import os
import re
import shutil
import zipfile
import hashlib
import time
import datetime
import logging
import subprocess
from django.conf import settings
from django.db import connections, connection
from django.utils import timezone
from django.contrib.auth import get_user_model
from backup.models import BackupSetting, BackupHistory

logger = logging.getLogger(__name__)
User = get_user_model()

# Matches rclone remote syntax like "gdrive:ERP_Backups/CloudVM" — requires 2+
# letters before the colon so a Windows drive letter ("C:", "G:") never
# matches this and keeps going through the plain local-path branch below.
RCLONE_REMOTE_RE = re.compile(r'^[A-Za-z][A-Za-z0-9_-]+:')

def is_rclone_remote(destination):
    """True if `destination` looks like an rclone remote (e.g. "gdrive:folder")
    rather than a local filesystem path (e.g. "C:\\folder" or "/mnt/folder")."""
    return bool(RCLONE_REMOTE_RE.match(destination)) and shutil.which('rclone') is not None

def get_share_link(remote_file_path):
    """Best-effort public/shareable link for a file already uploaded to an
    rclone remote (e.g. "onedrive:ERP_Backups/CloudVM_A1/backup.zip"). Returns
    None on any failure -- a missing link must never fail the backup itself,
    it just means the notification email won't have a clickable link for
    that destination."""
    try:
        result = subprocess.run(
            ['rclone', 'link', remote_file_path],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode == 0:
            return result.stdout.strip() or None
    except Exception as e:
        logger.warning(f"Could not generate share link for {remote_file_path}: {str(e)}")
    return None

def copy_to_cloud_folder(file_path, destination):
    """Copies file_path to destination, which is either a plain local/synced
    folder path (Windows OneDrive/GDrive desktop client folder — copied with
    shutil) or an rclone remote spec like "gdrive:ERP_Backups/CloudVM" (Linux
    servers with no desktop client — pushed with `rclone copy`). Returns the
    path/spec recorded for this copy, for backup_location bookkeeping.
    """
    if is_rclone_remote(destination):
        result = subprocess.run(
            ['rclone', 'copy', file_path, destination],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            raise Exception(f"rclone copy to {destination} failed: {result.stderr}")
        return f"{destination}/{os.path.basename(file_path)}"
    else:
        os.makedirs(destination, exist_ok=True)
        dest_path = os.path.join(destination, os.path.basename(file_path))
        shutil.copy2(file_path, dest_path)
        return dest_path

BACKUP_ZIP_RE = re.compile(r'^ERP_Backup_.*\.(zip|7z)$')
MIN_BACKUPS_TO_KEEP = 3


def _list_cloud_backups(destination):
    """[(name, modified datetime)] for the ERP backup zips in a cloud folder, or None if it
    could not be read (callers must then delete nothing)."""
    import json

    found = []
    if is_rclone_remote(destination):
        result = subprocess.run(
            ['rclone', 'lsjson', destination, '--files-only'], capture_output=True, text=True, timeout=300,
        )
        if result.returncode != 0:
            logger.error('rclone lsjson %s failed: %s', destination, result.stderr[-300:])
            return None
        try:
            entries = json.loads(result.stdout or '[]')
        except ValueError:
            logger.error('rclone lsjson %s returned unreadable output', destination)
            return None
        for entry in entries:
            if BACKUP_ZIP_RE.match(entry.get('Name', '')):
                modified = datetime.datetime.fromisoformat(entry['ModTime'].replace('Z', '+00:00'))
                found.append((entry['Name'], modified))
        return found
    if not os.path.isdir(destination):
        logger.error('Backup folder %s does not exist; nothing pruned', destination)
        return None
    for name in os.listdir(destination):
        full = os.path.join(destination, name)
        if BACKUP_ZIP_RE.match(name) and os.path.isfile(full):
            found.append((name, datetime.datetime.fromtimestamp(os.path.getmtime(full), tz=datetime.timezone.utc)))
    return found


def prune_cloud_backups(destination, retention_days, min_keep=MIN_BACKUPS_TO_KEEP, now=None, dry_run=False):
    """Delete ERP backup zips older than ``retention_days`` from one cloud folder.

    Safe by construction: only files named ``ERP_Backup_*.zip`` / ``.7z`` are touched, the newest
    ``min_keep`` are always kept (so a stalled backup job can never empty the folder),
    and if the folder cannot be listed nothing is deleted. Works for an rclone remote
    (``onedrive:ERP_Backups/CloudVM``) and for a desktop-synced folder path. Returns
    the names removed (or that would be removed, for ``dry_run``); never raises.
    """
    if not destination or not retention_days or retention_days <= 0:
        return []
    try:
        files = _list_cloud_backups(destination)
        if files is None:
            return []
        now = now or timezone.now()
        cutoff = now - datetime.timedelta(days=retention_days)
        newest_first = sorted(files, key=lambda item: item[1], reverse=True)
        protected = {name for name, _ in newest_first[:max(min_keep, 0)]}
        doomed = [name for name, modified in newest_first if modified < cutoff and name not in protected]
        if dry_run:
            return doomed
        removed = []
        for name in doomed:
            try:
                if is_rclone_remote(destination):
                    result = subprocess.run(
                        ['rclone', 'deletefile', f"{destination.rstrip('/')}/{name}"],
                        capture_output=True, text=True, timeout=300,
                    )
                    if result.returncode != 0:
                        raise Exception(result.stderr[-300:])
                else:
                    os.remove(os.path.join(destination, name))
                removed.append(name)
                logger.info('Pruned old backup %s from %s', name, destination)
            except Exception as exc:
                logger.error('Could not delete old backup %s from %s: %s', name, destination, exc)
        return removed
    except Exception:
        logger.exception('Backup pruning failed for %s', destination)
        return []


MEDIA_MIRROR_FOLDER = 'media_files'


def sync_media_incrementally(media_root, destination):
    """Mirror the media folder into ``<destination>/media_files`` - only files that are new
    or changed are transferred, and nothing already in the backup is ever deleted.

    Uploaded files (PO/WO PDFs, quotes, photos) never change once saved, so re-zipping and
    re-uploading the whole folder every day (about 36 MB a day, ~1 GB a month) was pure waste.
    Works for an rclone remote and for a desktop-synced folder path. Returns a small summary.
    """
    target = f"{destination.rstrip('/').rstrip(chr(92))}/{MEDIA_MIRROR_FOLDER}" if is_rclone_remote(destination) \
        else os.path.join(destination, MEDIA_MIRROR_FOLDER)
    if is_rclone_remote(destination):
        result = subprocess.run(
            ['rclone', 'copy', media_root, target, '--transfers', '4', '--stats', '0', '-v'],
            capture_output=True, text=True, timeout=3600,
        )
        if result.returncode != 0:
            raise Exception(f'rclone media copy to {target} failed: {result.stderr[-300:]}')
        transferred = sum(1 for line in (result.stderr or '').splitlines() if ': Copied (new)' in line or ': Copied (replaced' in line)
        return {'path': target, 'files_transferred': transferred}

    os.makedirs(target, exist_ok=True)
    transferred = 0
    for folder, _dirs, files in os.walk(media_root):
        relative = os.path.relpath(folder, media_root)
        out_dir = target if relative == '.' else os.path.join(target, relative)
        os.makedirs(out_dir, exist_ok=True)
        for name in files:
            source = os.path.join(folder, name)
            copy = os.path.join(out_dir, name)
            if os.path.exists(copy) and os.path.getsize(copy) == os.path.getsize(source) \
                    and int(os.path.getmtime(copy)) >= int(os.path.getmtime(source)):
                continue
            shutil.copy2(source, copy)
            transferred += 1
    return {'path': target, 'files_transferred': transferred}


def calculate_sha256(file_path):
    """Calculate the SHA-256 checksum of a file."""
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def get_database_engine():
    """Returns the type of database backend currently configured."""
    engine = settings.DATABASES['default']['ENGINE']
    if 'sqlite' in engine:
        return 'sqlite'
    elif 'postgresql' in engine:
        return 'postgresql'
    return 'unknown'

def run_sqlite_backup(dest_db_path):
    """Snapshot the SQLite database into a compact file.

    ``VACUUM INTO`` writes a consistent copy without the free pages and fragmentation a plain
    page-by-page copy carries over (and without touching the live database); if it is not
    available, fall back to the native backup API."""
    import sqlite3
    try:
        if os.path.exists(dest_db_path):
            os.remove(dest_db_path)
        with connections['default'].cursor() as cursor:
            cursor.execute("VACUUM INTO %s", [dest_db_path])
        return
    except Exception as exc:
        logger.warning("VACUUM INTO failed (%s); using the standard backup API instead", exc)
        if os.path.exists(dest_db_path):
            os.remove(dest_db_path)
    # Use connections['default'].cursor().connection to get the active sqlite3 connection
    src_conn = connections['default'].cursor().connection
    dest_conn = sqlite3.connect(dest_db_path)
    with dest_conn:
        src_conn.backup(dest_conn)
    dest_conn.close()

def run_postgresql_backup(dest_db_path):
    """Backup PostgreSQL database using pg_dump."""
    db_config = settings.DATABASES['default']
    db_name = db_config['NAME']
    db_user = db_config.get('USER', '')
    db_password = db_config.get('PASSWORD', '')
    db_host = db_config.get('HOST', 'localhost')
    db_port = db_config.get('PORT', '5432')

    env = os.environ.copy()
    if db_password:
        env['PGPASSWORD'] = db_password

    cmd = [
        'pg_dump',
        '-h', db_host,
        '-p', str(db_port),
        '-U', db_user,
        '-F', 'c', # Custom archive format
        '-b',      # Include large objects
        '-v',      # Verbose
        '-f', dest_db_path,
        db_name
    ]

    result = subprocess.run(cmd, env=env, capture_output=True, text=True)
    if result.returncode != 0:
        raise Exception(f"pg_dump failed: {result.stderr}")

def _setting(env_name, db_value):
    """DB-stored backup settings live in db.sqlite3, which gets wholesale
    overwritten every time a production->cloud data sync runs (see
    scripts/cloud/sync_db_to_cloud.bat) -- silently reverting a cloud
    deployment's backup routing back to whatever the syncing machine had
    configured (e.g. Windows paths that don't exist on Linux). An env var
    set in this deployment's own .env, which a data sync can't touch,
    always wins over the DB value when present.
    """
    return os.environ.get(env_name) or db_value

class _NullArchive:
    """Stands in for the zip writer when the archive was already produced by 7-Zip."""

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def write(self, *args, **kwargs):
        pass


def _open_zip(path, skip):
    return _NullArchive() if skip else zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)


def _int_setting(env_name, db_value):
    """Like _setting, for integers: an env var in this deployment's own .env wins over the DB value."""
    raw = os.environ.get(env_name)
    if raw not in (None, ''):
        try:
            return int(raw)
        except ValueError:
            logger.warning('%s=%r is not a number; using the stored setting', env_name, raw)
    return int(db_value or 0)


def create_backup(backup_type='AUTO', user=None, targets=None):
    """
    Creates a database backup, archives it, syncs to cloud folders, and enforces retention.
    Returns the created BackupHistory instance.

    `targets`: None (default) backs up to both OneDrive and Google Drive, as
    always. Pass {'onedrive'} or {'gdrive'} to run a backup that only syncs to
    that one destination -- used when the two clouds run on separate
    schedules (BackupSetting.onedrive_backup_time / gdrive_backup_time).
    """
    settings_obj = BackupSetting.get_settings()
    targets_label = 'both' if targets is None else ','.join(sorted(targets)) or 'both'

    # Create a new history entry
    history = BackupHistory.objects.create(
        backup_type=backup_type,
        targets=targets_label,
        start_time=timezone.now(),
        status='PENDING',
        created_by=user
    )

    temp_files = []

    # Trim short-lived data first so this backup (and every later one) stays small. A cleanup
    # problem must never stop a backup.
    try:
        from core.data_retention import prune_old_data
        prune_old_data()
    except Exception:
        logger.exception("Pre-backup data cleanup failed; continuing with the backup")

    try:
        # Create local backup folder if it doesn't exist. Kept inside the try so a
        # bad/missing backup path records a FAILED history instead of orphaning the
        # record as PENDING.
        local_dir = _setting('BACKUP_LOCAL_FOLDER_OVERRIDE', settings_obj.local_backup_folder)
        if not os.path.isabs(local_dir):
            local_dir = os.path.join(settings.BASE_DIR, local_dir)
        os.makedirs(local_dir, exist_ok=True)

        engine = get_database_engine()
        timestamp = timezone.localtime(timezone.now()).strftime("%Y-%m-%d_%H%M%S")
        erp_version = getattr(settings, 'ERP_SOFTWARE_VERSION', '1.0')
        
        # Temp database copy file
        temp_db_filename = f"db_temp_{timestamp}.bak"
        temp_db_path = os.path.join(local_dir, temp_db_filename)
        temp_files.append(temp_db_path)
        
        logger.info(f"Starting database backup for {engine}")
        
        if engine == 'sqlite':
            run_sqlite_backup(temp_db_path)
        elif engine == 'postgresql':
            run_postgresql_backup(temp_db_path)
        else:
            raise Exception(f"Unsupported database engine for automatic backups: {engine}")
            
        # Target ZIP name. Includes an instance label (e.g. "CloudVM") when set,
        # so backups from multiple servers syncing to the same Drive/OneDrive
        # account don't look identical.
        instance_label = getattr(settings, 'BACKUP_INSTANCE_LABEL', '')
        label_part = f"_{instance_label}" if instance_label else ""
        zip_filename = f"ERP_Backup_{engine}_v{erp_version}{label_part}_{timestamp}.zip"
        zip_filepath = os.path.join(local_dir, zip_filename)
        
        # Effective values: env override (immune to db-sync overwrites) wins
        # over the DB-stored setting when present. When `targets` restricts
        # this run to one destination (split-schedule mode), the other
        # destination is treated as unset for this run only.
        effective_onedrive = _setting('BACKUP_ONEDRIVE_FOLDER_OVERRIDE', settings_obj.cloud_onedrive_folder)
        effective_gdrive = _setting('BACKUP_GDRIVE_FOLDER_OVERRIDE', settings_obj.cloud_gdrive_folder)
        if targets is not None:
            if 'onedrive' not in targets:
                effective_onedrive = None
            if 'gdrive' not in targets:
                effective_gdrive = None
        effective_media_folder = _setting('BACKUP_MEDIA_FOLDER_OVERRIDE', settings_obj.media_cloud_folder)
        # Remember whether media has its own destination BEFORE the per-run adjustment below
        # blanks it for a run that doesn't cover that destination. Without this, the OneDrive
        # run (media is routed to Google Drive) lost track of that and bundled the whole media
        # folder into its database zip - roughly 35 MB of PDFs a day on a nearly-full OneDrive.
        media_has_own_destination = bool(effective_media_folder)
        # Media rides along with whichever run actually covers its own
        # destination remote -- in split mode, a gdrive: media folder only
        # backs up during the gdrive run, not the onedrive one (and vice
        # versa), so it isn't silently skipped or duplicated across both.
        if targets is not None and effective_media_folder:
            media_target = (
                'onedrive' if effective_media_folder.startswith('onedrive:')
                else 'gdrive' if effective_media_folder.startswith('gdrive:')
                else None
            )
            if media_target and media_target not in targets:
                effective_media_folder = None
        effective_include_media = os.environ.get('BACKUP_INCLUDE_MEDIA_OVERRIDE')
        effective_include_media = (
            effective_include_media == 'True' if effective_include_media is not None
            else settings_obj.include_media
        )

        # If a separate media destination is configured, media goes into its own
        # archive routed only there, instead of being bundled into the db zip
        # that's sent to both OneDrive/Google Drive folders below.
        separate_media = bool(effective_include_media and effective_media_folder)

        # Compress database and optional files. A database-only backup uses 7-Zip (LZMA2, less
        # than half the size) when it is installed and enabled; anything bundling media or logs,
        # or any 7-Zip problem, falls back to the ordinary deflate .zip below.
        bundles_extras = bool(
            (effective_include_media and not media_has_own_destination and hasattr(settings, 'MEDIA_ROOT') and settings.MEDIA_ROOT)
            or settings_obj.include_logs
        )
        wanted_format = os.environ.get('BACKUP_COMPRESSION_FORMAT') or settings_obj.compression_format
        made_7z = False
        if wanted_format == '7z' and not bundles_extras and engine == 'sqlite':
            from backup.archive import ArchiveError, create_7z
            archive_path_7z = os.path.splitext(zip_filepath)[0] + '.7z'
            try:
                create_7z(archive_path_7z, temp_db_path, 'db.sqlite3')
                zip_filepath, zip_filename, made_7z = archive_path_7z, os.path.basename(archive_path_7z), True
                logger.info("Compressed database backup with 7-Zip: %s", zip_filename)
            except ArchiveError as exc:
                logger.warning("7-Zip backup not used (%s); making a .zip instead", exc)
        logger.info("Compressing backup files into %s archive", '7z' if made_7z else 'ZIP')
        with _open_zip(zip_filepath, made_7z) as zip_file:
            # Add database copy to the zip as database file
            db_in_zip_name = "db.sqlite3" if engine == 'sqlite' else "db.dump"
            zip_file.write(temp_db_path, arcname=db_in_zip_name)

            # Optionally include media (bundled in the same zip, unless routed
            # separately via media_cloud_folder above)
            if effective_include_media and not media_has_own_destination and hasattr(settings, 'MEDIA_ROOT') and settings.MEDIA_ROOT:
                media_root = settings.MEDIA_ROOT
                if os.path.exists(media_root):
                    for root, dirs, files in os.walk(media_root):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, os.path.dirname(media_root))
                            zip_file.write(file_path, arcname=arcname)

            # Optionally include logs
            if settings_obj.include_logs:
                # We search for .log files in BASE_DIR or settings.LOGGING log files if configured
                for root, dirs, files in os.walk(settings.BASE_DIR):
                    # Only search top-level logs to avoid scanning virtualenv
                    if root == str(settings.BASE_DIR):
                        for file in files:
                            if file.endswith('.log'):
                                file_path = os.path.join(root, file)
                                zip_file.write(file_path, arcname=os.path.join("logs", file))
                                
        # Check if ZIP password protection/encryption is requested
        # python's zipfile doesn't support writing password protected files out-of-the-box
        # without external libraries like pyminizip. If encryption is enabled, we'll log it
        # or implement a lightweight warning/information.
        if settings_obj.enable_encryption and settings_obj.encryption_password:
            logger.info("ZIP encryption requested. Since Python standard zipfile lacks writing native encryption, "
                        "this will require a future extension or pyminizip. Storing ZIP normally but logging password protection requirement.")
            
        # Calculate checksum and size
        file_size = os.path.getsize(zip_filepath)
        checksum = calculate_sha256(zip_filepath)
        
        # Copy to cloud synchronization folders if defined. A failure here
        # (bad path, unconfigured rclone remote, etc.) must NOT be silently
        # swallowed — a backup that never leaves this VM defeats the whole
        # point of an off-site copy, so it has to show up as a real failure
        # on the dashboard, not a quiet log line under a green "SUCCESS".
        locations = [zip_filepath]
        cloud_errors = []
        cloud_sync_details = {}

        if effective_onedrive:
            try:
                od_path = copy_to_cloud_folder(zip_filepath, effective_onedrive)
                locations.append(od_path)
                logger.info(f"Successfully copied backup to OneDrive: {od_path}")
                cloud_sync_details['onedrive'] = {
                    'path': od_path,
                    'link': get_share_link(od_path) if is_rclone_remote(effective_onedrive) else None,
                    'synced_at': timezone.now().isoformat(),
                }
                # Keep OneDrive to a rolling window (default one week). Runs only after a
                # successful upload, and never fails the backup itself.
                retention_days = _int_setting('BACKUP_ONEDRIVE_RETENTION_DAYS_OVERRIDE', settings_obj.onedrive_retention_days)
                pruned = prune_cloud_backups(effective_onedrive, retention_days)
                if pruned:
                    cloud_sync_details['onedrive']['pruned'] = pruned
            except Exception as e:
                logger.error(f"Failed to copy to OneDrive folder: {str(e)}")
                cloud_errors.append(f"OneDrive: {str(e)}")

        if effective_gdrive:
            try:
                gd_path = copy_to_cloud_folder(zip_filepath, effective_gdrive)
                locations.append(gd_path)
                logger.info(f"Successfully copied backup to Google Drive: {gd_path}")
                cloud_sync_details['gdrive'] = {
                    'path': gd_path,
                    'link': get_share_link(gd_path) if is_rclone_remote(effective_gdrive) else None,
                    'synced_at': timezone.now().isoformat(),
                }
            except Exception as e:
                logger.error(f"Failed to copy to Google Drive folder: {str(e)}")
                cloud_errors.append(f"Google Drive: {str(e)}")

        # Media routed to its own destination: mirrored incrementally (new/changed files
        # only) rather than re-zipped and re-uploaded in full on every run.
        if separate_media and hasattr(settings, 'MEDIA_ROOT') and settings.MEDIA_ROOT and os.path.exists(settings.MEDIA_ROOT):
            try:
                media_summary = sync_media_incrementally(settings.MEDIA_ROOT, effective_media_folder)
                cloud_sync_details['media'] = {**media_summary, 'synced_at': timezone.now().isoformat()}
                logger.info(
                    "Media mirrored to %s (%s new/changed file(s))",
                    media_summary['path'], media_summary['files_transferred'],
                )
            except Exception as e:
                logger.error(f"Failed to sync media backup: {str(e)}")
                cloud_errors.append(f"Media: {str(e)}")

        # Finalize history record
        finish_time = timezone.now()
        duration = int((finish_time - history.start_time).total_seconds())

        history.finish_time = finish_time
        history.duration_seconds = max(1, duration)
        history.file_name = zip_filename
        history.file_size = file_size
        history.backup_location = ", ".join(locations)
        history.cloud_sync_details = cloud_sync_details
        if cloud_errors:
            history.status = 'FAILED'
            history.error_message = (
                'Local backup succeeded, but cloud sync failed — '
                + '; '.join(cloud_errors)
            )
        else:
            history.status = 'SUCCESS'
        history.sha256_checksum = checksum
        history.save()

        logger.info(f"Backup created successfully: {zip_filename}")

        # Run retention cleanup
        run_retention_cleanup(settings_obj)

        send_backup_notification(history, settings_obj)

        return history

    except Exception as e:
        logger.exception("Backup failed due to exception:")
        history.finish_time = timezone.now()
        history.status = 'FAILED'
        history.error_message = str(e)
        history.save()

        send_backup_notification(history, settings_obj)

        return history
        
    finally:
        # Clean up temp backup files
        for f in temp_files:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception as ex:
                    logger.error(f"Could not clean up temp file {f}: {str(ex)}")

def send_backup_notification(history, settings_obj):
    """Emails the outcome of one backup run -- success or failure -- with a
    separate link and sync time for OneDrive and Google Drive, so a real gap
    between the two (a stuck/slow remote, one destination failing silently in
    the past, etc.) is visible at a glance instead of buried in the dashboard.
    Never raises -- a notification failure must not turn a good backup into a
    reported failure."""
    if not settings_obj.enable_notifications:
        return
    recipients = [addr.strip() for addr in (settings_obj.notify_email or '').replace(';', ',').split(',') if addr.strip()]
    if not recipients:
        return

    try:
        from django.core.mail import send_mail
        from django.template.loader import render_to_string

        details = history.cloud_sync_details or {}
        onedrive = details.get('onedrive')
        gdrive = details.get('gdrive')

        time_gap_seconds = None
        if onedrive and gdrive and onedrive.get('synced_at') and gdrive.get('synced_at'):
            try:
                t1 = datetime.datetime.fromisoformat(onedrive['synced_at'])
                t2 = datetime.datetime.fromisoformat(gdrive['synced_at'])
                time_gap_seconds = abs((t2 - t1).total_seconds())
            except Exception:
                pass

        instance_label = getattr(settings, 'BACKUP_INSTANCE_LABEL', '') or 'ERP'
        scope_label = {
            'onedrive': ' (OneDrive only)', 'gdrive': ' (Google Drive only)',
        }.get(history.targets, '')
        subject = f"[{history.status}] {instance_label} Backup{scope_label} - {history.file_name or history.start_time.strftime('%Y-%m-%d %H:%M')}"

        body = render_to_string('backup/email_notification.html', {
            'history': history,
            'instance_label': instance_label,
            'scope_label': scope_label,
            'onedrive': onedrive,
            'gdrive': gdrive,
            'time_gap_seconds': time_gap_seconds,
        })

        send_mail(
            subject=subject,
            message=body,
            from_email=None,
            recipient_list=recipients,
            html_message=body,
            fail_silently=True,
        )
        history.notification_sent = True
        history.save(update_fields=['notification_sent'])
    except Exception:
        logger.exception("Failed to send backup notification email")


def run_retention_cleanup(settings_obj):
    """Deletes old backups according to the retention settings."""
    logger.info("Starting backup retention policy cleanup")
    
    # We retrieve all successful backups
    backups = BackupHistory.objects.filter(status='SUCCESS').order_name_desc() if hasattr(BackupHistory.objects, 'order_name_desc') else BackupHistory.objects.filter(status='SUCCESS').order_by('-start_time')
    
    daily_count = 0
    weekly_count = 0
    monthly_count = 0
    
    for backup in backups:
        # Determine schedule classification
        # For simplicity, we keep the last N backups based on settings
        # Keep last keep_daily daily (auto/manual), keep_weekly weekly, keep_monthly monthly
        # Since backups have a timestamp, we can group them
        # Let's count them:
        if backup.backup_type == 'AUTO':
            # We can classify them or keep total limits
            pass
            
    # Simple limit enforcement: Keep last N total backups for safety or clean up files that exceed limits
    # Daily limits
    auto_backups = BackupHistory.objects.filter(status='SUCCESS', backup_type='AUTO').order_by('-start_time')
    if auto_backups.count() > settings_obj.keep_daily:
        to_delete = auto_backups[settings_obj.keep_daily:]
        for backup in to_delete:
            delete_backup_files(backup)
            
    manual_backups = BackupHistory.objects.filter(status='SUCCESS', backup_type='MANUAL').order_by('-start_time')
    # Keep last 15 manual backups
    if manual_backups.count() > 15:
        to_delete = manual_backups[15:]
        for backup in to_delete:
            delete_backup_files(backup)

def delete_backup_files(backup_record):
    """Deletes physical backup files associated with a BackupHistory record."""
    if backup_record.backup_location:
        paths = [p.strip() for p in backup_record.backup_location.split(",")]
        for p in paths:
            if os.path.exists(p):
                try:
                    os.remove(p)
                    logger.info(f"Deleted old backup file: {p}")
                except Exception as e:
                    logger.error(f"Failed to delete backup file {p}: {str(e)}")
    backup_record.delete()
