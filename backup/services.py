import os
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
    """Safely backup SQLite database using the native Python/sqlite3 backup API."""
    import sqlite3
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

def create_backup(backup_type='AUTO', user=None):
    """
    Creates a database backup, archives it, syncs to cloud folders, and enforces retention.
    Returns the created BackupHistory instance.
    """
    settings_obj = BackupSetting.get_settings()
    
    # Create a new history entry
    history = BackupHistory.objects.create(
        backup_type=backup_type,
        start_time=timezone.now(),
        status='PENDING',
        created_by=user
    )
    
    temp_files = []

    try:
        # Create local backup folder if it doesn't exist. Kept inside the try so a
        # bad/missing backup path records a FAILED history instead of orphaning the
        # record as PENDING.
        local_dir = settings_obj.local_backup_folder
        if not os.path.isabs(local_dir):
            local_dir = os.path.join(settings.BASE_DIR, local_dir)
        os.makedirs(local_dir, exist_ok=True)

        engine = get_database_engine()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
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
            
        # Target ZIP name
        zip_filename = f"ERP_Backup_{engine}_v{erp_version}_{timestamp}.zip"
        zip_filepath = os.path.join(local_dir, zip_filename)
        
        # Compress database and optional files
        logger.info("Compressing backup files into ZIP archive")
        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Add database copy to the zip as database file
            db_in_zip_name = "db.sqlite3" if engine == 'sqlite' else "db.dump"
            zip_file.write(temp_db_path, arcname=db_in_zip_name)
            
            # Optionally include media
            if settings_obj.include_media and hasattr(settings, 'MEDIA_ROOT') and settings.MEDIA_ROOT:
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
        
        # Copy to cloud synchronization folders if defined
        locations = [zip_filepath]
        
        if settings_obj.cloud_onedrive_folder:
            od_dir = settings_obj.cloud_onedrive_folder
            try:
                os.makedirs(od_dir, exist_ok=True)
                od_path = os.path.join(od_dir, zip_filename)
                shutil.copy2(zip_filepath, od_path)
                locations.append(od_path)
                logger.info(f"Successfully copied backup to OneDrive: {od_path}")
            except Exception as e:
                logger.error(f"Failed to copy to OneDrive folder: {str(e)}")
                
        if settings_obj.cloud_gdrive_folder:
            gd_dir = settings_obj.cloud_gdrive_folder
            try:
                os.makedirs(gd_dir, exist_ok=True)
                gd_path = os.path.join(gd_dir, zip_filename)
                shutil.copy2(zip_filepath, gd_path)
                locations.append(gd_path)
                logger.info(f"Successfully copied backup to Google Drive: {gd_path}")
            except Exception as e:
                logger.error(f"Failed to copy to Google Drive folder: {str(e)}")
                
        # Finalize history record
        finish_time = timezone.now()
        duration = int((finish_time - history.start_time).total_seconds())
        
        history.finish_time = finish_time
        history.duration_seconds = max(1, duration)
        history.file_name = zip_filename
        history.file_size = file_size
        history.backup_location = ", ".join(locations)
        history.status = 'SUCCESS'
        history.sha256_checksum = checksum
        history.save()
        
        logger.info(f"Backup created successfully: {zip_filename}")
        
        # Run retention cleanup
        run_retention_cleanup(settings_obj)
        
        return history
        
    except Exception as e:
        logger.exception("Backup failed due to exception:")
        history.finish_time = timezone.now()
        history.status = 'FAILED'
        history.error_message = str(e)
        history.save()
        
        # Log to notification or system logs
        return history
        
    finally:
        # Clean up temp backup files
        for f in temp_files:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception as ex:
                    logger.error(f"Could not clean up temp file {f}: {str(ex)}")

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
