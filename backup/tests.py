import os
import zipfile
from django.test import TransactionTestCase as TestCase
from django.conf import settings
from django.utils import timezone
from django.contrib.auth import get_user_model
from backup.models import BackupSetting, BackupHistory, RestoreHistory
from backup.services import create_backup, get_database_engine, run_retention_cleanup
from backup.restore import simulate_restore, perform_restore

User = get_user_model()

class BackupSystemTests(TestCase):
    def setUp(self):
        self.user = User.objects.create_superuser(username='admin_test', password='password123', email='admin@test.com')
        # Setup settings
        self.setting = BackupSetting.get_settings()
        self.setting.local_backup_folder = 'test_backups'
        self.setting.compression_format = 'zip'      # these tests open the archive with zipfile
        self.setting.save()

    def tearDown(self):
        # Clean up any generated test backups
        test_dir = os.path.join(settings.BASE_DIR, 'test_backups')
        if os.path.exists(test_dir):
            import shutil
            shutil.rmtree(test_dir)
            
        test_restore_dir = os.path.join(settings.BASE_DIR, 'backups', 'temp_restore')
        if os.path.exists(test_restore_dir):
            shutil.rmtree(test_restore_dir)

    def test_settings_initialization(self):
        """Test default settings retrieve or create behavior."""
        setting = BackupSetting.get_settings()
        self.assertTrue(setting.backup_enabled)
        self.assertEqual(setting.frequency, 'DAILY')

    def test_create_backup_sqlite(self):
        """Test backup creation process."""
        if get_database_engine() != 'sqlite':
            self.skipTest("Skipping SQLite-specific backup tests because another database engine is configured.")
            
        history = create_backup(backup_type='MANUAL', user=self.user)
        
        self.assertEqual(history.status, 'SUCCESS')
        self.assertEqual(history.backup_type, 'MANUAL')
        self.assertEqual(history.created_by, self.user)
        self.assertTrue(history.file_name.startswith('ERP_Backup_sqlite_'))
        self.assertTrue(history.file_name.endswith('.zip'))
        self.assertIsNotNone(history.sha256_checksum)
        self.assertTrue(history.file_size > 0)
        
        # Verify physical file existence
        file_path = os.path.join(settings.BASE_DIR, 'test_backups', history.file_name)
        self.assertTrue(os.path.exists(file_path))
        
        # Verify zip contents
        with zipfile.ZipFile(file_path, 'r') as zf:
            self.assertIn('db.sqlite3', zf.namelist())

    def test_restore_simulation_and_execution(self):
        """Test dry-run simulation and safe SQLite database restore."""
        if get_database_engine() != 'sqlite':
            self.skipTest("Skipping SQLite-specific restore tests.")
            
        # Create a backup
        history = create_backup(backup_type='MANUAL', user=self.user)
        self.assertEqual(history.status, 'SUCCESS')
        
        # Verify simulated restore
        sim = simulate_restore(history.id)
        self.assertTrue(sim.is_valid)
        self.assertTrue(sim.db_size > 0)
        self.assertEqual(sim.file_count, 1)
        
        # Run restore
        success = perform_restore(history.id, user=self.user)
        self.assertTrue(success)
        
        # Check restore history entry
        restore_entry = RestoreHistory.objects.filter(backup=history).first()
        self.assertIsNotNone(restore_entry)
        self.assertEqual(restore_entry.status, 'SUCCESS')
        self.assertEqual(restore_entry.executed_by, self.user)

    def test_retention_cleanup(self):
        """Test that retention policies successfully purge older backup records/files."""
        self.setting.keep_daily = 3
        self.setting.save()
        
        # Create 5 mock backup records
        for i in range(5):
            h = BackupHistory.objects.create(
                backup_type='AUTO',
                status='SUCCESS',
                start_time=timezone.now() - timezone.timedelta(days=i),
                file_name=f"test_backup_{i}.zip",
                backup_location=os.path.join(settings.BASE_DIR, 'test_backups', f"test_backup_{i}.zip")
            )
            # Create dummy physical files so deletion code is hit without throwing
            local_dir = os.path.join(settings.BASE_DIR, 'test_backups')
            os.makedirs(local_dir, exist_ok=True)
            with open(h.backup_location, 'w') as f:
                f.write("dummy db contents")
                
        # Trigger cleanup
        run_retention_cleanup(self.setting)
        
        # Verify that only 3 automatic backups remain
        remaining = BackupHistory.objects.filter(backup_type='AUTO', status='SUCCESS')
        self.assertEqual(remaining.count(), 3)


class OneDriveRetentionTests(TestCase):
    """OneDrive keeps a rolling window (default one week); older ERP backup zips are removed automatically."""

    def setUp(self):
        import shutil
        import tempfile

        self.folder = tempfile.mkdtemp(prefix='onedrive_test_')
        self.addCleanup(shutil.rmtree, self.folder, True)
        self.now = timezone.now()

    def make_zip(self, days_old, label='CloudVM_A1'):
        import datetime

        name = f'ERP_Backup_sqlite_v1_{label}_{days_old:02d}d.zip'
        path = os.path.join(self.folder, name)
        with open(path, 'wb') as handle:
            handle.write(b'x')
        stamp = (self.now - datetime.timedelta(days=days_old, hours=1)).timestamp()
        os.utime(path, (stamp, stamp))
        return name

    def remaining(self):
        return sorted(os.listdir(self.folder))

    def test_new_installations_default_to_one_week(self):
        self.assertEqual(BackupSetting.get_settings().onedrive_retention_days, 7)

    def test_files_older_than_the_window_are_deleted_and_recent_ones_kept(self):
        from backup.services import prune_cloud_backups

        names = {days: self.make_zip(days) for days in (14, 9, 8, 6, 3, 1)}
        removed = prune_cloud_backups(self.folder, 7, now=self.now)
        self.assertEqual(sorted(removed), sorted([names[14], names[9], names[8]]))
        self.assertEqual(self.remaining(), sorted([names[6], names[3], names[1]]))

    def test_the_newest_three_are_always_kept_even_if_everything_is_old(self):
        from backup.services import prune_cloud_backups

        names = [self.make_zip(days) for days in (30, 29, 28, 27, 26)]
        prune_cloud_backups(self.folder, 7, now=self.now)
        self.assertEqual(self.remaining(), sorted(names[2:]))        # 28, 27, 26 days old survive

    def test_only_erp_backup_zips_are_ever_touched(self):
        from backup.services import prune_cloud_backups

        for other in ('notes.txt', 'photo.zip', 'ERP_Backup_notes.txt'):
            with open(os.path.join(self.folder, other), 'w') as handle:
                handle.write('keep me')
        for days in (40, 39, 38, 37, 36):
            self.make_zip(days)
        prune_cloud_backups(self.folder, 7, now=self.now)
        for other in ('notes.txt', 'photo.zip', 'ERP_Backup_notes.txt'):
            self.assertIn(other, self.remaining())

    def test_zero_days_means_keep_everything(self):
        from backup.services import prune_cloud_backups

        for days in (40, 39, 38, 37, 36):
            self.make_zip(days)
        self.assertEqual(prune_cloud_backups(self.folder, 0, now=self.now), [])
        self.assertEqual(len(self.remaining()), 5)

    def test_dry_run_reports_without_deleting(self):
        from backup.services import prune_cloud_backups

        for days in (20, 19, 18, 2, 1):
            self.make_zip(days)
        would_remove = prune_cloud_backups(self.folder, 7, now=self.now, dry_run=True)
        self.assertEqual(len(would_remove), 2)            # 20d and 19d; the newest three (18d, 2d, 1d) stay
        self.assertEqual(len(self.remaining()), 5)

    def test_missing_or_unset_folder_deletes_nothing_and_does_not_raise(self):
        from backup.services import prune_cloud_backups

        self.assertEqual(prune_cloud_backups(os.path.join(self.folder, 'nope'), 7), [])
        self.assertEqual(prune_cloud_backups('', 7), [])
        self.assertEqual(prune_cloud_backups(None, 7), [])

    def test_a_stalled_backup_job_cannot_empty_the_folder(self):
        from backup.services import prune_cloud_backups

        for days in (60, 59, 58):
            self.make_zip(days)
        prune_cloud_backups(self.folder, 7, now=self.now)
        self.assertEqual(len(self.remaining()), 3)

    def test_rclone_remote_lists_then_deletes_only_expired_backups(self):
        import json
        import subprocess
        from unittest import mock

        from backup.services import prune_cloud_backups

        def entry(name, days_old):
            when = (self.now - __import__('datetime').timedelta(days=days_old)).strftime('%Y-%m-%dT%H:%M:%SZ')
            return {'Name': name, 'Path': name, 'ModTime': when, 'Size': 10, 'IsDir': False}

        listing = [
            entry('ERP_Backup_sqlite_v1_CloudVM_A1_old.zip', 20),
            entry('ERP_Backup_sqlite_v1_CloudVM_A1_mid.zip', 9),
            entry('ERP_Backup_sqlite_v1_CloudVM_A1_5d.zip', 5),
            entry('ERP_Backup_sqlite_v1_CloudVM_A1_3d.zip', 3),
            entry('ERP_Backup_sqlite_v1_CloudVM_A1_1d.zip', 1),
            entry('readme.txt', 90),
        ]
        calls = []

        def fake_run(cmd, **kwargs):
            calls.append(cmd)
            if cmd[1] == 'lsjson':
                return subprocess.CompletedProcess(cmd, 0, stdout=json.dumps(listing), stderr='')
            return subprocess.CompletedProcess(cmd, 0, stdout='', stderr='')

        with mock.patch('backup.services.is_rclone_remote', return_value=True), \
                mock.patch('backup.services.subprocess.run', side_effect=fake_run):
            removed = prune_cloud_backups('onedrive:ERP_Backups/CloudVM_A1', 7, now=self.now)

        self.assertEqual(sorted(removed), ['ERP_Backup_sqlite_v1_CloudVM_A1_mid.zip', 'ERP_Backup_sqlite_v1_CloudVM_A1_old.zip'])
        deletes = sorted(cmd[2] for cmd in calls if cmd[1] == 'deletefile')
        self.assertEqual(deletes, [
            'onedrive:ERP_Backups/CloudVM_A1/ERP_Backup_sqlite_v1_CloudVM_A1_mid.zip',
            'onedrive:ERP_Backups/CloudVM_A1/ERP_Backup_sqlite_v1_CloudVM_A1_old.zip',
        ])
        self.assertFalse(any('readme' in ' '.join(cmd) for cmd in calls if cmd[1] == 'deletefile'))

    def test_failed_rclone_listing_deletes_nothing(self):
        import subprocess
        from unittest import mock

        from backup.services import prune_cloud_backups

        def fake_run(cmd, **kwargs):
            return subprocess.CompletedProcess(cmd, 1, stdout='', stderr='auth expired')

        with mock.patch('backup.services.is_rclone_remote', return_value=True), \
                mock.patch('backup.services.subprocess.run', side_effect=fake_run) as run:
            self.assertEqual(prune_cloud_backups('onedrive:ERP_Backups/CloudVM_A1', 7), [])
        self.assertEqual([c.args[0][1] for c in run.call_args_list], ['lsjson'])     # never got as far as deleting

    def test_one_failed_delete_does_not_stop_the_rest(self):
        import json
        import subprocess
        from unittest import mock

        from backup.services import prune_cloud_backups

        import datetime

        def entry(name, days_old):
            when = (self.now - datetime.timedelta(days=days_old)).strftime('%Y-%m-%dT%H:%M:%SZ')
            return {'Name': name, 'Path': name, 'ModTime': when, 'Size': 10, 'IsDir': False}

        listing = [entry(f'ERP_Backup_x_{d}.zip', d) for d in (30, 20, 5, 4, 3)]

        def fake_run(cmd, **kwargs):
            if cmd[1] == 'lsjson':
                return subprocess.CompletedProcess(cmd, 0, stdout=json.dumps(listing), stderr='')
            if cmd[2].endswith('ERP_Backup_x_30.zip'):
                return subprocess.CompletedProcess(cmd, 1, stdout='', stderr='locked')
            return subprocess.CompletedProcess(cmd, 0, stdout='', stderr='')

        with mock.patch('backup.services.is_rclone_remote', return_value=True), \
                mock.patch('backup.services.subprocess.run', side_effect=fake_run):
            removed = prune_cloud_backups('onedrive:b', 7, now=self.now)
        self.assertEqual(removed, ['ERP_Backup_x_20.zip'])

    def test_env_override_wins_over_the_stored_value(self):
        from unittest import mock

        from backup.services import _int_setting

        self.assertEqual(_int_setting('BACKUP_ONEDRIVE_RETENTION_DAYS_OVERRIDE', 7), 7)
        with mock.patch.dict(os.environ, {'BACKUP_ONEDRIVE_RETENTION_DAYS_OVERRIDE': '10'}):
            self.assertEqual(_int_setting('BACKUP_ONEDRIVE_RETENTION_DAYS_OVERRIDE', 7), 10)
        with mock.patch.dict(os.environ, {'BACKUP_ONEDRIVE_RETENTION_DAYS_OVERRIDE': 'abc'}):
            self.assertEqual(_int_setting('BACKUP_ONEDRIVE_RETENTION_DAYS_OVERRIDE', 7), 7)

    def test_a_successful_onedrive_upload_triggers_pruning_and_a_failed_one_does_not(self):
        from unittest import mock

        from backup.models import BackupSetting
        from backup.services import create_backup, get_database_engine

        if get_database_engine() != 'sqlite':
            self.skipTest('SQLite-specific backup test')
        settings_obj = BackupSetting.get_settings()
        settings_obj.local_backup_folder = 'test_backups_prune'
        settings_obj.cloud_onedrive_folder = self.folder
        settings_obj.onedrive_retention_days = 7
        settings_obj.save()
        for days in (30, 29, 28, 27):
            self.make_zip(days)
        try:
            create_backup(backup_type='MANUAL', targets={'onedrive'})
            left = self.remaining()
            self.assertEqual(len(left), 3)                                   # new upload + the 2 newest old ones
            self.assertFalse(any('_30d' in n or '_29d' in n for n in left))   # the oldest were pruned

            with mock.patch('backup.services.copy_to_cloud_folder', side_effect=Exception('OneDrive down')), \
                    mock.patch('backup.services.prune_cloud_backups') as prune:
                create_backup(backup_type='MANUAL', targets={'onedrive'})
            prune.assert_not_called()
        finally:
            import shutil

            shutil.rmtree(os.path.join(settings.BASE_DIR, 'test_backups_prune'), ignore_errors=True)


class MediaRoutingTests(TestCase):
    """Media routed to its own destination must never be bundled into the database zip of another run."""

    def setUp(self):
        import shutil
        import tempfile
        from unittest import mock

        from backup.models import BackupSetting
        from backup.services import get_database_engine

        if get_database_engine() != 'sqlite':
            self.skipTest('SQLite-specific backup test')
        self.media_root = tempfile.mkdtemp(prefix='media_root_')
        os.makedirs(os.path.join(self.media_root, 'planning'))
        with open(os.path.join(self.media_root, 'planning', 'po.pdf'), 'wb') as handle:
            handle.write(b'%PDF-1.4 test')
        self.onedrive = tempfile.mkdtemp(prefix='od_')
        for path in (self.media_root, self.onedrive):
            self.addCleanup(shutil.rmtree, path, True)
        settings_obj = BackupSetting.get_settings()
        settings_obj.local_backup_folder = 'test_backups_media'
        settings_obj.cloud_onedrive_folder = self.onedrive
        settings_obj.cloud_gdrive_folder = 'gdrive:ERP_Backups/Test'
        settings_obj.media_cloud_folder = 'gdrive:ERP_Backups/Test/Media'
        settings_obj.include_media = True
        settings_obj.onedrive_retention_days = 0
        settings_obj.compression_format = 'zip'      # the fake uploader inspects archives with zipfile
        settings_obj.save()
        self.addCleanup(shutil.rmtree, os.path.join(settings.BASE_DIR, 'test_backups_media'), True)

        self.uploads = []

        def fake_copy(file_path, destination):
            import zipfile as _zip

            with _zip.ZipFile(file_path) as archive:
                self.uploads.append((os.path.basename(file_path), destination, archive.namelist()))
            return f'{destination}/{os.path.basename(file_path)}'

        patchers = [
            mock.patch('backup.services.copy_to_cloud_folder', side_effect=fake_copy),
            mock.patch('backup.services.get_share_link', return_value=None),
            mock.patch.object(settings, 'MEDIA_ROOT', self.media_root),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

    def test_onedrive_run_contains_only_the_database_when_media_goes_to_google_drive(self):
        from backup.services import create_backup

        create_backup(backup_type='MANUAL', targets={'onedrive'})
        names = [name for upload in self.uploads for name in upload[2]]
        self.assertEqual(len(self.uploads), 1)
        self.assertEqual(names, ['db.sqlite3'])

    def test_google_drive_run_sends_the_database_and_mirrors_media_incrementally(self):
        from unittest import mock

        from backup.services import create_backup

        with mock.patch('backup.services.sync_media_incrementally',
                        return_value={'path': 'gdrive:ERP_Backups/Test/Media/media_files', 'files_transferred': 2}) as sync:
            history = create_backup(backup_type='MANUAL', targets={'gdrive'})
        names = [n for upload in self.uploads for n in upload[2]]
        self.assertEqual(names, ['db.sqlite3'])                       # the database zip carries no media
        sync.assert_called_once_with(self.media_root, 'gdrive:ERP_Backups/Test/Media')
        self.assertEqual(history.cloud_sync_details['media']['files_transferred'], 2)
        self.assertFalse(any(n.startswith('ERP_Media_') for n, _, _ in self.uploads))   # no daily media zip any more

    def test_a_failed_media_sync_is_reported_not_hidden(self):
        from unittest import mock

        from backup.services import create_backup

        with mock.patch('backup.services.sync_media_incrementally', side_effect=Exception('drive full')):
            history = create_backup(backup_type='MANUAL', targets={'gdrive'})
        self.assertIn('Media: drive full', history.error_message or '')

    def test_media_is_still_bundled_when_it_has_no_destination_of_its_own(self):
        from backup.models import BackupSetting
        from backup.services import create_backup

        settings_obj = BackupSetting.get_settings()
        settings_obj.media_cloud_folder = ''
        settings_obj.save()
        create_backup(backup_type='MANUAL', targets={'onedrive'})
        names = [name for upload in self.uploads for name in upload[2]]
        self.assertTrue(any(n.endswith('po.pdf') for n in names))      # old behaviour kept for single-zip setups


class IncrementalMediaSyncTests(TestCase):
    def setUp(self):
        import shutil
        import tempfile

        self.src = tempfile.mkdtemp(prefix='media_src_')
        self.dst = tempfile.mkdtemp(prefix='media_dst_')
        for path in (self.src, self.dst):
            self.addCleanup(shutil.rmtree, path, True)

    def write(self, relative, content=b'pdf'):
        full = os.path.join(self.src, relative)
        os.makedirs(os.path.dirname(full), exist_ok=True)
        with open(full, 'wb') as handle:
            handle.write(content)
        return full

    def test_first_run_copies_everything_and_the_second_copies_nothing(self):
        from backup.services import sync_media_incrementally

        self.write('planning/po_docs/a.pdf')
        self.write('item_requests/q.pdf')
        first = sync_media_incrementally(self.src, self.dst)
        second = sync_media_incrementally(self.src, self.dst)
        self.assertEqual(first['files_transferred'], 2)
        self.assertEqual(second['files_transferred'], 0)
        self.assertTrue(os.path.isfile(os.path.join(self.dst, 'media_files', 'planning', 'po_docs', 'a.pdf')))

    def test_only_new_files_are_transferred_later(self):
        from backup.services import sync_media_incrementally

        self.write('planning/po_docs/a.pdf')
        sync_media_incrementally(self.src, self.dst)
        self.write('planning/po_docs/b.pdf')
        self.assertEqual(sync_media_incrementally(self.src, self.dst)['files_transferred'], 1)

    def test_a_changed_file_is_sent_again(self):
        from backup.services import sync_media_incrementally

        path = self.write('planning/po_docs/a.pdf', b'one')
        sync_media_incrementally(self.src, self.dst)
        self.write('planning/po_docs/a.pdf', b'two-longer')
        os.utime(path, (os.path.getmtime(path) + 10, os.path.getmtime(path) + 10))
        self.assertEqual(sync_media_incrementally(self.src, self.dst)['files_transferred'], 1)

    def test_files_removed_from_the_server_stay_in_the_backup(self):
        from backup.services import sync_media_incrementally

        path = self.write('planning/po_docs/a.pdf')
        sync_media_incrementally(self.src, self.dst)
        os.remove(path)
        sync_media_incrementally(self.src, self.dst)
        self.assertTrue(os.path.isfile(os.path.join(self.dst, 'media_files', 'planning', 'po_docs', 'a.pdf')))

    def test_rclone_remote_uses_copy_never_sync(self):
        import subprocess
        from unittest import mock

        from backup.services import sync_media_incrementally

        stderr = '2026/10/21 INFO  : a.pdf: Copied (new)\n2026/10/21 INFO  : b.pdf: Copied (new)\n'
        completed = subprocess.CompletedProcess([], 0, stdout='', stderr=stderr)
        with mock.patch('backup.services.is_rclone_remote', return_value=True), \
                mock.patch('backup.services.subprocess.run', return_value=completed) as run:
            summary = sync_media_incrementally('/app/media', 'gdrive:ERP_Backups/CloudVM_A1/Media')
        command = run.call_args.args[0]
        self.assertEqual(command[:2], ['rclone', 'copy'])
        self.assertNotIn('sync', command)
        self.assertEqual(command[3], 'gdrive:ERP_Backups/CloudVM_A1/Media/media_files')
        self.assertEqual(summary['files_transferred'], 2)

    def test_a_failed_rclone_copy_raises_so_the_backup_reports_it(self):
        import subprocess
        from unittest import mock

        from backup.services import sync_media_incrementally

        with mock.patch('backup.services.is_rclone_remote', return_value=True),                 mock.patch('backup.services.subprocess.run',
                           return_value=subprocess.CompletedProcess([], 1, stdout='', stderr='quota exceeded')):
            with self.assertRaises(Exception) as ctx:
                sync_media_incrementally('/app/media', 'gdrive:x')
        self.assertIn('quota exceeded', str(ctx.exception))


class SevenZipBackupTests(TestCase):
    """Database backups are 7-Zip archives when 7-Zip is available, and plain zips otherwise."""

    def setUp(self):
        import shutil
        import tempfile

        from backup.archive import find_7zip

        self.have_7z = bool(find_7zip())
        self.work = tempfile.mkdtemp(prefix='sevenz_')
        self.addCleanup(shutil.rmtree, self.work, True)
        self.addCleanup(shutil.rmtree, os.path.join(settings.BASE_DIR, 'test_backups_7z'), True)

    def make_sqlite(self, path, rows=2000):
        import sqlite3

        con = sqlite3.connect(path)
        con.execute('CREATE TABLE t (id INTEGER PRIMARY KEY, body TEXT)')
        con.executemany('INSERT INTO t (body) VALUES (?)', [(f'<tr><td>row {i}</td><td>some repeated html text</td></tr>',) for i in range(rows)])
        con.commit()
        con.close()
        return path

    def test_7z_round_trip_is_valid_and_smaller_than_a_zip(self):
        if not self.have_7z:
            self.skipTest('7-Zip not installed here')
        import zipfile

        from backup.archive import create_7z, extract_all, list_entries

        db = self.make_sqlite(os.path.join(self.work, 'src.db'), rows=20000)
        archive = create_7z(os.path.join(self.work, 'out.7z'), db, 'db.sqlite3')
        self.assertEqual(list(list_entries(archive)), ['db.sqlite3'])
        self.assertEqual(list_entries(archive)['db.sqlite3'], os.path.getsize(db))
        extract_all(archive, os.path.join(self.work, 'x'))
        self.assertEqual(open(db, 'rb').read(), open(os.path.join(self.work, 'x', 'db.sqlite3'), 'rb').read())
        zip_path = os.path.join(self.work, 'out.zip')
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.write(db, 'db.sqlite3')
        self.assertLess(os.path.getsize(archive), os.path.getsize(zip_path))

    def test_zip_archives_are_still_listed_and_extracted(self):
        import zipfile

        from backup.archive import extract_all, list_entries

        path = os.path.join(self.work, 'a.zip')
        with zipfile.ZipFile(path, 'w') as z:
            z.writestr('db.sqlite3', b'hello')
        self.assertEqual(list_entries(path), {'db.sqlite3': 5})
        extract_all(path, os.path.join(self.work, 'out'))
        self.assertEqual(open(os.path.join(self.work, 'out', 'db.sqlite3'), 'rb').read(), b'hello')

    def test_a_missing_7zip_binary_is_reported_clearly(self):
        from unittest import mock

        from backup.archive import ArchiveError, create_7z, list_entries

        with mock.patch('backup.archive.find_7zip', return_value=None):
            with self.assertRaises(ArchiveError):
                create_7z(os.path.join(self.work, 'x.7z'), self.make_sqlite(os.path.join(self.work, 's.db')), 'db.sqlite3')
            with self.assertRaises(ArchiveError):
                list_entries(os.path.join(self.work, 'x.7z'))

    def _run_backup(self, **overrides):
        from backup.models import BackupSetting
        from backup.services import create_backup, get_database_engine

        if get_database_engine() != 'sqlite':
            self.skipTest('SQLite-specific backup test')
        settings_obj = BackupSetting.get_settings()
        settings_obj.local_backup_folder = 'test_backups_7z'
        settings_obj.cloud_onedrive_folder = None
        settings_obj.cloud_gdrive_folder = None
        settings_obj.media_cloud_folder = ''
        settings_obj.include_media = False
        for key, value in overrides.items():
            setattr(settings_obj, key, value)
        settings_obj.save()
        return create_backup(backup_type='MANUAL')

    def test_a_database_backup_is_a_7z_when_7zip_is_available(self):
        if not self.have_7z:
            self.skipTest('7-Zip not installed here')
        from backup.restore import simulate_restore

        history = self._run_backup(compression_format='7z')
        self.assertEqual(history.status, 'SUCCESS')
        self.assertTrue(history.file_name.endswith('.7z'))
        self.assertTrue(os.path.isfile(os.path.join(settings.BASE_DIR, 'test_backups_7z', history.file_name)))
        result = simulate_restore(history.id)
        self.assertTrue(result.is_valid, result.message)       # the restore path can read a .7z too

    def test_zip_format_can_be_chosen(self):
        history = self._run_backup(compression_format='zip')
        self.assertTrue(history.file_name.endswith('.zip'))

    def test_without_7zip_the_backup_falls_back_to_a_zip(self):
        from unittest import mock

        with mock.patch('backup.archive.find_7zip', return_value=None):
            history = self._run_backup(compression_format='7z')
        self.assertEqual(history.status, 'SUCCESS')
        self.assertTrue(history.file_name.endswith('.zip'))

    def test_a_backup_that_bundles_logs_stays_a_zip(self):
        history = self._run_backup(compression_format='7z', include_logs=True)
        self.assertTrue(history.file_name.endswith('.zip'))

    def test_env_override_of_the_format_wins(self):
        from unittest import mock

        with mock.patch.dict(os.environ, {'BACKUP_COMPRESSION_FORMAT': 'zip'}):
            history = self._run_backup(compression_format='7z')
        self.assertTrue(history.file_name.endswith('.zip'))

    def test_onedrive_pruning_covers_7z_files_too(self):
        from backup.services import BACKUP_ZIP_RE

        self.assertTrue(BACKUP_ZIP_RE.match('ERP_Backup_sqlite_v1_CloudVM_A1_2026-10-21_131517.7z'))
        self.assertTrue(BACKUP_ZIP_RE.match('ERP_Backup_sqlite_v1_CloudVM_A1_2026-10-21_131517.zip'))
        self.assertFalse(BACKUP_ZIP_RE.match('ERP_Backup_notes.txt'))
        self.assertFalse(BACKUP_ZIP_RE.match('other.7z'))

    def test_sqlite_snapshot_is_compact_and_valid(self):
        import sqlite3

        from backup.services import run_sqlite_backup
        from core.models import ChangeLog

        for i in range(300):
            ChangeLog.objects.create(entity_type='job_card', record_id=i, record_label=f'JC-{i}', action='update',
                                     field_changes={'x': 'y' * 200}, change_reason='r')
        ChangeLog.objects.all().delete()          # leaves free pages behind in the live database
        out = os.path.join(self.work, 'snap.db')
        run_sqlite_backup(out)
        con = sqlite3.connect(out)
        self.assertEqual(con.execute('PRAGMA integrity_check').fetchone()[0], 'ok')
        self.assertEqual(con.execute('PRAGMA freelist_count').fetchone()[0], 0)
        con.close()


class PreBackupCleanupTests(TestCase):
    """create_backup trims short-lived data first, and a cleanup failure never stops a backup."""

    def setUp(self):
        import shutil

        from backup.models import BackupSetting
        from backup.services import get_database_engine

        if get_database_engine() != 'sqlite':
            self.skipTest('SQLite-specific backup test')
        settings_obj = BackupSetting.get_settings()
        settings_obj.local_backup_folder = 'test_backups_retention'
        settings_obj.cloud_onedrive_folder = None
        settings_obj.cloud_gdrive_folder = None
        settings_obj.compression_format = 'zip'
        settings_obj.save()
        self.addCleanup(shutil.rmtree, os.path.join(settings.BASE_DIR, 'test_backups_retention'), True)

    def test_old_notifications_are_removed_before_the_backup_is_taken(self):
        import datetime

        from core.models import Notification

        user = User.objects.create_user('cleanup_user', password='pw12345678')
        row = Notification.objects.create(user=user, event_type='x', title='old')
        Notification.objects.filter(pk=row.pk).update(created_at=timezone.now() - datetime.timedelta(days=90))
        history = create_backup(backup_type='MANUAL')
        self.assertEqual(history.status, 'SUCCESS')
        self.assertEqual(Notification.objects.count(), 0)

    def test_a_cleanup_failure_does_not_stop_the_backup(self):
        from unittest import mock

        with mock.patch('core.data_retention.prune_old_data', side_effect=RuntimeError('boom')):
            history = create_backup(backup_type='MANUAL')
        self.assertEqual(history.status, 'SUCCESS')
