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
