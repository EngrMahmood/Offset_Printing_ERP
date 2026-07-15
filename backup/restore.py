import os
import zipfile
import shutil
import logging
import sqlite3
import datetime
import subprocess
from django.conf import settings
from django.db import connections, transaction
from django.utils import timezone
from django.contrib.auth import get_user_model
from backup.models import BackupHistory, RestoreHistory
from backup.services import get_database_engine, calculate_sha256

logger = logging.getLogger(__name__)
User = get_user_model()

class RestoreSimulationResult:
    def __init__(self, is_valid=True, message="", db_size=0, file_count=0):
        self.is_valid = is_valid
        self.message = message
        self.db_size = db_size
        self.file_count = file_count

def simulate_restore(backup_record_id):
    """
    Perform a dry run / simulation of the restore to verify integrity of the backup file.
    Does not write to the active database.
    """
    try:
        backup = BackupHistory.objects.get(id=backup_record_id)
    except BackupHistory.DoesNotExist:
        return RestoreSimulationResult(is_valid=False, message="Backup record not found in database.")

    # Find file path
    if not backup.backup_location:
        return RestoreSimulationResult(is_valid=False, message="No backup file location stored in database.")
    
    paths = [p.strip() for p in backup.backup_location.split(",")]
    zip_path = None
    for p in paths:
        if os.path.exists(p):
            zip_path = p
            break
            
    if not zip_path:
        return RestoreSimulationResult(is_valid=False, message="Physical backup ZIP file could not be found at any listed location.")

    # Validate checksum if it exists
    if backup.sha256_checksum:
        current_checksum = calculate_sha256(zip_path)
        if current_checksum != backup.sha256_checksum:
            return RestoreSimulationResult(is_valid=False, message="File integrity check failed: Checksum mismatch.")

    # Inspect ZIP structure
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_file:
            namelist = zip_file.namelist()
            engine = get_database_engine()
            
            db_in_zip_name = "db.sqlite3" if engine == 'sqlite' else "db.dump"
            
            if db_in_zip_name not in namelist:
                return RestoreSimulationResult(is_valid=False, message=f"Backup archive is missing the core database file: {db_in_zip_name}")
            
            # Get db file size in zip
            db_info = zip_file.getinfo(db_in_zip_name)
            return RestoreSimulationResult(
                is_valid=True,
                message="Backup file structure and integrity verified successfully. Ready to restore.",
                db_size=db_info.file_size,
                file_count=len(namelist)
            )
    except zipfile.BadZipFile:
        return RestoreSimulationResult(is_valid=False, message="The backup file is not a valid ZIP archive.")
    except Exception as e:
        return RestoreSimulationResult(is_valid=False, message=f"Verification failed: {str(e)}")

def perform_restore(backup_record_id, user=None):
    """
    Restores the database from a backup record.
    """
    settings_dir = os.path.join(settings.BASE_DIR, 'backups', 'temp_restore')
    os.makedirs(settings_dir, exist_ok=True)
    
    restore_record = RestoreHistory.objects.create(
        timestamp=timezone.now(),
        status='FAILED',
        executed_by=user
    )

    try:
        backup = BackupHistory.objects.get(id=backup_record_id)
        restore_record.backup = backup
        restore_record.save()
        
        # Verify first
        sim = simulate_restore(backup_record_id)
        if not sim.is_valid:
            raise Exception(f"Restore verification failed: {sim.message}")
            
        paths = [p.strip() for p in backup.backup_location.split(",")]
        zip_path = None
        for p in paths:
            if os.path.exists(p):
                zip_path = p
                break
                
        # Extract files to temporary directory
        logger.info(f"Extracting backup zip for restoration: {zip_path}")
        with zipfile.ZipFile(zip_path, 'r') as zip_file:
            zip_file.extractall(settings_dir)
            
        engine = get_database_engine()
        
        if engine == 'sqlite':
            extracted_db_path = os.path.join(settings_dir, "db.sqlite3")
            if not os.path.exists(extracted_db_path):
                raise Exception("Extracted SQLite database file not found in temp directory.")
                
            logger.info("Performing SQLite native backup-restore process")
            # Open source connection (the extracted database file)
            src_conn = sqlite3.connect(extracted_db_path)
            
            # Get connection to the active live database
            dest_conn = connections['default'].cursor().connection
            
            # Backup from src_conn to dest_conn (restores the database page-by-page safely)
            with src_conn:
                src_conn.backup(dest_conn)
                
            src_conn.close()
            logger.info("SQLite native backup-restore process completed successfully")
            
        elif engine == 'postgresql':
            extracted_db_path = os.path.join(settings_dir, "db.dump")
            if not os.path.exists(extracted_db_path):
                raise Exception("Extracted PostgreSQL dump file not found in temp directory.")
                
            logger.info("Performing PostgreSQL pg_restore process")
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
                'pg_restore',
                '-h', db_host,
                '-p', str(db_port),
                '-U', db_user,
                '-d', db_name,
                '-c', # Clean database objects before restoring
                '-v',
                extracted_db_path
            ]
            
            result = subprocess.run(cmd, env=env, capture_output=True, text=True)
            if result.returncode != 0:
                raise Exception(f"pg_restore failed: {result.stderr}")
                
            logger.info("PostgreSQL pg_restore completed successfully")
            
        # Optional: restore media files from zip if media exists
        # We can extract any media folder contents back into settings.MEDIA_ROOT
        if hasattr(settings, 'MEDIA_ROOT') and settings.MEDIA_ROOT:
            media_src = os.path.join(settings_dir, 'media')
            if os.path.exists(media_src):
                logger.info("Restoring media folder files")
                os.makedirs(settings.MEDIA_ROOT, exist_ok=True)
                for root, dirs, files in os.walk(media_src):
                    for file in files:
                        src_file = os.path.join(root, file)
                        rel_path = os.path.relpath(src_file, media_src)
                        dest_file = os.path.join(settings.MEDIA_ROOT, rel_path)
                        os.makedirs(os.path.dirname(dest_file), exist_ok=True)
                        shutil.copy2(src_file, dest_file)
                        
        # Re-create the RestoreHistory record in the newly restored database
        RestoreHistory.objects.create(
            backup_id=backup_record_id,
            timestamp=timezone.now(),
            status='SUCCESS',
            executed_by=user
        )
        
        # Also ensure that the restored backup history record itself is marked as SUCCESS
        # and has its metadata preserved in the restored database
        try:
            restored_backup, created = BackupHistory.objects.get_or_create(id=backup_record_id)
            restored_backup.status = 'SUCCESS'
            restored_backup.file_name = backup.file_name
            restored_backup.file_size = backup.file_size
            restored_backup.backup_location = backup.backup_location
            restored_backup.sha256_checksum = backup.sha256_checksum
            restored_backup.finish_time = backup.finish_time
            restored_backup.duration_seconds = backup.duration_seconds
            restored_backup.save()
        except Exception as ex:
            logger.error(f"Could not update restored backup record: {str(ex)}")
            
        logger.info(f"Database successfully restored from backup ID {backup_record_id}")
        return True
        
    except Exception as e:
        logger.exception("Restore operation failed:")
        restore_record.status = 'FAILED'
        restore_record.error_message = str(e)
        restore_record.save()
        raise e
        
    finally:
        # Clean up temp folder
        if os.path.exists(settings_dir):
            try:
                shutil.rmtree(settings_dir)
            except Exception as ex:
                logger.error(f"Could not clean up restore temp folder: {str(ex)}")
