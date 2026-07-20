import os
import time
import datetime
import threading
import logging
from django.conf import settings
from django.utils import timezone
from django.db import connection

logger = logging.getLogger(__name__)

# Global flag to control thread execution (used for cleanup/reloads)
_scheduler_running = False

class BackupSchedulerThread(threading.Thread):
    def __init__(self):
        super().__init__()
        self.daemon = True
        self.name = "DjangoBackupScheduler"

    def run(self):
        global _scheduler_running
        _scheduler_running = True
        logger.info("Backup background scheduler thread started.")
        
        # Give Django time to initialize completely
        time.sleep(10)
        
        while _scheduler_running:
            try:
                # Close old connections to avoid database pool problems on tick
                connection.close_if_unusable_or_obsolete()
                
                if self.is_backup_due():
                    logger.info("Scheduler detected that a backup is due. Starting backup thread...")
                    # Run backup in a separate worker thread so scheduler doesn't lock up
                    worker = threading.Thread(target=self.trigger_backup, name="BackupWorker")
                    worker.start()
            except Exception as e:
                logger.error(f"Error in backup scheduler tick: {str(e)}")
            
            # Tick every 60 seconds
            time.sleep(60)

    def is_backup_due(self):
        from backup.models import BackupSetting, BackupHistory
        
        # Guard clause: check if settings table exists yet
        try:
            settings_obj = BackupSetting.get_settings()
        except Exception:
            return False

        if not settings_obj.backup_enabled:
            return False
            
        now = timezone.localtime(timezone.now())
        scheduled_time = settings_obj.backup_time
        
        # Ensure we are past the scheduled time today
        if now.time() < scheduled_time:
            return False
            
        # Frequency checks
        if settings_obj.frequency == 'WEEKLY':
            # Sunday is 6 (Monday is 0, Sunday is 6)
            if now.weekday() != 6:
                return False
        elif settings_obj.frequency == 'MONTHLY':
            if now.day != 1:
                return False
                
        today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)

        # Do not launch a new backup if one is already in progress. Without this,
        # a slow or stuck run would cause the 60s tick to spawn a new worker every
        # minute (the "retry storm" seen as many Pending rows).
        in_progress = BackupHistory.objects.filter(
            status='PENDING',
            backup_type='AUTO',
            start_time__gte=today_start,
        ).exists()
        if in_progress:
            return False

        # Only ONE automatic attempt per scheduled window: if any AUTO backup has
        # already been attempted today (success OR failure), don't auto-retry every
        # minute. Admins can retry via "Create Backup Now". This is what stops the
        # every-minute storm even when a run fails.
        already_attempted = BackupHistory.objects.filter(
            backup_type='AUTO',
            start_time__gte=today_start,
        ).exists()

        return not already_attempted

    def trigger_backup(self):
        from backup.services import create_backup
        try:
            # Run the backup service
            create_backup(backup_type='AUTO', user=None)
        except Exception as e:
            logger.error(f"Error executing backup task: {str(e)}")


def start_scheduler():
    import sys
    # Do not run scheduler under unit tests
    if 'test' in sys.argv or any('test' in arg for arg in sys.argv):
        logger.info("Running under test runner. Skipping scheduler start.")
        return

    # The in-process background scheduler is disabled by default. Daily backups
    # are driven by the OS scheduler (Windows Task Scheduler -> manage.py run_backup),
    # which runs in a clean one-shot process and does not depend on the web server.
    # To re-enable the in-process thread (e.g. on a dev box), set
    # BACKUP_INPROCESS_SCHEDULER = True in settings.
    if not getattr(settings, 'BACKUP_INPROCESS_SCHEDULER', False):
        logger.info("In-process backup scheduler disabled (using OS scheduler). Skipping.")
        return

    # Only start the scheduler in the main process thread to avoid running duplicate schedulers
    # In Django dev server, RUN_MAIN=true is set in the reloader subprocess
    if settings.DEBUG and os.environ.get('RUN_MAIN') != 'true':
        logger.info("DEBUG is True and RUN_MAIN is not true. Skipping scheduler start in parent process.")
        return

    # Guard to prevent duplicate threads starting
    for thread in threading.enumerate():
        if thread.name == "DjangoBackupScheduler":
            logger.info("Backup scheduler thread is already running.")
            return

    scheduler = BackupSchedulerThread()
    scheduler.start()
