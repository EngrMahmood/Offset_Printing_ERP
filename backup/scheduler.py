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

        # Already completed successfully today? Nothing more to do this window.
        if BackupHistory.objects.filter(
            status='SUCCESS', backup_type='AUTO', start_time__gte=today_start,
        ).exists():
            return False

        # A run is already in progress? Don't launch a second one.
        if BackupHistory.objects.filter(
            status='PENDING', backup_type='AUTO', start_time__gte=today_start,
        ).exists():
            return False

        # Cooldown after a failed/partial attempt: wait before retrying so a
        # persistent failure can't spawn a new backup on every 60s tick (the
        # "retry storm"). Failures are retried at most every 15 minutes; the
        # first SUCCESS above stops retries for the rest of the day.
        cooldown_start = now - datetime.timedelta(minutes=15)
        if BackupHistory.objects.filter(
            backup_type='AUTO', start_time__gte=cooldown_start,
        ).exists():
            return False

        return True

    def trigger_backup(self):
        from backup.services import create_backup
        try:
            # Run the backup service
            create_backup(backup_type='AUTO', user=None)
        except Exception as e:
            logger.error(f"Error executing backup task: {str(e)}")


def start_scheduler():
    import sys

    # Never run the scheduler under the test runner.
    if 'test' in sys.argv or any('test' in arg for arg in sys.argv):
        logger.info("Running under test runner. Skipping scheduler start.")
        return

    # Opt-out switch. Defaults to ON so auto-backup works out of the box on ANY
    # server the same way -- dev or production, DEBUG on or off, runserver
    # (with or without --noreload) or a WSGI server (gunicorn/waitress/uWSGI/IIS).
    # Set BACKUP_INPROCESS_SCHEDULER = False in settings only to disable it.
    if getattr(settings, 'BACKUP_INPROCESS_SCHEDULER', True) is False:
        logger.info("In-process backup scheduler disabled via settings. Skipping.")
        return

    # Avoid a duplicate scheduler under Django's autoreloader. The reloader runs
    # two processes: a watcher parent and a child with RUN_MAIN=true. Start only
    # in the child. This guard does NOT apply to `runserver --noreload` or to WSGI
    # servers (no 'runserver' in argv), where the scheduler must start normally --
    # this is exactly the case (`--noreload`) that previously left production with
    # no auto-backup.
    running_under_reloader = ('runserver' in sys.argv) and ('--noreload' not in sys.argv)
    if running_under_reloader and os.environ.get('RUN_MAIN') != 'true':
        logger.info("Reloader parent detected; the child process will start the scheduler.")
        return

    # Guard against duplicate threads within a single process.
    for thread in threading.enumerate():
        if thread.name == "DjangoBackupScheduler":
            logger.info("Backup scheduler thread is already running.")
            return

    BackupSchedulerThread().start()
    logger.info("In-process backup scheduler started.")
