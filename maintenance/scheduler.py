"""In-process scheduler for preventive maintenance.

Follows the project's existing convention (see bot/scheduler.py and
backup/scheduler.py): there is no Celery/cron, so a small daemon thread started
from AppConfig.ready() does the periodic work. Every run is idempotent - the
database refuses a second ticket for the same plan and due occurrence - so
overlapping ticks, restarts, or two servers cannot create duplicates.

Disable with MAINTENANCE_INPROCESS_SCHEDULER=False (set it on any standby
server that mirrors the primary's database). Manual fallback at any time:

    python manage.py generate_due_pm_records
"""
import logging
import os
import sys
import threading
import time

from django.conf import settings
from django.db import connection

logger = logging.getLogger(__name__)

THREAD_NAME = 'DjangoMaintenancePMScheduler'
TICK_SECONDS = 30 * 60
STARTUP_DELAY_SECONDS = 60

_running = False


class PMSchedulerThread(threading.Thread):
    def __init__(self):
        super().__init__()
        self.daemon = True
        self.name = THREAD_NAME

    def run(self):
        global _running
        _running = True
        logger.info('Maintenance PM scheduler thread started.')
        time.sleep(STARTUP_DELAY_SECONDS)
        while _running:
            try:
                connection.close_if_unusable_or_obsolete()
                run_once()
            except Exception:
                logger.exception('Error in maintenance PM scheduler tick')
            finally:
                connection.close()
            time.sleep(TICK_SECONDS)


def run_once():
    from django.contrib.auth import get_user_model

    from maintenance.services import generate_due_pm_records, remind_overdue_pm

    actor = get_user_model().objects.filter(is_superuser=True).order_by('id').first()
    if actor is None:
        logger.warning('PM scheduler: no superuser to attribute generated tickets to; skipping.')
        return None
    result = generate_due_pm_records(actor=actor)
    remind_overdue_pm(actor=actor)
    if result['created'] or result['failed']:
        logger.info('PM scheduler: %d created, %d failed.', len(result['created']), len(result['failed']))
    return result


def stop_scheduler():
    global _running
    _running = False


def start_scheduler():
    if 'test' in sys.argv or any('test' in arg for arg in sys.argv):
        return
    # Only run for the long-lived server, not for migrate/shell/other commands.
    if getattr(settings, 'MAINTENANCE_INPROCESS_SCHEDULER', True) is False:
        logger.info('Maintenance PM scheduler disabled via settings. Skipping.')
        return
    argv = ' '.join(sys.argv)
    is_server = 'runserver' in argv or 'daphne' in argv or 'asgi' in argv.lower() or 'gunicorn' in argv
    if not is_server:
        return
    if ('runserver' in sys.argv) and ('--noreload' not in sys.argv) and os.environ.get('RUN_MAIN') != 'true':
        return
    for thread in threading.enumerate():
        if thread.name == THREAD_NAME:
            return
    PMSchedulerThread().start()
    logger.info('Maintenance PM scheduler started.')
