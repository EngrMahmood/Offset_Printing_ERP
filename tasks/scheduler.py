"""In-process reminder scheduler for the tasks app.

Structural copy of bot/scheduler.py's guard code (autoreloader child-only
start, test-runner skip, duplicate-thread-name guard, opt-out setting) — that
thread has proven itself in this codebase already. Kept as a separate thread
rather than folded into bot/scheduler.py because task reminders have a flat
per-task N-day interval, not the cron-like BotAutomation schedule shape; see
bot/models.py's own docstring on that boundary.
"""
import logging
import os
import sys
import threading
import time

from django.conf import settings
from django.db import connection

logger = logging.getLogger(__name__)

THREAD_NAME = 'DjangoTaskReminderScheduler'

# Reminders are day-granularity, so 60s polling (like the bot scheduler) is
# unnecessary — 5 minutes is more than enough and cuts DB load.
TICK_SECONDS = 300

# Offset from the bot scheduler's 15s startup delay so both threads don't
# hit the DB at the exact same moment on boot.
STARTUP_DELAY_SECONDS = 20

_scheduler_running = False


class TaskReminderSchedulerThread(threading.Thread):
    def __init__(self):
        super().__init__()
        self.daemon = True
        self.name = THREAD_NAME

    def run(self):
        global _scheduler_running
        _scheduler_running = True
        logger.info('Task reminder scheduler thread started.')

        time.sleep(STARTUP_DELAY_SECONDS)

        while _scheduler_running:
            try:
                connection.close_if_unusable_or_obsolete()
                self.tick()
            except Exception:
                logger.exception('Error in task reminder scheduler tick')
            time.sleep(TICK_SECONDS)

    def tick(self):
        from tasks.reminders import due_reminder_tasks

        try:
            due_tasks = due_reminder_tasks()
        except Exception:
            # Table may not exist yet (pre-migrate); stay quiet and retry later.
            logger.debug('Task reminder scheduler could not read settings/tasks yet.', exc_info=True)
            return

        for task in due_tasks:
            worker = threading.Thread(
                target=self.execute,
                args=(task.pk,),
                name=f'TaskReminderWorker-{task.pk}',
                daemon=True,
            )
            worker.start()

    @staticmethod
    def execute(task_id):
        from tasks.models import Task
        from tasks.reminders import send_reminder_email

        try:
            connection.close_if_unusable_or_obsolete()
            task = Task.objects.filter(pk=task_id).first()
            if task is None:
                return
            send_reminder_email(task)
        except Exception:
            logger.exception('Task reminder worker crashed for task id %s', task_id)
        finally:
            connection.close()


def stop_scheduler():
    """Used by tests and shutdown paths to end the loop after the current tick."""
    global _scheduler_running
    _scheduler_running = False


def start_scheduler():
    # Never run under the test runner.
    if 'test' in sys.argv or any('test' in arg for arg in sys.argv):
        logger.info('Running under test runner. Skipping task reminder scheduler start.')
        return

    # Opt-out switch, defaulting to ON. Set TASK_REMINDER_INPROCESS_SCHEDULER =
    # False in settings to disable.
    if getattr(settings, 'TASK_REMINDER_INPROCESS_SCHEDULER', True) is False:
        logger.info('In-process task reminder scheduler disabled via settings. Skipping.')
        return

    # Under Django's autoreloader two processes run: a watcher parent and a
    # child with RUN_MAIN=true. Start only in the child.
    running_under_reloader = ('runserver' in sys.argv) and ('--noreload' not in sys.argv)
    if running_under_reloader and os.environ.get('RUN_MAIN') != 'true':
        logger.info('Reloader parent detected; the child process will start the task reminder scheduler.')
        return

    for thread in threading.enumerate():
        if thread.name == THREAD_NAME:
            logger.info('Task reminder scheduler thread is already running.')
            return

    TaskReminderSchedulerThread().start()
    logger.info('In-process task reminder scheduler started.')
