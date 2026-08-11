"""In-process bot scheduler.

Structural port of backup/scheduler.py — this project has no Celery, and that
thread has proven itself on the Windows LAN server under runserver, runserver
--noreload, and Daphne alike. Every guard below is carried over deliberately.

Differences from the backup scheduler: it iterates many BotAutomation rows
instead of one singleton setting, and the due-ness rules live in bot/schedule.py
as pure functions so they are testable without threads.
"""
import logging
import os
import threading
import time

from django.conf import settings
from django.db import connection

logger = logging.getLogger(__name__)

THREAD_NAME = 'DjangoBotScheduler'

# Seconds between ticks. Matches the backup scheduler; a bot's send_time has
# minute resolution, so 60s is enough.
TICK_SECONDS = 60

# Let Django finish booting (apps, DB, cache) before the first query.
STARTUP_DELAY_SECONDS = 15

_scheduler_running = False


class BotSchedulerThread(threading.Thread):
    def __init__(self):
        super().__init__()
        self.daemon = True
        self.name = THREAD_NAME

    def run(self):
        global _scheduler_running
        _scheduler_running = True
        logger.info('Bot background scheduler thread started.')

        time.sleep(STARTUP_DELAY_SECONDS)

        while _scheduler_running:
            try:
                # Avoid handing a stale connection to the tick's queries.
                connection.close_if_unusable_or_obsolete()
                self.tick()
            except Exception:
                # Never let a tick kill the thread — the next one may succeed.
                logger.exception('Error in bot scheduler tick')
            time.sleep(TICK_SECONDS)

    def tick(self):
        from bot.schedule import due_bots

        try:
            bots = due_bots()
        except Exception:
            # Table may not exist yet (pre-migrate); stay quiet and retry later.
            logger.debug('Bot scheduler could not read automations yet.', exc_info=True)
            return

        for bot in bots:
            logger.info('Bot scheduler: %s is due. Dispatching.', bot.code)
            # Each run gets its own thread so one slow report cannot stall the
            # tick loop (and therefore every other bot).
            worker = threading.Thread(
                target=self.execute,
                args=(bot.pk,),
                name=f'BotWorker-{bot.code}',
                daemon=True,
            )
            worker.start()

    @staticmethod
    def execute(bot_id):
        from bot.models import BotAutomation
        from bot.services import run_bot

        try:
            connection.close_if_unusable_or_obsolete()
            bot = BotAutomation.objects.filter(pk=bot_id).first()
            if bot is None:
                return
            execution = run_bot(bot)
            logger.info('Bot %s finished with status %s.', bot.code, execution.status)
        except Exception:
            # run_bot already records failures; this only catches a failure to
            # even reach it (e.g. the DB went away between tick and worker).
            logger.exception('Bot worker crashed for bot id %s', bot_id)
        finally:
            connection.close()


def stop_scheduler():
    """Used by tests and shutdown paths to end the loop after the current tick."""
    global _scheduler_running
    _scheduler_running = False


def start_scheduler():
    import sys

    # Never run under the test runner — tests drive run_bot/is_due directly.
    if 'test' in sys.argv or any('test' in arg for arg in sys.argv):
        logger.info('Running under test runner. Skipping bot scheduler start.')
        return

    # Opt-out switch, defaulting to ON so the bots behave the same on every
    # server (dev or production, runserver or ASGI). Set
    # BOT_INPROCESS_SCHEDULER = False in settings to disable.
    if getattr(settings, 'BOT_INPROCESS_SCHEDULER', True) is False:
        logger.info('In-process bot scheduler disabled via settings. Skipping.')
        return

    # Under Django's autoreloader two processes run: a watcher parent and a
    # child with RUN_MAIN=true. Start only in the child. This must NOT apply to
    # `runserver --noreload` or to ASGI/WSGI servers, where the scheduler is the
    # only thing that would ever run the bots.
    running_under_reloader = ('runserver' in sys.argv) and ('--noreload' not in sys.argv)
    if running_under_reloader and os.environ.get('RUN_MAIN') != 'true':
        logger.info('Reloader parent detected; the child process will start the bot scheduler.')
        return

    for thread in threading.enumerate():
        if thread.name == THREAD_NAME:
            logger.info('Bot scheduler thread is already running.')
            return

    BotSchedulerThread().start()
    logger.info('In-process bot scheduler started.')
