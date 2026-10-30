"""In-process scheduler for the plan: 18:30 draft, 08:30 re-balance, 08:45 email.

Same pattern as bot/scheduler.py. It only runs on a server that is allowed to send (the primary): it
follows BOT_INPROCESS_SCHEDULER (False on the standby) as well as its own PLAN_INPROCESS_SCHEDULER, and
the bot master switch. PlanRunLog makes every task once-per-day even across restarts, and a task that
was missed (the server was down at 18:30) still runs later the same day."""
import datetime
import logging
import os
import threading
import time

from django.conf import settings
from django.db import connection
from django.utils import timezone

logger = logging.getLogger(__name__)

THREAD_NAME = 'DjangoPlanScheduler'
TICK_SECONDS = 60
STARTUP_DELAY_SECONDS = 25
EMAIL_RETRY_MINUTES = 5
EMAIL_MAX_ATTEMPTS = 4
# A missed task still runs until this long after its time.
CATCH_UP = {'draft': datetime.timedelta(hours=5), 'rebalance': datetime.timedelta(hours=3), 'email': datetime.timedelta(hours=4)}

_scheduler_running = False


def _due(task, now, scheduled_time):
    due_at = timezone.make_aware(datetime.datetime.combine(now.date(), scheduled_time))
    return due_at <= now <= due_at + CATCH_UP[task]


def _claim(task, today):
    """Create today's run row; False when this task already ran (or is running) today."""
    from production_plan.models import PlanRunLog

    row, created = PlanRunLog.objects.get_or_create(task=task, run_date=today)
    return row if created else None


def run_tick(now=None):
    """One scheduler pass. Returns the tasks it ran (for tests)."""
    from bot.models import BotGlobalSettings
    from production_plan import engine, services
    from production_plan.models import PlanRunLog, PlanSettings, ProductionPlan

    now = timezone.localtime(now or timezone.now())
    config = PlanSettings.get()
    ran = []
    if not config.scheduler_enabled or not BotGlobalSettings.get_settings().automation_enabled:
        return ran
    today = now.date()

    # Evening draft: needs a working day to plan for.
    if now.time() >= config.draft_time and _due('draft', now, config.draft_time) and engine.next_working_day(today, config) and today.weekday() != 6:
        row = _claim('draft', today)
        if row is not None:
            try:
                plan = services.build_draft(today=today, settings_obj=config)
                row.status, row.detail = 'done', f'plan {plan.pk} for {plan.morning_date}'
            except Exception as exc:
                logger.exception('Plan draft failed')
                row.status, row.detail = 'failed', str(exc)[:500]
            row.finished_at = now
            row.save()
            ran.append('draft')

    if engine.is_working_day(today, config):
        if _due('rebalance', now, config.rebalance_time):
            row = _claim('rebalance', today)
            if row is not None:
                try:
                    plan = services.rebalance_today(today=today, settings_obj=config)
                    row.status, row.detail = 'done', f'plan {plan.pk}' if plan else 'no plan'
                except Exception as exc:
                    logger.exception('Plan re-balance failed')
                    row.status, row.detail = 'failed', str(exc)[:500]
                row.finished_at = now
                row.save()
                ran.append('rebalance')

        if _due('email', now, config.email_time):
            ran += _email_step(now, today, config)
    return ran


def _email_step(now, today, config):
    from production_plan import services
    from production_plan.models import PlanRunLog, ProductionPlan

    row = PlanRunLog.objects.filter(task='email', run_date=today).first()
    if row is not None and (row.status in ('done', 'gave_up') or
                            (row.status == 'retry' and now < timezone.localtime(row.finished_at) + datetime.timedelta(minutes=EMAIL_RETRY_MINUTES))):
        return []
    plan = ProductionPlan.objects.filter(morning_date=today).exclude(status='superseded').first()
    if plan is None:
        return []                                         # the re-balance has not produced today's plan yet
    if row is None:
        row = PlanRunLog.objects.create(task='email', run_date=today)
    else:
        row.attempts += 1
    try:
        services.send_plan_email(plan, config)
        row.status, row.detail = 'done', f'plan {plan.pk} ({plan.status})'
    except Exception as exc:
        logger.exception('Plan email failed (attempt %s)', row.attempts)
        plan.email_status, plan.email_error = 'failed', str(exc)[:500]
        plan.save(update_fields=['email_status', 'email_error'])
        row.status = 'gave_up' if row.attempts >= EMAIL_MAX_ATTEMPTS else 'retry'
        row.detail = str(exc)[:500]
    row.finished_at = now
    row.save()
    return ['email']


class PlanSchedulerThread(threading.Thread):
    def __init__(self):
        super().__init__()
        self.daemon = True
        self.name = THREAD_NAME

    def run(self):
        global _scheduler_running
        _scheduler_running = True
        logger.info('Production plan scheduler thread started.')
        time.sleep(STARTUP_DELAY_SECONDS)
        while _scheduler_running:
            try:
                connection.close_if_unusable_or_obsolete()
                run_tick()
            except Exception:
                logger.exception('Error in production plan scheduler tick')
            time.sleep(TICK_SECONDS)


def stop_scheduler():
    global _scheduler_running
    _scheduler_running = False


def start_scheduler():
    import sys

    if 'test' in sys.argv or any('test' in arg for arg in sys.argv):
        return
    # Only a server that sends mail runs it (the standby has BOT_INPROCESS_SCHEDULER=False).
    if getattr(settings, 'BOT_INPROCESS_SCHEDULER', True) is False or getattr(settings, 'PLAN_INPROCESS_SCHEDULER', True) is False:
        logger.info('Production plan scheduler disabled via settings. Skipping.')
        return
    running_under_reloader = ('runserver' in sys.argv) and ('--noreload' not in sys.argv)
    if running_under_reloader and os.environ.get('RUN_MAIN') != 'true':
        return
    for thread in threading.enumerate():
        if thread.name == THREAD_NAME:
            return
    PlanSchedulerThread().start()
