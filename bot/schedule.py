"""Due-ness and next-run calculation for bots.

Pure functions over a BotAutomation + a local datetime, so the scheduling rules
can be unit-tested without starting the scheduler thread.
"""
from __future__ import annotations

import calendar
import datetime

from django.utils import timezone

from bot.models import (
    FREQUENCY_DAILY,
    FREQUENCY_MONTHLY,
    FREQUENCY_WEEKLY,
    STATUS_FAILED,
    STATUS_SENT,
    STATUS_SKIPPED,
    TRIGGER_AUTO,
)


def _clamp_day(year: int, month: int, day: int) -> int:
    """A bot set to the 31st still fires in February — on the last day."""
    return min(day, calendar.monthrange(year, month)[1])


def matches_calendar_day(bot, day: datetime.date) -> bool:
    """Is `day` a day this bot is configured to run on (ignoring the clock)?"""
    if bot.start_date and day < bot.start_date:
        return False
    if bot.end_date and day > bot.end_date:
        return False

    if bot.frequency == FREQUENCY_MONTHLY:
        target = _clamp_day(day.year, day.month, bot.day_of_month or 1)
        return day.day == target

    weekdays = bot.weekday_numbers
    if bot.frequency == FREQUENCY_WEEKLY:
        # Weekly with no explicit weekday defaults to Monday.
        return day.weekday() in (weekdays or {0})

    # DAILY: blank weekdays means every day, otherwise only the listed ones.
    return not weekdays or day.weekday() in weekdays


def _scheduled_datetime(bot, day: datetime.date):
    """Timezone-aware datetime for this bot's send time on `day`."""
    naive = datetime.datetime.combine(day, bot.send_time)
    return timezone.make_aware(naive, timezone.get_current_timezone())


def calculate_next_run(bot, from_dt=None):
    """First scheduled moment strictly after `from_dt`. None if the bot's
    end_date has already passed."""
    now = timezone.localtime(from_dt or timezone.now())
    day = now.date()
    # 400 days covers a monthly bot whose day_of_month sits just past end_date.
    for offset in range(0, 400):
        candidate_day = day + datetime.timedelta(days=offset)
        if bot.end_date and candidate_day > bot.end_date:
            return None
        if not matches_calendar_day(bot, candidate_day):
            continue
        candidate = _scheduled_datetime(bot, candidate_day)
        if candidate > now:
            return candidate
    return None


def _todays_executions(bot, now):
    from bot.models import BotExecution

    day_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    return BotExecution.objects.filter(bot=bot, trigger=TRIGGER_AUTO, started_at__gte=day_start)


def is_due(bot, now=None) -> bool:
    """Should the scheduler run this bot right now?

    Mirrors the guards proven in backup/scheduler.py: only after the scheduled
    time, at most one successful run per window, no second run while one is in
    flight, and a cooldown between retries so a persistent failure can't spawn
    a run on every 60s tick.
    """
    if not bot.is_active:
        return False

    now = timezone.localtime(now or timezone.now())

    if not matches_calendar_day(bot, now.date()):
        return False
    if now.time() < bot.send_time:
        return False

    executions = _todays_executions(bot, now)

    # Already delivered (or deliberately skipped as empty) in this window.
    if executions.filter(status__in=(STATUS_SENT, STATUS_SKIPPED)).exists():
        return False

    # A run is already in flight — never launch a second one.
    if executions.filter(status='PENDING').exists():
        return False

    failures = executions.filter(status=STATUS_FAILED)
    failure_count = failures.count()
    if failure_count:
        if failure_count > bot.retry_count:
            return False
        cooldown = now - datetime.timedelta(minutes=bot.retry_interval_minutes or 15)
        if failures.filter(started_at__gte=cooldown).exists():
            return False

    return True


def due_bots(now=None):
    """Active bots that should run on this tick."""
    from bot.models import BotAutomation, BotGlobalSettings

    if not BotGlobalSettings.get_settings().automation_enabled:
        return []

    now = timezone.localtime(now or timezone.now())
    return [bot for bot in BotAutomation.objects.filter(is_active=True) if is_due(bot, now)]
