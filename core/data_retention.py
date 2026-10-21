"""Housekeeping for tables that grow without bound but are only useful while recent.

Run before every backup (backup.services.create_backup) and by ``manage.py prune_old_data``.
Notifications are transient alerts; the bot email history keeps a full HTML copy of every
report email (about 67 KB each), which nobody re-reads after a couple of weeks. Both were
over half of the database.
"""
import datetime
import logging

from django.conf import settings
from django.utils import timezone

logger = logging.getLogger(__name__)

BODY_PLACEHOLDER = '[Email body removed after {days} days to save space. Subject, recipients and status are kept.]'
BATCH = 5000


def prune_old_data(notification_days=None, bot_body_days=None, now=None, dry_run=False):
    """Returns {'notifications': n, 'bot_bodies': n}: rows removed / bodies cleared (or that would be)."""
    from bot.models import BotExecution
    from core.models import Notification

    now = now or timezone.now()
    if notification_days is None:
        notification_days = getattr(settings, 'NOTIFICATION_RETENTION_DAYS', 30)
    if bot_body_days is None:
        bot_body_days = getattr(settings, 'BOT_EMAIL_BODY_RETENTION_DAYS', 14)
    result = {'notifications': 0, 'bot_bodies': 0}

    if notification_days and notification_days > 0:
        cutoff = now - datetime.timedelta(days=notification_days)
        old = Notification.objects.filter(created_at__lt=cutoff)
        if dry_run:
            result['notifications'] = old.count()
        else:
            # In batches, so a first run over ~75,000 rows never holds the database lock for long.
            while True:
                ids = list(old.values_list('pk', flat=True)[:BATCH])
                if not ids:
                    break
                Notification.objects.filter(pk__in=ids).delete()
                result['notifications'] += len(ids)

    if bot_body_days and bot_body_days > 0:
        cutoff = now - datetime.timedelta(days=bot_body_days)
        placeholder = BODY_PLACEHOLDER.format(days=bot_body_days)
        stale = BotExecution.objects.filter(started_at__lt=cutoff).exclude(rendered_body='').exclude(rendered_body=placeholder)
        if dry_run:
            result['bot_bodies'] = stale.count()
        else:
            result['bot_bodies'] = stale.update(rendered_body=placeholder)

    if not dry_run and (result['notifications'] or result['bot_bodies']):
        logger.info(
            'Data retention: removed %s old notification(s), cleared %s old bot email body(ies)',
            result['notifications'], result['bot_bodies'],
        )
    return result
