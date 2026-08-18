"""Due-ness and send logic for recurring task reminders. Pure functions over
a Task + TaskNotificationSettings + an optional `now`, so the scheduling rules
are testable without starting the scheduler thread — mirrors bot/schedule.py."""
from __future__ import annotations

import logging

from django.conf import settings
from django.core.mail import send_mail
from django.template.loader import render_to_string
from django.utils import timezone

from .models import REMIND_FROM_OVERDUE, Task, TaskNotificationSettings, TaskReminderLog

logger = logging.getLogger(__name__)


def _last_reminder_or_created_date(task):
    last = task.reminder_logs.filter(status=TaskReminderLog.STATUS_SENT).order_by('-sent_at').first()
    return (last.sent_at if last else task.created_at).date()


def is_reminder_due(task, settings_obj, now=None) -> bool:
    if not settings_obj.reminders_enabled:
        return False
    if task.status not in ('pending', 'in_progress'):
        return False
    if not (task.assignee_id or task.assigned_team_id):
        return False

    now = now or timezone.now()

    if settings_obj.remind_from == REMIND_FROM_OVERDUE:
        if task.due_date is None or now.date() <= task.due_date:
            return False
        anchor = max(task.due_date, _last_reminder_or_created_date(task))
    else:
        anchor = _last_reminder_or_created_date(task)

    return (now.date() - anchor).days >= settings_obj.reminder_interval_days


def due_reminder_tasks(now=None):
    settings_obj = TaskNotificationSettings.get_solo()
    if not settings_obj.reminders_enabled:
        return []
    candidates = (
        Task.objects.filter(status__in=['pending', 'in_progress'])
        .exclude(assignee__isnull=True, assigned_team__isnull=True)
        .select_related('assignee', 'assigned_team')
    )
    return [t for t in candidates if is_reminder_due(t, settings_obj, now)]


def send_reminder_email(task):
    """Sends the recurring "still pending" reminder and writes a TaskReminderLog
    row. Never raises to the caller — failures are recorded, not propagated,
    matching bot/services.py::run_bot's outer try/except."""
    from .emails import resolve_recipients, task_detail_url

    settings_obj = TaskNotificationSettings.get_solo()
    if not settings_obj.reminders_enabled:
        return

    recipients = resolve_recipients(task)
    if not recipients:
        # No valid recipient email — skip silently, no log row (a false
        # "sent" would corrupt the dedup anchor for _last_reminder_or_created_date).
        return

    detail_url = task_detail_url(task)
    is_team_assignment = bool(task.assigned_team_id)
    today = timezone.now().date()
    days_pending = (today - task.created_at.date()).days
    days_overdue = (today - task.due_date).days if task.due_date and today > task.due_date else 0

    try:
        for user, addr in recipients:
            body = render_to_string('tasks/email/reminder.html', {
                'task': task,
                'recipient': user,
                'is_team_assignment': is_team_assignment,
                'team': task.assigned_team,
                'days_pending': days_pending,
                'days_overdue': days_overdue,
                'detail_url': detail_url,
            })
            send_mail(
                subject=f"Reminder: task still pending — {task.title}",
                message=body,
                from_email=None,
                recipient_list=[addr],
                fail_silently=False,
            )
        TaskReminderLog.objects.create(
            task=task,
            status=TaskReminderLog.STATUS_SENT,
            recipients=', '.join(addr for _, addr in recipients),
        )
    except Exception as exc:
        logger.exception('Failed to send task reminder email for task id %s', task.pk)
        TaskReminderLog.objects.create(
            task=task,
            status=TaskReminderLog.STATUS_FAILED,
            recipients=', '.join(addr for _, addr in recipients),
            error_message=str(exc),
        )
