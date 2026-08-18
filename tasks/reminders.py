"""Due-ness and send logic for recurring task reminders. Pure functions over
a Task + TaskNotificationSettings + an optional `now`, so the scheduling rules
are testable without starting the scheduler thread — mirrors bot/schedule.py."""
from __future__ import annotations

import logging

from django.core.mail import EmailMessage
from django.template.loader import render_to_string
from django.utils import timezone

from .models import REMIND_FROM_OVERDUE, Task, TaskNotificationLog, TaskNotificationSettings

logger = logging.getLogger(__name__)


def _last_reminder_or_created_date(task):
    last = task.notification_logs.filter(
        kind=TaskNotificationLog.KIND_REMINDER, status=TaskNotificationLog.STATUS_SENT
    ).order_by('-sent_at').first()
    return (last.sent_at if last else task.created_at).date()


def effective_interval_days(task, settings_obj):
    """The per-task override if the task set one, else the global default.
    Checked with `is not None` rather than truthiness -- 0 is a valid
    override (remind every tick) and must not be treated as "unset"."""
    if task.reminder_interval_days is not None:
        return task.reminder_interval_days
    return settings_obj.reminder_interval_days


def is_reminder_due(task, settings_obj, now=None) -> bool:
    if not settings_obj.reminders_enabled:
        return False
    if task.status not in ('pending', 'in_progress'):
        return False
    if not (task.assignee_id or task.assigned_team_id):
        return False

    now = now or timezone.now()
    interval_days = effective_interval_days(task, settings_obj)

    if settings_obj.remind_from == REMIND_FROM_OVERDUE:
        if task.due_date is None or now.date() <= task.due_date:
            return False
        anchor = max(task.due_date, _last_reminder_or_created_date(task))
    else:
        anchor = _last_reminder_or_created_date(task)

    return (now.date() - anchor).days >= interval_days


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
    """Sends a single recurring "still pending" reminder (To = every resolved
    recipient, Cc/Bcc = the task's own overrides) and writes a
    TaskNotificationLog row. Never raises to the caller — failures are
    recorded, not propagated, matching bot/services.py::run_bot's outer
    try/except."""
    from .emails import _split_addresses, resolve_recipients, task_detail_url

    settings_obj = TaskNotificationSettings.get_solo()
    if not settings_obj.reminders_enabled:
        return

    recipients = resolve_recipients(task)
    if not recipients:
        # No valid recipient email — skip silently, no log row (a false
        # "sent" would corrupt the dedup anchor for _last_reminder_or_created_date).
        return

    to_addrs = [addr for _, addr in recipients]
    cc_addrs = _split_addresses(task.cc_emails)
    bcc_addrs = _split_addresses(task.bcc_emails)
    detail_url = task_detail_url(task)
    is_team_assignment = bool(task.assigned_team_id)
    today = timezone.now().date()
    days_pending = (today - task.created_at.date()).days
    days_overdue = (today - task.due_date).days if task.due_date and today > task.due_date else 0

    try:
        body = render_to_string('tasks/email/reminder.html', {
            'task': task,
            'recipients': [u for u, _ in recipients],
            'is_team_assignment': is_team_assignment,
            'team': task.assigned_team,
            'days_pending': days_pending,
            'days_overdue': days_overdue,
            'detail_url': detail_url,
        })
        message = EmailMessage(
            subject=f"Reminder: task still pending — {task.title}",
            body=body,
            from_email=None,
            to=to_addrs,
            cc=cc_addrs,
            bcc=bcc_addrs,
        )
        message.send(fail_silently=False)
        TaskNotificationLog.objects.create(
            task=task,
            kind=TaskNotificationLog.KIND_REMINDER,
            status=TaskNotificationLog.STATUS_SENT,
            recipients_to=', '.join(to_addrs),
            recipients_cc=', '.join(cc_addrs),
            recipients_bcc=', '.join(bcc_addrs),
        )
    except Exception as exc:
        logger.exception('Failed to send task reminder email for task id %s', task.pk)
        TaskNotificationLog.objects.create(
            task=task,
            kind=TaskNotificationLog.KIND_REMINDER,
            status=TaskNotificationLog.STATUS_FAILED,
            recipients_to=', '.join(to_addrs),
            recipients_cc=', '.join(cc_addrs),
            recipients_bcc=', '.join(bcc_addrs),
            error_message=str(exc),
        )
