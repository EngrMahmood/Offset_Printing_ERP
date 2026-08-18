"""Assignment-notice email + shared recipient resolution for the tasks app.
tasks/reminders.py reuses resolve_recipients()/_split_addresses() for the
recurring reminder send."""
import logging

from django.conf import settings
from django.core.mail import EmailMessage
from django.template.loader import render_to_string

logger = logging.getLogger(__name__)


def _split_addresses(raw):
    """Split a comma/newline/semicolon separated address blob into a clean list."""
    if not raw:
        return []
    parts = raw.replace(';', ',').replace('\n', ',').replace('\r', ',').split(',')
    return [part.strip() for part in parts if part.strip()]


def resolve_recipients(task):
    """De-duplicated list of (User, email) pairs for a task's assignee and/or
    team members. `email` is always the OFFICIAL email (core.models.notification_email
    — UserProfile.official_email if set, else User.email), never a raw personal
    address. Users with no email on file at all are skipped."""
    from core.models import notification_email

    users = []
    if task.assignee_id and task.assignee:
        users.append(task.assignee)
    if task.assigned_team_id and task.assigned_team:
        users.extend(task.assigned_team.members.filter(is_active=True))

    seen_ids = set()
    result = []
    for u in users:
        if u.id in seen_ids:
            continue
        seen_ids.add(u.id)
        addr = notification_email(u)
        if addr:
            result.append((u, addr))
    return result


def task_detail_url(task):
    base = getattr(settings, 'TASK_APP_BASE_URL', '').rstrip('/')
    return f"{base}/tasks/{task.pk}/"


def send_assignment_email(task):
    """Sends a single "you've been assigned" email to every current recipient
    (To), with the task's own CC/BCC overrides applied, and logs the attempt.
    Never raises — a notification failure must not break the save() that
    triggered it."""
    from .models import TaskNotificationLog, TaskNotificationSettings

    settings_obj = TaskNotificationSettings.get_solo()
    if not settings_obj.assignment_email_enabled:
        return

    recipients = resolve_recipients(task)
    if not recipients:
        return

    to_addrs = [addr for _, addr in recipients]
    cc_addrs = _split_addresses(task.cc_emails)
    bcc_addrs = _split_addresses(task.bcc_emails)
    detail_url = task_detail_url(task)
    is_team_assignment = bool(task.assigned_team_id)

    try:
        body = render_to_string('tasks/email/assigned.html', {
            'task': task,
            'recipients': [u for u, _ in recipients],
            'is_team_assignment': is_team_assignment,
            'team': task.assigned_team,
            'detail_url': detail_url,
        })
        message = EmailMessage(
            subject=f"Task assigned: {task.title}",
            body=body,
            from_email=None,
            to=to_addrs,
            cc=cc_addrs,
            bcc=bcc_addrs,
        )
        message.send(fail_silently=False)
        TaskNotificationLog.objects.create(
            task=task,
            kind=TaskNotificationLog.KIND_ASSIGNMENT,
            status=TaskNotificationLog.STATUS_SENT,
            recipients_to=', '.join(to_addrs),
            recipients_cc=', '.join(cc_addrs),
            recipients_bcc=', '.join(bcc_addrs),
        )
    except Exception as exc:
        logger.exception('Failed to send task assignment email for task id %s', task.pk)
        TaskNotificationLog.objects.create(
            task=task,
            kind=TaskNotificationLog.KIND_ASSIGNMENT,
            status=TaskNotificationLog.STATUS_FAILED,
            recipients_to=', '.join(to_addrs),
            recipients_cc=', '.join(cc_addrs),
            recipients_bcc=', '.join(bcc_addrs),
            error_message=str(exc),
        )
