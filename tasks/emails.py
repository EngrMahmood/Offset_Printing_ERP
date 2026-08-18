"""Assignment-notice email + shared recipient resolution for the tasks app.
tasks/reminders.py reuses resolve_recipients() for the recurring reminder send."""
import logging

from django.conf import settings
from django.core.mail import send_mail
from django.template.loader import render_to_string

logger = logging.getLogger(__name__)


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
    """Sends the immediate "you've been assigned" email to every current
    recipient. Never raises — a notification failure must not break the
    save() that triggered it."""
    from .models import TaskNotificationSettings

    settings_obj = TaskNotificationSettings.get_solo()
    if not settings_obj.assignment_email_enabled:
        return

    recipients = resolve_recipients(task)
    if not recipients:
        return

    detail_url = task_detail_url(task)
    is_team_assignment = bool(task.assigned_team_id)

    for user, addr in recipients:
        try:
            body = render_to_string('tasks/email/assigned.html', {
                'task': task,
                'recipient': user,
                'is_team_assignment': is_team_assignment,
                'team': task.assigned_team,
                'detail_url': detail_url,
            })
            send_mail(
                subject=f"Task assigned to you: {task.title}",
                message=body,
                from_email=None,
                recipient_list=[addr],
                fail_silently=True,
            )
        except Exception:
            logger.exception('Failed to send task assignment email for task id %s to %s', task.pk, addr)
