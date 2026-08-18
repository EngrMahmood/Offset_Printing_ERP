"""Fires the assignment email exactly when a Task's assignee/team actually
changes — mirrors core/signals.py::capture_previous_status + trigger_job_notifications
on PlanningJob, the codebase's established pattern for model-transition side effects."""
from django.db.models.signals import pre_save, post_save
from django.dispatch import receiver

from .models import Task


@receiver(pre_save, sender=Task)
def capture_previous_assignment(sender, instance, **kwargs):
    if instance.pk:
        try:
            old = Task.objects.get(pk=instance.pk)
            instance._previous_assignee_id = old.assignee_id
            instance._previous_assigned_team_id = old.assigned_team_id
        except Task.DoesNotExist:
            instance._previous_assignee_id = None
            instance._previous_assigned_team_id = None
    else:
        instance._previous_assignee_id = None
        instance._previous_assigned_team_id = None


@receiver(post_save, sender=Task)
def send_assignment_email_on_change(sender, instance, created, **kwargs):
    prev_assignee = getattr(instance, '_previous_assignee_id', None)
    prev_team = getattr(instance, '_previous_assigned_team_id', None)
    assignee_changed = instance.assignee_id != prev_assignee
    team_changed = instance.assigned_team_id != prev_team

    if (created or assignee_changed or team_changed) and (instance.assignee_id or instance.assigned_team_id):
        from tasks.emails import send_assignment_email
        send_assignment_email(instance)
