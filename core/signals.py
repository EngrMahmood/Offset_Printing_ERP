from django.db.models.signals import post_save
from django.dispatch import receiver
from django.contrib.auth import get_user_model
from django.db import IntegrityError

from .models import UserProfile

User = get_user_model()


@receiver(post_save, sender=User)
def create_user_profile(sender, instance, created, **kwargs):
    """Auto-create UserProfile when a new User is created"""
    if created:
        try:
            UserProfile.objects.get_or_create(user=instance)
        except IntegrityError:
            # Profile already exists, skip
            pass


@receiver(post_save, sender=User)
def ensure_user_profile_exists(sender, instance, **kwargs):
    """Ensure UserProfile exists for every user (handles edge cases)"""
    try:
        profile = instance.profile
    except UserProfile.DoesNotExist:
        try:
            UserProfile.objects.create(user=instance)
        except IntegrityError:
            # Race condition - another process created it
            pass


from django.db.models.signals import pre_save

@receiver(pre_save, sender='planning.PlanningJob')
def capture_previous_status(sender, instance, **kwargs):
    if instance.pk:
        try:
            old_instance = sender.objects.get(pk=instance.pk)
            instance._previous_status = old_instance.status
        except sender.DoesNotExist:
            instance._previous_status = None
    else:
        instance._previous_status = None


@receiver(post_save, sender='planning.PlanningJob')
def trigger_job_notifications(sender, instance, created, **kwargs):
    previous_status = getattr(instance, '_previous_status', None)
    current_status = instance.status

    if current_status != previous_status or (created and current_status == 'pending_qc'):
        from core.notifications import notify_event
        actor = getattr(instance, 'updated_by', None) or getattr(instance, 'created_by', None)
        
        if current_status == 'pending_qc':
            notify_event('job.pending_qc', instance=instance, actor=actor)
        elif current_status == 'qc_approved':
            notify_event('job.qc_approved', instance=instance, actor=actor)
        elif current_status == 'released':
            notify_event('job.released', instance=instance, actor=actor)


@receiver(post_save, sender='core.Dispatch')
def trigger_dispatch_notifications(sender, instance, created, **kwargs):
    if created:
        from core.notifications import notify_event
        actor = getattr(instance, 'created_by', None)
        notify_event('dispatch.created', instance=instance, actor=actor)


@receiver(post_save, sender='core.EditOverrideRequest')
def trigger_override_notifications(sender, instance, created, **kwargs):
    if created:
        from core.notifications import notify_event
        actor = getattr(instance, 'requested_by', None)
        notify_event('override.requested', instance=instance, actor=actor)


@receiver(post_save, sender='core.Production')
def trigger_production_notifications(sender, instance, created, **kwargs):
    if created:
        from core.notifications import notify_event
        actor = getattr(instance, 'created_by', None)
        notify_event('production.submitted', instance=instance, actor=actor)


@receiver(post_save, sender='core.PasswordResetRequest')
def trigger_password_reset_notifications(sender, instance, created, **kwargs):
    if created:
        from core.models import Notification
        admins = User.objects.filter(is_active=True, profile__role='admin')
        for admin in admins:
            Notification.objects.create(
                user=admin,
                event_type='password_reset_request',
                title='Password Reset Request',
                message=f"User/Email '{instance.username_or_email}' has requested a password reset.",
                link='/admin/core/passwordresetrequest/',
                entity_type='password_reset_request',
                entity_id=instance.id,
            )

