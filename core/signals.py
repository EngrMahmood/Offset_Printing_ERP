import logging

from django.db.models.signals import post_save
from django.dispatch import receiver
from django.contrib.auth import get_user_model
from django.db import IntegrityError

from .models import UserProfile

logger = logging.getLogger(__name__)

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
    
    # Auto-sync JobCard status based on dispatch completion ratio
    job_card = instance.job_card
    if job_card:
        _sync_job_card_dispatch_status(job_card)


from django.db.models.signals import post_delete

@receiver(post_delete, sender='core.Dispatch')
def sync_job_card_status_on_dispatch_delete(sender, instance, **kwargs):
    job_card = instance.job_card
    if job_card:
        _sync_job_card_dispatch_status(job_card)


def _sync_job_card_dispatch_status(job_card):
    dispatch_ratio = job_card.dispatch_completion_percent
    if dispatch_ratio >= 95:
        if job_card.status == 'in_production':
            from core.jobcard_service import transition_job_card_status
            try:
                transition_job_card_status(job_card, 'completed', reason='System: Dispatch completion reached >= 95%')
            except Exception:
                pass
    else:
        if job_card.status == 'completed':
            from core.jobcard_service import transition_job_card_status
            try:
                transition_job_card_status(job_card, 'in_production', reason='System: Dispatch completion fell below 95%')
            except Exception:
                pass


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


@receiver(post_save, sender='core.Production')
def split_merge_group_printing_entry(sender, instance, created, **kwargs):
    """Auto-split a lead job's printing run to every member SKU card.

    When a ganged (merge-group) sheet is printed, the operator records ONE
    printing entry on the lead job card. Each member SKU physically comes off
    the same sheets, so we mirror that entry onto every member's card, counting
    pieces at the SKU's ups share on the combined sheet (item.allocated_ups).
    """
    if not created or instance.entry_type != 'printing' or not instance.is_active:
        return
    if instance.merge_parent_id:  # already a derived child — never re-split
        return

    from core.models import JobCard, Production

    def _card_for(pjob):
        try:
            return pjob.job_card
        except (JobCard.DoesNotExist, AttributeError):
            return None

    job_card = instance.job_card
    planning_job = getattr(job_card, 'planning_job', None)
    if not planning_job:
        return
    group = planning_job.active_merge_group
    if not group or group.lead_job_id != planning_job.id:
        return

    items = list(group.items.select_related('planning_job'))
    lead_item = next((it for it in items if it.planning_job_id == planning_job.id), None)
    if lead_item is None:
        return

    # The lead's own entry counts at its combined-sheet ups, not its full-sheet ups.
    Production.objects.filter(pk=instance.pk).update(
        merge_allocated_ups=lead_item.allocated_ups,
    )

    for item in items:
        if item.planning_job_id == planning_job.id:
            continue
        member_card = _card_for(item.planning_job)
        if not member_card:
            continue  # member has no job card yet; nothing to attribute
        try:
            Production.objects.create(
                job_card=member_card,
                entry_type='printing',
                date=instance.date,
                shift=instance.shift,
                machine=instance.machine,
                output_sheets=instance.output_sheets,
                waste_sheets=0,          # run waste is counted once, on the lead entry
                impressions=0,           # run impressions counted once, on the lead entry
                merge_parent=instance,
                merge_allocated_ups=item.allocated_ups,
                status=instance.status,
                created_by=instance.created_by,
            )
        except Exception:
            # A single member failing must never block the operator's entry.
            logger.exception(
                'Merge split failed for %s on group %s',
                item.planning_job.jc_number, group.code,
            )


@receiver(post_save, sender='core.Production')
def sync_sku_preferred_machine_from_production(sender, instance, **kwargs):
    """Part C: learn a default machine per SKU from actual printing runs,
    so future planning can suggest it. Skipped when the SKU master has an
    explicit manual lock (machine_name_locked)."""
    if instance.entry_type != 'printing' or not instance.machine_id:
        return
    job_card = instance.job_card
    sku = (getattr(job_card, 'SKU', '') or '').strip()
    if not sku:
        return

    from planning.models import SkuRecipe
    recipe = SkuRecipe.objects.filter(sku__iexact=sku).first()
    if not recipe or recipe.machine_name_locked:
        return
    machine_name = instance.machine.name
    if recipe.machine_name != machine_name:
        recipe.machine_name = machine_name
        recipe.save(update_fields=['machine_name'])


@receiver(post_save, sender='planning.PlanningJob')
@receiver(post_save, sender='core.Machine')
def bump_reports_cache_version(sender, instance, **kwargs):
    """Invalidate cached report payloads (e.g. Machine Planning) so
    priority/master-data edits are reflected immediately instead of
    waiting for the cache timeout."""
    from reports.report_engine.engine import bump_cache_version
    bump_cache_version()


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

