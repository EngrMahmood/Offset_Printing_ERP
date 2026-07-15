from django.utils import timezone
from core.models import JobCard, ProductionWipStatus, JobCardWipStatus, ChangeLog, Production, Dispatch
from django.db.models import Sum

def get_or_create_wip_status(name, user=None):
    """Safely get or create a WIP status by name."""
    status, created = ProductionWipStatus.objects.get_or_create(
        name=name,
        defaults={'created_by': user}
    )
    return status

def get_system_calculated_status_name(job_card):
    """
    Calculates the system WIP status strictly from logged data (production/dispatch entries).
    """
    # 1. Dispatch check
    dispatches = Dispatch.objects.filter(job_card=job_card, is_active=True)
    total_dispatched = dispatches.aggregate(total=Sum('dispatch_qty'))['total'] or 0
    if total_dispatched >= job_card.order_qty and job_card.order_qty > 0:
        return 'Completed'
    if total_dispatched > 0:
        return 'Partial Dispatch'

    # 2. Packing check
    packing_records = Production.objects.filter(job_card=job_card, is_active=True, entry_type='packing')
    total_packed = packing_records.aggregate(total=Sum('packing_qty'))['total'] or 0
    if total_packed >= job_card.order_qty and job_card.order_qty > 0:
        return 'Ready for Dispatch'
    if packing_records.exists():
        return 'Sorting / Packing'

    # 3. Printing check
    printing_records = Production.objects.filter(job_card=job_card, is_active=True, entry_type='printing')
    
    from production.printing_pass_helpers import get_job_card_pass_count
    total_passes = get_job_card_pass_count(job_card)
    final_pass_exists = printing_records.filter(print_pass_number=total_passes, output_sheets__gt=0).exists()
    
    if final_pass_exists:
        return 'Printing Completed'

    if printing_records.exists() or job_card.workflow_status == 'released':
        return 'Printing'

    return 'Not Set'

def update_wip_status_for_job(job_card, target_status_name, user=None, is_manual=False, force=False):
    """
    Updates the WIP status of a Job Card.
    If is_manual is False (auto-transition) and the job already has a manual override,
    the update is skipped unless force=True.
    """
    wip_status = get_or_create_wip_status(target_status_name, user=user)
    
    # Check existing status
    existing = JobCardWipStatus.objects.filter(job_card=job_card).first()
    
    if existing:
        # If it was manually set, skip auto updates unless forced
        if existing.is_manual and not is_manual and not force:
            return False
            
        old_status_name = existing.status.name
        if old_status_name == target_status_name:
            # Already set to this status
            if existing.is_manual != is_manual:
                existing.is_manual = is_manual
                existing.updated_by = user
                existing.save(update_fields=['is_manual', 'updated_by'])
            return False
            
        existing.status = wip_status
        existing.is_manual = is_manual
        existing.updated_by = user
        existing.save()
    else:
        old_status_name = 'Not Set'
        existing = JobCardWipStatus.objects.create(
            job_card=job_card,
            status=wip_status,
            is_manual=is_manual,
            updated_by=user
        )
        
    # Log the transition in ChangeLog
    ChangeLog.objects.create(
        entity_type='job_card',
        record_id=job_card.pk,
        record_label=str(job_card),
        action='update',
        changed_by=user,
        change_reason=f"WIP Status updated to '{target_status_name}' ({'Manual Override' if is_manual else 'Auto-transition'})",
        field_changes={
            'wip_status': {
                'label': 'WIP Status',
                'from': old_status_name,
                'to': target_status_name
            }
        }
    )
    return True

def evaluate_and_update_job_wip_status(job_card, user=None, force=False):
    """
    Evaluates the current operational state of a Job Card and updates its WIP status
    in the database if not manually overridden by a supervisor.
    """
    target_status = get_system_calculated_status_name(job_card)
    if target_status == 'Not Set':
        return False
    return update_wip_status_for_job(job_card, target_status, user=user, is_manual=False, force=force)
