from datetime import timedelta

from django.db import transaction
from django.utils import timezone

from supply_chain.models import ItemRequest

from .models import MaintenanceActivityLog, MaintenanceApproval, MaintenanceRecord, PreventiveMaintenancePlan


def generate_record_no(record):
    """Assign a per-year sequential record number, e.g. MNT-2026-0042."""
    year = timezone.now().year
    prefix = f'MNT-{year}-'
    with transaction.atomic():
        last = (
            MaintenanceRecord.objects.select_for_update()
            .filter(record_no__startswith=prefix)
            .order_by('-record_no')
            .first()
        )
        next_seq = 1
        if last and last.record_no:
            try:
                next_seq = int(last.record_no.rsplit('-', 1)[-1]) + 1
            except ValueError:
                next_seq = 1
        record.record_no = f'{prefix}{next_seq:04d}'
        record.save(update_fields=['record_no'])
    return record.record_no


def log_activity(record, actor, action, from_status='', to_status='', note=''):
    MaintenanceActivityLog.objects.create(
        record=record, actor=actor, action=action,
        from_status=from_status, to_status=to_status, note=note,
    )


def raise_spare_part_demand(spare_part, user):
    from supply_chain.item_request_service import generate_request_no

    record = spare_part.record
    item_request = ItemRequest.objects.create(
        request_type=_maintenance_request_type(),
        request_date=timezone.now(),
        item_title=spare_part.description,
        machine=record.machine,
        uom=spare_part.uom,
        specifications=f'Spare part for maintenance record {record.record_no}',
        required_quantity=spare_part.quantity,
        department=_maintenance_department(),
        existing_sku=spare_part.existing_sku,
        raised_by=user,
    )
    generate_request_no(item_request)
    spare_part.item_request = item_request
    spare_part.save(update_fields=['item_request'])
    log_activity(record, user, 'Raised spare part demand', note=item_request.request_no)
    return item_request


def raise_service_demand(service_job, user):
    from supply_chain.item_request_service import generate_request_no
    from supply_chain.models import ItemRequestType

    record = service_job.record
    item_request = ItemRequest.objects.create(
        request_type=ItemRequestType.objects.filter(code='SRV', is_active=True).first(),
        request_date=timezone.now(),
        item_title=f'Outsourced repair — {record.machine}',
        machine=record.machine,
        specifications=service_job.scope,
        required_quantity=1,
        department=_maintenance_department(),
        raised_by=user,
    )
    generate_request_no(item_request)
    service_job.item_request = item_request
    service_job.save(update_fields=['item_request'])
    log_activity(record, user, 'Raised service demand', note=item_request.request_no)
    return item_request


def _maintenance_request_type():
    from supply_chain.models import ItemRequestType

    return ItemRequestType.objects.filter(code='MNT', is_active=True).first()


def _maintenance_department():
    from supply_chain.models import ItemRequestDepartment

    dept, _ = ItemRequestDepartment.objects.get_or_create(name='Maintenance')
    return dept


def submit_for_approval(record, user):
    MaintenanceApproval.objects.create(record=record, actor=user, action='SUBMIT')
    log_activity(record, user, 'Submitted for approval', to_status=record.status)


def review_record(record, user, action, comment=''):
    """Manager approves or rejects a PENDING_APPROVAL maintenance record."""
    if record.status != 'PENDING_APPROVAL':
        raise ValueError('Only records pending approval can be reviewed.')

    from_status = record.status
    if action == 'APPROVE':
        record.status = 'REPORTED'
    elif action == 'REJECT':
        record.status = 'REJECTED'
    else:
        raise ValueError('Invalid review action.')

    record.save(update_fields=['status'])
    MaintenanceApproval.objects.create(record=record, actor=user, action=action, comment=comment)
    log_activity(record, user, f'Record {action.lower()}d', from_status=from_status, to_status=record.status, note=comment)
    return record


def soft_delete_record(record, user, reason=''):
    record.is_active = False
    record.deleted_by = user
    record.deleted_at = timezone.now()
    record.save(update_fields=['is_active', 'deleted_by', 'deleted_at'])
    MaintenanceApproval.objects.create(record=record, actor=user, action='DELETE', comment=reason)
    log_activity(record, user, 'Record deleted', note=reason)


def generate_due_pm_records(as_of=None, actor=None):
    """Create MaintenanceRecords for any active PM plan whose next_due_at has passed (DAYS interval only)."""
    as_of = as_of or timezone.now().date()
    created = []
    plans = PreventiveMaintenancePlan.objects.filter(
        is_active=True, interval_type='DAYS', next_due_at__isnull=False, next_due_at__lte=as_of,
    )
    for plan in plans:
        record = MaintenanceRecord.objects.create(
            machine=plan.machine,
            reported_date=as_of,
            reported_by=actor,
            maintenance_type='PREVENTIVE',
            priority='MEDIUM',
            fault_description=f'Preventive maintenance due: {plan.title}',
            status='PENDING_APPROVAL',
        )
        generate_record_no(record)
        log_activity(record, actor, 'Auto-generated from PM plan', to_status=record.status, note=plan.title)
        plan.last_done_at = as_of
        plan.next_due_at = as_of + timedelta(days=plan.interval_value)
        plan.save(update_fields=['last_done_at', 'next_due_at'])
        created.append(record)
    return created
