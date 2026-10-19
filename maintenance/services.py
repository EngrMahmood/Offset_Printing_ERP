import logging
from datetime import date, timedelta

from django.db import transaction
from django.urls import reverse
from django.utils import timezone

from core.notifications import notify_roles, notify_users
from supply_chain.models import ItemRequest

from .models import (
    MachineDowntime, MaintenanceActivityLog, MaintenanceApproval, MaintenanceChecklistItem, MaintenanceRecord,
    PreventiveMaintenancePlan,
)

logger = logging.getLogger(__name__)

ENGINEER_ROLES = ('maintenance_engineer', 'production_manager')


def _detail_link(record):
    try:
        return reverse('maintenance:record_detail', args=[record.pk])
    except Exception:
        return ''


def notify_complaint_raised(record, actor):
    notify_roles(
        ENGINEER_ROLES,
        event_type='maintenance.complaint_raised',
        title=f'New complaint: {record.record_no} — {record.machine}',
        message=f'{actor.username} reported a fault on {record.machine}: {record.fault_description}',
        link=_detail_link(record),
        entity_type='maintenancerecord',
        entity_id=record.pk,
        actor=actor,
    )


def notify_triaged(record, actor):
    if not record.reported_by_id:
        return
    notify_users(
        [record.reported_by],
        event_type='maintenance.triaged',
        title=f'{record.record_no} picked up by {actor.username}',
        message=f'Your complaint on {record.machine} is being assessed by maintenance.',
        link=_detail_link(record),
        entity_type='maintenancerecord',
        entity_id=record.pk,
        actor=actor,
    )


def notify_status_update(record, actor):
    if not record.reported_by_id:
        return
    notify_users(
        [record.reported_by],
        event_type='maintenance.status_update',
        title=f'{record.record_no} — {record.get_status_display()}',
        message=f'Your reported issue on {record.machine} is now {record.get_status_display()}.',
        link=_detail_link(record),
        entity_type='maintenancerecord',
        entity_id=record.pk,
        actor=actor,
    )


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


def open_downtime(record, reason='BREAKDOWN'):
    """Open downtime for the record's machine — one interval per physical stoppage.

    If the machine already has an open interval (from this or another ticket) it
    is reused rather than duplicated, so the same stoppage is never counted twice.
    """
    existing = MachineDowntime.objects.filter(machine_id=record.machine_id, end_at__isnull=True).first()
    if existing:
        if existing.record_id is None:
            existing.record = record
            existing.save(update_fields=['record'])
        return None
    return MachineDowntime.objects.create(
        machine=record.machine, record=record, start_at=timezone.now(), reason=reason,
    )


def release_downtime(record, actor, reason, handover_only=False):
    """End this ticket's open downtime — unless another stopped ticket still needs it.

    If another unresolved, stopped ticket exists on the machine the interval is
    handed to it and stays open (the machine is not restored). Otherwise it is
    closed, recording who closed it and why. ``handover_only`` never closes.
    """
    intervals = list(record.downtime_intervals.filter(end_at__isnull=True))
    if not intervals:
        return
    successor = (
        MaintenanceRecord.objects.filter(
            machine_id=record.machine_id, is_active=True, machine_stopped=True,
            status__in=MaintenanceRecord.OPEN_STATUSES,
        ).exclude(pk=record.pk).order_by('created_at').first()
    )
    for interval in intervals:
        if successor:
            interval.record = successor
            interval.save(update_fields=['record'])
            log_activity(record, actor, 'Downtime handed over', note=f'Machine still stopped by {successor.record_no}')
            log_activity(successor, actor, 'Downtime taken over', note=f'From {record.record_no}')
        elif not handover_only:
            interval.end_at = timezone.now()
            interval.closed_by = actor
            interval.closed_reason = (reason or '')[:255]
            interval.save(update_fields=['end_at', 'closed_by', 'closed_reason'])
            log_activity(record, actor, 'Downtime ended', note=reason)


def notify_repair_completed(record, actor):
    """Ask the engineers to test and verify the finished repair."""
    notify_roles(
        ENGINEER_ROLES,
        event_type='maintenance.verification_needed',
        title=f'Verify repair: {record.record_no} — {record.machine}',
        message=f'{record.record_no} is reported complete. Run the machine test and verify the repair.',
        link=_detail_link(record),
        entity_type='maintenancerecord',
        entity_id=record.pk,
        actor=actor,
    )


def notify_waiting(record, actor):
    if not record.reported_by_id:
        return
    notify_users(
        [record.reported_by],
        event_type='maintenance.waiting',
        title=f'{record.record_no} — {record.get_status_display()}',
        message=f'The repair on {record.machine} is paused: {record.get_status_display().lower()}.',
        link=_detail_link(record),
        entity_type='maintenancerecord',
        entity_id=record.pk,
        actor=actor,
    )


def notify_assigned(record, actor):
    if not record.assigned_to_id or record.assigned_to_id == getattr(actor, 'pk', None):
        return
    notify_users(
        [record.assigned_to],
        event_type='maintenance.assigned',
        title=f'Assigned to you: {record.record_no} — {record.machine}',
        message=f'{record.record_no} ({record.get_priority_display()}) was assigned to you.',
        link=_detail_link(record),
        entity_type='maintenancerecord',
        entity_id=record.pk,
        actor=actor,
    )


def open_breakdown_downtime_if_needed(record):
    """Called at triage once the engineer classifies a record as BREAKDOWN."""
    if record.maintenance_type != 'BREAKDOWN':
        return None
    if not record.machine_stopped:
        record.machine_stopped = True
        record.save(update_fields=['machine_stopped'])
    return open_downtime(record, reason='BREAKDOWN')


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
    release_downtime(record, user, 'Ticket deleted')


def machine_impression_total(machine):
    """Cumulative impressions run on ``machine``, summed from Production entries.

    The ERP keeps no dedicated machine counter, so this is the best available
    source: every active production entry logged against the machine. It is only
    as complete as the shop floor's production entries - impressions run but never
    logged are invisible - which is why impression plans are shown with their
    source and why a plan with no baseline is never guessed at.
    """
    from django.db.models import Sum

    from core.models import Production

    total = Production.objects.filter(machine=machine, is_active=True).aggregate(total=Sum('impressions'))['total']
    return int(total or 0)


def pm_overdue_since(record):
    """Date a still-open PM ticket became due, or None (not a PM ticket / not date based)."""
    if not record.pm_plan_id or not record.pm_due_key or record.pm_due_key.startswith('imp:'):
        return None
    try:
        return date.fromisoformat(record.pm_due_key)
    except ValueError:
        return None


def _create_pm_record(plan, actor, as_of, due_key, reason):
    record = MaintenanceRecord.objects.create(
        machine=plan.machine,
        reported_date=as_of,
        reported_by=actor,
        maintenance_type='PREVENTIVE',
        priority='MEDIUM',
        fault_description=f'Preventive maintenance due: {plan.title}',
        proposed_solution=plan.instructions,
        # Already fully classified (a scheduled PM job) - straight to DIAGNOSED.
        status='DIAGNOSED',
        assigned_to=plan.assigned_to,
        pm_plan=plan,
        pm_due_key=due_key,
    )
    generate_record_no(record)
    items = [line.strip() for line in (plan.checklist or '').splitlines() if line.strip()]
    MaintenanceChecklistItem.objects.bulk_create([
        MaintenanceChecklistItem(record=record, position=index, text=text[:255]) for index, text in enumerate(items)
    ])
    log_activity(record, actor, 'Auto-generated from PM plan', to_status=record.status, note=f'{plan.title} - {reason}')
    return record


def _advance_days(plan, as_of):
    """Anchored schedule: step from the due date, never from the moment the job ran."""
    step = timedelta(days=max(plan.interval_value, 1))
    next_due = plan.next_due_at
    while next_due <= as_of:
        next_due += step
    return next_due


def generate_due_pm_records(as_of=None, actor=None):
    """Raise one ticket per due occurrence of every active PM plan. Safe to run repeatedly.

    Returns ``{'created': [...], 'skipped': int, 'failed': [(plan, message)]}``.
    A database constraint on (plan, due occurrence) makes a duplicate impossible
    even if two runs overlap; one broken plan never stops the others.
    """
    as_of = as_of or timezone.localdate()
    result = {'created': [], 'skipped': 0, 'failed': []}
    plans = PreventiveMaintenancePlan.objects.filter(is_active=True).select_related('machine', 'assigned_to')
    for plan in plans:
        try:
            with transaction.atomic():
                record = _process_plan(plan, actor, as_of)
        except Exception as exc:
            logger.exception('PM generation failed for plan %s', plan.pk)
            result['failed'].append((plan, str(exc)))
            continue
        if record is None:
            result['skipped'] += 1
        else:
            result['created'].append(record)
    for record in result['created']:
        try:
            notify_pm_due(record, actor)
        except Exception:
            logger.exception('PM-due notification failed for record %s', record.pk)
    return result


def _process_plan(plan, actor, as_of):
    if plan.interval_type == 'DAYS':
        if not plan.next_due_at or plan.next_due_at > as_of:
            return None
        due_key = plan.next_due_at.isoformat()
        already = MaintenanceRecord.objects.filter(pm_plan=plan, pm_due_key=due_key).exists()
        record = None if already else _create_pm_record(plan, actor, as_of, due_key, f'due {due_key}')
        plan.next_due_at = _advance_days(plan, as_of)
        plan.save(update_fields=['next_due_at'])
        return record

    if plan.interval_type == 'IMPRESSIONS':
        total = machine_impression_total(plan.machine)
        if plan.next_due_impressions is None:
            # No baseline yet: start counting from the machine's current total
            # rather than inventing one.
            plan.next_due_impressions = total + max(plan.interval_value, 1)
            plan.save(update_fields=['next_due_impressions'])
            return None
        if total < plan.next_due_impressions:
            return None
        due_key = f'imp:{plan.next_due_impressions}'
        already = MaintenanceRecord.objects.filter(pm_plan=plan, pm_due_key=due_key).exists()
        record = None if already else _create_pm_record(
            plan, actor, as_of, due_key, f'{total:,} impressions run (due at {plan.next_due_impressions:,})',
        )
        step = max(plan.interval_value, 1)
        while plan.next_due_impressions <= total:
            plan.next_due_impressions += step
        plan.save(update_fields=['next_due_impressions'])
        return record
    return None


def notify_pm_due(record, actor):
    recipients = [record.assigned_to] if record.assigned_to_id else []
    if recipients:
        notify_users(
            recipients,
            event_type='maintenance.pm_due',
            title=f'PM due: {record.record_no} - {record.machine}',
            message=f'Preventive maintenance "{record.pm_plan.title}" is due on {record.machine}.',
            link=_detail_link(record),
            entity_type='maintenancerecord',
            entity_id=record.pk,
            actor=actor,
        )
    notify_roles(
        ENGINEER_ROLES,
        event_type='maintenance.pm_due',
        title=f'PM due: {record.record_no} - {record.machine}',
        message=f'Preventive maintenance "{record.pm_plan.title}" is due on {record.machine}.',
        link=_detail_link(record),
        entity_type='maintenancerecord',
        entity_id=record.pk,
        actor=actor,
    )


OVERDUE_REMINDER_EVERY = timedelta(days=7)


def remind_overdue_pm(as_of=None, actor=None):
    """Nudge engineers about date-based PM tickets still open past their due date.

    One reminder per ticket per week, remembered in the activity log, so running
    this every few minutes never repeats it.
    """
    as_of = as_of or timezone.localdate()
    now = timezone.now()
    sent = []
    open_pm = MaintenanceRecord.objects.filter(
        is_active=True, pm_plan__isnull=False, status__in=MaintenanceRecord.OPEN_STATUSES,
    ).select_related('machine', 'pm_plan', 'assigned_to')
    for record in open_pm:
        due = pm_overdue_since(record)
        if not due or due >= as_of:
            continue
        last = record.activity_log.filter(action='Overdue reminder').order_by('-created_at').first()
        if last and now - last.created_at < OVERDUE_REMINDER_EVERY:
            continue
        days = (as_of - due).days
        recipients = [record.assigned_to] if record.assigned_to_id else []
        try:
            if recipients:
                notify_users(
                    recipients, event_type='maintenance.pm_overdue',
                    title=f'PM overdue: {record.record_no} - {record.machine}',
                    message=f'"{record.pm_plan.title}" was due {days} day(s) ago.',
                    link=_detail_link(record), entity_type='maintenancerecord', entity_id=record.pk, actor=actor,
                )
            notify_roles(
                ENGINEER_ROLES, event_type='maintenance.pm_overdue',
                title=f'PM overdue: {record.record_no} - {record.machine}',
                message=f'"{record.pm_plan.title}" was due {days} day(s) ago.',
                link=_detail_link(record), entity_type='maintenancerecord', entity_id=record.pk, actor=actor,
            )
        except Exception:
            logger.exception('PM-overdue notification failed for record %s', record.pk)
            continue
        log_activity(record, actor, 'Overdue reminder', note=f'{days} day(s) overdue')
        sent.append(record)
    return sent


def pm_plan_status(plan, today=None, impression_total=None):
    """Plain-language state of a plan for the PM screen: overdue / due soon / upcoming."""
    today = today or timezone.localdate()
    if not plan.is_active:
        return 'inactive', 'Disabled'
    if plan.interval_type == 'DAYS':
        if not plan.next_due_at:
            return 'unset', 'No due date set'
        delta = (plan.next_due_at - today).days
        if delta < 0:
            return 'overdue', f'Overdue by {-delta} day(s)'
        if delta <= 7:
            return 'soon', 'Due today' if delta == 0 else f'Due in {delta} day(s)'
        return 'upcoming', f'Due in {delta} days'
    if plan.next_due_impressions is None or impression_total is None:
        return 'unset', 'Waiting for first baseline'
    left = plan.next_due_impressions - impression_total
    if left <= 0:
        return 'overdue', 'Impression limit reached'
    if left <= max(plan.interval_value // 10, 1):
        return 'soon', f'{left:,} impressions to go'
    return 'upcoming', f'{left:,} impressions to go'
