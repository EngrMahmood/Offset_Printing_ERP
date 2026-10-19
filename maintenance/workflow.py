"""The single authoritative workflow for maintenance tickets.

Every status change — from a view, a command or a test — goes through
``transition_record``. Views and templates never encode their own rules.

Ticket status, machine downtime and machine availability are deliberately three
separate things:

* A ticket reaching COMPLETED means the technician says the work is done. It
  does *not* mean the machine is back in production.
* Downtime for a stopped machine stays open until a repair is VERIFIED (a passed
  test run) or an authorised user confirms the machine is restored when
  cancelling. It is never closed merely because a ticket changed state.
* If another unresolved ticket on the same machine is also stopping it, the
  downtime interval is handed over to that ticket instead of being closed.
* ``core.Machine.is_under_maintenance`` (operational colours) is a planning
  concept maintained elsewhere; nothing here redefines it.
"""
import logging
from decimal import Decimal

from django.db import transaction
from django.utils import timezone

from .models import MaintenanceRecord
from .services import (
    log_activity, notify_repair_completed, notify_status_update, notify_waiting, release_downtime,
)

logger = logging.getLogger(__name__)

TRANSITIONS = {
    'REPORTED': {'DIAGNOSED', 'CANCELLED'},
    'DIAGNOSED': {'AWAITING_PARTS', 'AWAITING_VENDOR', 'IN_PROGRESS', 'CANCELLED'},
    'AWAITING_PARTS': {'IN_PROGRESS', 'CANCELLED'},
    'AWAITING_VENDOR': {'IN_PROGRESS', 'CANCELLED'},
    'IN_PROGRESS': {'AWAITING_PARTS', 'AWAITING_VENDOR', 'COMPLETED', 'CANCELLED'},
    'COMPLETED': {'VERIFIED', 'IN_PROGRESS'},   # IN_PROGRESS = test failed, back to repair
    'VERIFIED': {'CLOSED'},
    'CLOSED': {'IN_PROGRESS'},                  # reopen — restricted, needs a reason
}
# DIAGNOSED is reached only through the triage screen (it captures the
# classification), never through the generic status control.
TRIAGE_ONLY = {'DIAGNOSED'}
REOPEN_ROLES = ('admin', 'manager')

# Plain-English button labels for the shop floor and engineers.
ACTION_LABELS = {
    'DIAGNOSED': 'Diagnose',
    'AWAITING_PARTS': 'Waiting for Parts',
    'AWAITING_VENDOR': 'Waiting for Vendor',
    'IN_PROGRESS': 'Start / Resume Repair',
    'COMPLETED': 'Repair Completed',
    'VERIFIED': 'Test Passed — Verify',
    'CLOSED': 'Close Ticket',
    'CANCELLED': 'Cancel Ticket',
}
STATUS_HELP = {
    'REPORTED': 'Reported — waiting for maintenance to look at it.',
    'DIAGNOSED': 'Assessed — a repair plan is set.',
    'AWAITING_PARTS': 'Waiting for spare parts.',
    'AWAITING_VENDOR': 'Waiting for an outside vendor.',
    'IN_PROGRESS': 'A technician is repairing it.',
    'COMPLETED': 'Repair finished — needs a test run and verification.',
    'VERIFIED': 'Repair tested and verified.',
    'CLOSED': 'Closed.',
    'CANCELLED': 'Cancelled.',
}

MACHINE_STATE_CHOICES = [
    ('restored', 'The machine is running again (end the downtime)'),
    ('still_down', 'The machine is still down — another active ticket covers it'),
]


class TransitionError(ValueError):
    """A rejected transition. The message is safe to show to users."""


def allowed_transitions(record, user=None):
    """Statuses the record may move to next (optionally filtered for ``user``)."""
    options = set(TRANSITIONS.get(record.status, set())) - TRIAGE_ONLY
    if 'IN_PROGRESS' in options and record.status == 'CLOSED':
        role = _role_of(user)
        if user is not None and not (getattr(user, 'is_superuser', False) or role in REOPEN_ROLES):
            options.discard('IN_PROGRESS')
    return sorted(options)


def _role_of(user):
    profile = getattr(user, 'profile', None)
    return (getattr(profile, 'role', '') or '').strip().lower()


def open_stopped_downtime(record):
    return record.downtime_intervals.filter(end_at__isnull=True).first()


def _other_stopped_tickets(record):
    return MaintenanceRecord.objects.filter(
        machine_id=record.machine_id, is_active=True, machine_stopped=True,
        status__in=MaintenanceRecord.OPEN_STATUSES,
    ).exclude(pk=record.pk)


def transition_record(record, to_status, actor, note='', labour_hours=None, machine_state=''):
    """Move ``record`` to ``to_status`` or raise ``TransitionError`` with no changes made."""
    note = (note or '').strip()
    with transaction.atomic():
        record = MaintenanceRecord.objects.select_for_update().get(pk=record.pk)
        from_status = record.status

        if to_status in TRIAGE_ONLY:
            raise TransitionError('Use the Triage screen to diagnose a complaint.')
        if to_status not in TRANSITIONS.get(from_status, set()):
            raise TransitionError(
                f'This ticket is {record.get_status_display()} and cannot move to '
                f'{dict(MaintenanceRecord.STATUS_CHOICES).get(to_status, to_status)}.'
            )

        is_reopen = from_status == 'CLOSED'
        if is_reopen:
            if not (getattr(actor, 'is_superuser', False) or _role_of(actor) in REOPEN_ROLES):
                raise TransitionError('Only a manager can reopen a closed ticket.')
            if not note:
                raise TransitionError('A reason is required to reopen a closed ticket.')

        if to_status == 'CANCELLED':
            _prepare_cancellation(record, note, machine_state)

        if to_status == 'COMPLETED':
            _validate_completion(record, labour_hours)

        if to_status == 'IN_PROGRESS' and from_status == 'COMPLETED' and not note:
            raise TransitionError('Say why the test failed or the repair is being resumed.')

        if to_status == 'VERIFIED' and record.test_result != 'PASS':
            raise TransitionError('Record a passed machine test before verifying the repair.')

        now = timezone.now()
        record.status = to_status
        if to_status == 'IN_PROGRESS':
            if not record.work_start_at:
                record.work_start_at = now
            if from_status in ('COMPLETED', 'CLOSED'):
                record.work_end_at = None
                record.verified_by = None
                record.verified_at = None
        if to_status == 'COMPLETED':
            record.work_end_at = now
            if labour_hours not in (None, ''):
                record.labour_hours = Decimal(str(labour_hours))
        if to_status == 'VERIFIED':
            record.verified_by = actor
            record.verified_at = now
        record.save()

        if to_status == 'VERIFIED':
            if record.pm_plan_id:
                record.pm_plan.last_done_at = timezone.localdate()
                record.pm_plan.save(update_fields=['last_done_at'])
            release_downtime(record, actor, 'Repair verified — machine tested and restored')
        elif to_status == 'CLOSED':
            release_downtime(record, actor, 'Ticket closed')
        elif to_status == 'CANCELLED':
            if machine_state == 'restored' or not record.machine_stopped:
                release_downtime(record, actor, f'Cancelled — machine confirmed restored. {note}'.strip())
            else:
                release_downtime(record, actor, 'Cancelled — downtime continued by another ticket', handover_only=True)

        log_activity(record, actor, 'Status changed', from_status=from_status, to_status=to_status, note=note)

        transaction.on_commit(lambda: _notify_after_commit(record, actor, to_status))
    return record


def _prepare_cancellation(record, note, machine_state):
    if not note:
        raise TransitionError('A reason is required to cancel a ticket.')
    if record.machine_stopped and open_stopped_downtime(record):
        if machine_state not in ('restored', 'still_down'):
            raise TransitionError(
                'This ticket recorded the machine as stopped. Confirm whether the machine is running again '
                'or is still down before cancelling.'
            )
        if machine_state == 'still_down' and not _other_stopped_tickets(record).exists():
            raise TransitionError(
                'The machine is still down but no other active ticket covers it. Raise or keep a ticket open '
                'for the fault first, or confirm the machine is running.'
            )


def _validate_completion(record, labour_hours):
    missing = []
    # A scheduled preventive job has no fault to diagnose.
    if record.maintenance_type != 'PREVENTIVE' and not (record.root_cause or '').strip():
        missing.append('diagnosis / root cause')
    if not (record.repair_action or '').strip():
        missing.append('repair action performed')
    if missing:
        raise TransitionError('Fill in the ' + ' and '.join(missing) + ' before completing the repair.')
    if record.checklist_items.filter(is_done=False).exists():
        raise TransitionError('Tick off every item on the PM checklist before completing the job.')
    if labour_hours not in (None, ''):
        try:
            value = Decimal(str(labour_hours))
        except (TypeError, ValueError, ArithmeticError):
            raise TransitionError('Labour hours must be a number.')
        if value < 0:
            raise TransitionError('Labour hours cannot be negative.')


def _notify_after_commit(record, actor, to_status):
    """Notifications must never undo or break a committed business change."""
    try:
        if to_status == 'COMPLETED':
            notify_repair_completed(record, actor)
        if to_status in ('AWAITING_PARTS', 'AWAITING_VENDOR'):
            notify_waiting(record, actor)
        if to_status in ('COMPLETED', 'VERIFIED', 'CLOSED'):
            notify_status_update(record, actor)
    except Exception:
        logger.exception('Maintenance notification failed for record %s (%s)', record.pk, to_status)
