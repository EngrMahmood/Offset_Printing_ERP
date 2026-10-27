"""Flexo planning workflow — the Flexo counterpart of planning/services.py and
the job-card status machine, without any offset sheet/plate/merge logic."""
import math
from decimal import Decimal

from django.core.exceptions import ValidationError
from django.db import transaction
from django.utils import timezone

from .models import (
    CANCEL_REASON_CHOICES, JOB_CANCEL_REQUEST_TYPE, FlexoJobCard, FlexoJobChangeRequest,
    FlexoJobEvent, FlexoPlanningJob, FlexoSkuRecipe,
)
from .numbering import allocate_next_flexo_jc_number

# Spec fields copied from the SKU master onto a job (and re-synced from it).
RECIPE_JOB_FIELDS = [
    'job_name', 'material', 'color_spec', 'size_w_mm', 'size_h_mm', 'ups', 'roll_size',
    'meter_per_roll', 'pcs_per_roll', 'wastage_pct', 'unit_type', 'die_code', 'teeth',
    'purchase_material_origin', 'daily_demand',
]


def user_is_admin(user):
    profile = getattr(user, 'profile', None)
    return getattr(user, 'is_superuser', False) or (
        profile is not None and getattr(profile, 'normalized_role', '') == 'admin'
    )


def log_event(job, action, *, actor=None, from_status='', to_status='', reason=''):
    return FlexoJobEvent.objects.create(
        planning_job=job, action=action, from_status=from_status, to_status=to_status,
        reason=reason or '', actor=actor,
    )


# ---------------------------------------------------------------------------
# SKU master
# ---------------------------------------------------------------------------

def apply_recipe_to_job(job, recipe):
    job.sku_recipe = recipe
    job.sku = recipe.sku
    for field in RECIPE_JOB_FIELDS:
        setattr(job, field, getattr(recipe, field))
    if recipe.machine_id:
        job.machine_name = recipe.machine
    if recipe.department_id:
        job.department = recipe.department
    if recipe.default_unit_cost is not None:
        job.unit_cost = recipe.default_unit_cost
    return job


def approved_recipe_for(sku):
    return FlexoSkuRecipe.objects.filter(
        sku__iexact=sku, is_active=True, master_data_status='approved',
    ).first()


@transaction.atomic
def create_job_from_recipe(recipe, *, actor, po_number='', po_received_date=None,
                           delivery_date=None, order_qty=0, stock_qty=0, destination='',
                           priority=1, remarks='', unit_type=None, unit_cost=None, department=None):
    if not recipe.is_approved:
        raise ValidationError(f'SKU {recipe.sku} is not approved; planning jobs need an approved SKU master.')
    if not order_qty or order_qty <= 0:
        raise ValidationError('Order qty must be greater than zero.')

    today = timezone.localdate()
    job = FlexoPlanningJob(
        jc_number=allocate_next_flexo_jc_number(),
        po_number=po_number or '',
        po_received_date=po_received_date,
        delivery_date=delivery_date,
        plan_date=today,
        plan_month=today.strftime('%m-%b'),
        order_qty=order_qty,
        stock_qty=stock_qty or 0,
        destination=destination or '',
        priority=priority,
        remarks=remarks or '',
        created_by=actor,
        repeat_flag='Repeat' if FlexoPlanningJob.objects.filter(sku_recipe=recipe).exists() else 'New',
    )
    apply_recipe_to_job(job, recipe)
    if unit_type in ('piece', 'roll'):
        job.unit_type = unit_type
    if unit_cost is not None:
        job.unit_cost = unit_cost
    if department is not None:
        job.department = department
    job.amount = (Decimal(job.unit_cost) * order_qty).quantize(Decimal('0.01')) if job.unit_cost is not None else None
    job.full_clean(exclude=['jc_number'])
    job.save()
    log_event(job, 'create', actor=actor, to_status='draft')
    return job


# ---------------------------------------------------------------------------
# Job card
# ---------------------------------------------------------------------------

def ensure_job_card(job):
    card = getattr(job, 'job_card', None)
    if card is None:
        card = FlexoJobCard(planning_job=job, job_card_no=job.jc_number)
    card.sku = job.sku
    card.job_name = job.job_name
    card.po_number = job.po_number
    card.po_received_date = job.po_received_date
    card.order_qty = job.order_qty
    card.unit_type = job.unit_type
    card.delivery_location = job.destination
    card.department = job.department
    card.daily_demand = job.daily_demand
    card.material_type = job.material
    card.size_w_mm = job.size_w_mm
    card.size_h_mm = job.size_h_mm
    card.print_roll_size = job.roll_size
    card.roll_meter_qty = job.print_meters
    card.ups = job.ups
    card.machine = job.machine_name
    card.wastage = job.wastage_meters
    card.color = job.color_spec
    card.mtr_per_roll = '' if job.meter_per_roll is None else format(job.meter_per_roll.normalize(), 'f')
    if job.daily_demand and job.daily_demand > 0:
        card.days_plan = math.ceil(Decimal(job.order_qty) / job.daily_demand)
    card.die_code = job.die_code
    card.is_active = True
    card.status = job.status
    card.save()
    return card


def _sync_card_status(job, status):
    card = getattr(job, 'job_card', None)
    if card:
        card.status = status
        card.save(update_fields=['status', 'updated_at'])


# ---------------------------------------------------------------------------
# Status workflow: draft -> pending_qc -> qc_approved -> released
# ---------------------------------------------------------------------------

def get_qc_submission_blockers(job):
    blockers = []
    recipe = job.sku_recipe
    if recipe is None:
        blockers.append(f'SKU master for {job.sku or "this job"} is missing.')
    elif not recipe.is_approved:
        blockers.append(f'SKU master for {job.sku} is not approved; QC submission is blocked until approval.')
    if not job.order_qty:
        blockers.append('Order qty (pcs) is required.')
    if not job.po_received_date:
        blockers.append('PO received date is required.')
    if not job.ups:
        blockers.append('Ups is required.')
    if job.machine_name_id is None:
        blockers.append('Machine is required.')
    if job.print_meters in (None, 0):
        blockers.append('Print meters could not be calculated; check ups, meter per roll, pcs per roll and order qty.')
    return blockers


def _transition(job, action, to_status, *, actor, reason='', allowed_from=()):
    current = job.status
    if allowed_from and current not in allowed_from:
        raise ValidationError(f'Transition not allowed from {current} to {to_status}.')
    with transaction.atomic():
        job.status = to_status
        job.issued_to_production = to_status in {'released', 'in_production'}
        update = ['status', 'issued_to_production', 'updated_at']
        if action in {'reject_qc', 'reopen'}:
            job.rejection_reason = reason if action == 'reject_qc' else ''
            update.append('rejection_reason')
        job.save(update_fields=update)
        if to_status != 'draft' or getattr(job, 'job_card', None) is not None:
            ensure_job_card(job)
        _sync_card_status(job, to_status)
        log_event(job, action, actor=actor, from_status=current, to_status=to_status, reason=reason)
    return job


def submit_to_qc(job, *, actor):
    blockers = get_qc_submission_blockers(job)
    if blockers:
        raise ValidationError(blockers)
    return _transition(job, 'submit_qc', 'pending_qc', actor=actor, allowed_from=('draft',))


def approve_qc(job, *, actor, reason=''):
    return _transition(job, 'approve_qc', 'qc_approved', actor=actor, reason=reason, allowed_from=('pending_qc',))


def reject_qc(job, *, actor, reason):
    if not (reason or '').strip():
        raise ValidationError('Rejection reason is required.')
    return _transition(job, 'reject_qc', 'draft', actor=actor, reason=reason, allowed_from=('pending_qc',))


def release_for_production(job, *, actor, reason=''):
    return _transition(job, 'release', 'released', actor=actor, reason=reason, allowed_from=('qc_approved',))


def reopen_to_draft(job, *, actor, reason=''):
    return _transition(job, 'reopen', 'draft', actor=actor, reason=reason,
                       allowed_from=('qc_approved', 'released', 'in_production', 'completed'))


# ---------------------------------------------------------------------------
# Hold / archive / cancel
# ---------------------------------------------------------------------------

def hold_job(job, *, actor, reason):
    if not (reason or '').strip():
        raise ValidationError('A hold reason is required to place a job on hold.')
    job.is_on_hold, job.hold_reason, job.hold_by, job.hold_at = True, reason.strip(), actor, timezone.now()
    job.save(update_fields=['is_on_hold', 'hold_reason', 'hold_by', 'hold_at', 'updated_at'])
    log_event(job, 'hold', actor=actor, reason=reason)


def release_hold(job, *, actor):
    job.is_on_hold, job.hold_reason, job.hold_by, job.hold_at = False, '', None, None
    job.save(update_fields=['is_on_hold', 'hold_reason', 'hold_by', 'hold_at', 'updated_at'])
    log_event(job, 'release_hold', actor=actor)


def archive_job(job, *, actor, reason=''):
    job.is_active = False
    job.archive_reason, job.archived_by, job.archived_at = reason or '', actor, timezone.now()
    job.save(update_fields=['is_active', 'archive_reason', 'archived_by', 'archived_at', 'updated_at'])
    log_event(job, 'archive', actor=actor, reason=reason)


def restore_job(job, *, actor, reason=''):
    if job.is_cancelled:
        raise ValidationError('Cancelled jobs cannot be restored.')
    job.is_active = True
    job.restored_by, job.restored_at, job.restore_reason = actor, timezone.now(), reason or ''
    job.save(update_fields=['is_active', 'restored_by', 'restored_at', 'restore_reason', 'updated_at'])
    log_event(job, 'restore', actor=actor, reason=reason)


def _validate_cancel_input(reason, reason_code):
    reason = (reason or '').strip()
    if not reason:
        raise ValidationError({'reason': 'A cancellation reason is required.'})
    if reason_code not in {code for code, _ in CANCEL_REASON_CHOICES}:
        raise ValidationError({'reason_code': 'Select a valid cancellation reason code.'})
    return reason


def cancel_job(job, *, actor, reason, reason_code):
    reason = _validate_cancel_input(reason, reason_code)
    blockers = job.cancellation_blockers()
    if blockers:
        raise ValidationError(blockers)
    now = timezone.now()
    previous = job.status
    with transaction.atomic():
        job.is_cancelled, job.cancel_reason, job.cancel_reason_code = True, reason, reason_code
        job.cancelled_by, job.cancelled_at, job.status = actor, now, 'cancelled'
        job.is_active = False
        job.archive_reason, job.archived_by, job.archived_at = f'Cancelled: {reason}', actor, now
        job.save(update_fields=[
            'is_cancelled', 'cancel_reason', 'cancel_reason_code', 'cancelled_by', 'cancelled_at',
            'status', 'is_active', 'archive_reason', 'archived_by', 'archived_at', 'updated_at',
        ])
        card = getattr(job, 'job_card', None)
        if card:
            card.is_active = False
            card.status = 'cancelled'
            card.save(update_fields=['is_active', 'status', 'updated_at'])
        log_event(job, 'cancel', actor=actor, from_status=previous, to_status='cancelled', reason=reason)
    return job


def request_cancellation(job, *, actor, reason, reason_code):
    reason = _validate_cancel_input(reason, reason_code)
    if job.is_cancelled:
        raise ValidationError('This job is already cancelled.')
    if FlexoJobChangeRequest.objects.filter(
        planning_job=job, request_type=JOB_CANCEL_REQUEST_TYPE, status='pending',
    ).exists():
        raise ValidationError('A cancellation request is already pending for this job.')
    return FlexoJobChangeRequest.objects.create(
        planning_job=job, request_type=JOB_CANCEL_REQUEST_TYPE,
        reason=f'[{reason_code}] {reason}', requested_by=actor,
    )


# ---------------------------------------------------------------------------
# Change requests (reopen / cancel a locked job — Production Manager approves)
# ---------------------------------------------------------------------------

def request_reopen(job, *, actor, reason):
    if job.status not in {'qc_approved', 'released', 'in_production'}:
        raise ValidationError('Reopen requests apply to QC approved, released or in production jobs only.')
    if not (reason or '').strip():
        raise ValidationError('A justification reason is required to submit a reopen request.')
    if job.change_requests.filter(status='pending').exists():
        raise ValidationError('There is already a pending request for this job.')
    return FlexoJobChangeRequest.objects.create(
        planning_job=job, request_type='reopen_to_draft', reason=reason.strip(), requested_by=actor,
    )


@transaction.atomic
def approve_change_request(change_request, *, actor):
    if change_request.status != 'pending':
        raise ValidationError('This request is already resolved.')
    job = change_request.planning_job
    if change_request.is_cancellation:
        cancel_job(job, actor=actor, reason=change_request.cancel_reason_text,
                   reason_code=change_request.cancel_reason_code)
    else:
        reopen_to_draft(job, actor=actor, reason=change_request.reason)
    change_request.status = 'approved'
    change_request.approved_by = actor
    change_request.approved_at = timezone.now()
    change_request.save(update_fields=['status', 'approved_by', 'approved_at'])
    return change_request


def reject_change_request(change_request, *, actor, reason=''):
    if change_request.status != 'pending':
        raise ValidationError('This request is already resolved.')
    change_request.status = 'rejected'
    change_request.approved_by = actor
    change_request.approved_at = timezone.now()
    change_request.rejection_reason = reason or ''
    change_request.save(update_fields=['status', 'approved_by', 'approved_at', 'rejection_reason'])
    return change_request


def error_messages(exc):
    """Flatten a ValidationError into a list of strings for messages.error()."""
    if hasattr(exc, 'message_dict'):
        return [m for msgs in exc.message_dict.values() for m in msgs]
    return list(exc.messages)
