from __future__ import annotations

import re

from django.core.exceptions import ValidationError
from django.db import transaction
from django.db.models import Q
from django.utils import timezone

from planning.models import PLANNING_STAGE_DONE

from .models import (
    ChangeLog,
    Department,
    JOB_CARD_PLANNING_APPROVAL_STATUSES,
    JOB_CARD_PLANNING_EDITABLE_STATUSES,
    JOB_CARD_STATUS_ALIASES,
    JOB_CARD_STATUS_CHOICES,
    JobCard,
    Machine,
    Material,
)

TRANSITION_ACTIONS = {
    'submit_to_qc': ('planning_approved', 'submit'),
    'enqueue_qc': ('pending_qc', 'submit'),
    'approve_qc': ('qc_approved', 'approve'),
    'reject_qc': ('qc_rejected', 'reject'),
    'enqueue_pm': ('pending_pm_approval', 'submit'),
    'approve_pm': ('production_approved', 'approve'),
    'reject_pm': ('pm_rejected', 'reject'),
    'release': ('released', 'release'),
    'start_production': ('in_production', 'start_production'),
    'complete': ('completed', 'complete'),
    'close': ('closed', 'close'),
    'reopen': ('draft', 'update'),
}

STATUS_ACTION_MAP = {
    target_status: action_name
    for action_name, (target_status, action_name) in TRANSITION_ACTIONS.items()
}


def normalize_job_card_status(raw_value, default='pending_data'):
    value = (raw_value or '').strip().lower()
    value = JOB_CARD_STATUS_ALIASES.get(value, value)
    status_values = {choice_value for choice_value, _ in JOB_CARD_STATUS_CHOICES}
    if value in status_values:
        return value
    return default


def job_card_queue_statuses(queue_name):
    queue_map = {
        'planning': {'draft', 'pending_data'},
        'qc': {'planning_approved', 'pending_qc'},
        'production_manager': {'qc_approved'},
        'production': {'production_approved'},
    }
    return queue_map.get(queue_name, set())


def job_card_queue_queryset(queue_name):
    statuses = job_card_queue_statuses(queue_name)
    queryset = JobCard.objects.filter(is_active=True)
    if statuses:
        if queue_name == 'qc':
            # Keep QC visibility resilient when PlanningJob already moved to pending_qc
            # but JobCard status sync has not yet been applied.
            queryset = queryset.filter(
                Q(status__in=statuses) | Q(planning_job__status='pending_qc')
            )
        else:
            queryset = queryset.filter(status__in=statuses)
    return queryset.select_related('planning_job', 'material', 'machine_name', 'department', 'created_by')


def _resolve_by_name(model_class, raw_value):
    raw_value = (raw_value or '').strip()
    if not raw_value:
        return None

    normalized_value = re.sub(r'[\s_-]+', ' ', raw_value).strip()
    exact_match = model_class.objects.filter(name__iexact=normalized_value).first()
    if exact_match:
        return exact_match

    startswith_matches = model_class.objects.filter(name__istartswith=normalized_value)
    if startswith_matches.count() == 1:
        return startswith_matches.first()

    contains_matches = model_class.objects.filter(name__icontains=normalized_value)
    if contains_matches.count() == 1:
        return contains_matches.first()

    return model_class.objects.filter(name__iexact=raw_value).first()


def resolve_total_impressions_required(planning_job):
    """Return job-card impression target; includes pass count when configured."""
    impressions = planning_job.planned_total_impressions
    if impressions is None:
        impressions = planning_job.calculated_planned_total_impressions
    if impressions is not None:
        return int(impressions)
    return int(planning_job.calculated_sheets_required or 0)


def ensure_job_card_from_planning_job(planning_job, actor=None):
    """Create or refresh the linked JobCard from a PlanningJob source record."""
    with transaction.atomic():
        try:
            job_card = planning_job.job_card
            created = False
        except JobCard.DoesNotExist:
            job_card = None
            created = True

        material = _resolve_by_name(Material, getattr(planning_job, 'material', ''))
        machine = _resolve_by_name(Machine, getattr(planning_job, 'machine_name', ''))
        department = _resolve_by_name(Department, getattr(planning_job, 'department', ''))

        defaults = {
            'planning_job': planning_job,
            'job_card_no': planning_job.jc_number,
            'month': planning_job.plan_date.strftime('%B') if planning_job.plan_date else (planning_job.plan_month or ''),
            'po_date': planning_job.po_approval_date or planning_job.po_received_date,
            'PO_No': planning_job.po_number,
            'SKU': planning_job.sku,
            'material': material,
            'colour': (planning_job.color_spec or '').strip(),
            'application': planning_job.application,
            'order_qty': int(planning_job.order_qty or 0),
            'total_impressions_required': resolve_total_impressions_required(planning_job),
            'ups': planning_job.ups or None,
            'print_sheet_size': planning_job.print_sheet_size or '',
            'plate_set_no': planning_job.plate_set_no or '',
            'wastage': int(planning_job.wastage_sheets or 0),
            'total_sheet_quantity': planning_job.calculated_sheets_required,
            'purchase_sheet_size': planning_job.purchase_sheet_size or '',
            'purchase_sheet_ups': planning_job.purchase_sheet_ups or None,
            'remarks': planning_job.remarks_display or '',
            'destination': planning_job.destination or '',
            'machine_name': machine,
            'department': department,
            'die_cutting': getattr(planning_job, 'die_cutting_display', '') or '',
            'total_colors': planning_job.number_of_colors or 0,
            'is_print_job': not planning_job.is_cut_and_pack(),
            'created_by': planning_job.created_by or actor,
            'status': 'pending_data',
        }

        if job_card is None:
            job_card = JobCard.objects.create(**defaults)
            return job_card, True

        if job_card.workflow_status in JOB_CARD_PLANNING_EDITABLE_STATUSES:
            for field_name, value in defaults.items():
                setattr(job_card, field_name, value)
            if not job_card.created_by and defaults['created_by']:
                job_card.created_by = defaults['created_by']
            job_card.save()
        elif not job_card.planning_job_id:
            job_card.planning_job = planning_job
            job_card.save(update_fields=['planning_job'])

        return job_card, created


def ensure_job_card_from_pending_qc_planning_job(planning_job, actor=None):
    job_card, created = ensure_job_card_from_planning_job(planning_job, actor=actor)
    if planning_job.workflow_status == 'pending_qc' and job_card.workflow_status in JOB_CARD_PLANNING_EDITABLE_STATUSES:
        submit_to_qc(job_card, actor=actor, reason='Backfilled QC routing from PlanningJob save')
    return job_card, created


def log_job_card_workflow_change(job_card, actor, action, reason='', before_status='', extra_message=''):
    field_changes = {
        'status': {
            'label': 'Status',
            'from': before_status or '-',
            'to': job_card.workflow_status,
        }
    }
    if extra_message:
        field_changes['note'] = {
            'label': 'Note',
            'from': '-',
            'to': extra_message,
        }

    ChangeLog.objects.create(
        entity_type='job_card',
        record_id=job_card.pk,
        record_label=str(job_card),
        action=action,
        changed_by=actor,
        change_reason=reason,
        field_changes=field_changes,
    )


def job_card_completion_blockers(job_card):
    """Reasons a job card cannot become 'completed' yet. Empty list = clear
    to proceed. Applied to both the automatic 95%-dispatch completion signal
    (core.signals._sync_job_card_dispatch_status) and the manual Close tool
    (core.job_card_finalization) via transition_job_card_status, so a job can
    never reach 'completed'/'closed' — auto or forced — without production
    data behind it.

    A repeat job fulfilled entirely from carried-forward stock
    (PlanningJob.stock_qty) can legitimately have zero *new* packing entries
    this cycle, so packed pcs and stock qty are counted together against
    what's actually been dispatched, not checked as a bare "packed > 0". The
    same is true of printing: a job dispatched entirely out of stock never
    needed a print run of its own, so "no printing logged" is only a real
    blocker while packed + stock still falls short of what shipped.
    """
    reasons = []

    stock_qty = job_card.planning_job.stock_qty if job_card.planning_job_id else 0
    covered = job_card.total_packed_pcs + (stock_qty or 0)

    if job_card.is_print_job and job_card.total_printed_pcs <= 0 and covered < job_card.total_dispatch:
        reasons.append('No printing entries have been logged for this job.')

    if covered <= 0:
        reasons.append('No packing entries have been logged for this job.')
    elif covered < job_card.total_dispatch:
        reasons.append(
            f'Packed + stock ({covered}) is less than dispatched ({job_card.total_dispatch}) — '
            'packing data looks incomplete for the amount already shipped.'
        )
    return reasons


def transition_job_card_status(job_card: JobCard, target_status, actor=None, reason=''):
    target_status = normalize_job_card_status(target_status, default='')
    if not target_status:
        raise ValidationError({'status': 'Unknown Job Card workflow status.'})

    planning_status_map = {
        'draft': ('draft', False),
        'planning_approved': ('pending_qc', False),
        'pending_qc': ('pending_qc', False),
        'qc_rejected': ('draft', False),
        'qc_approved': ('qc_approved', False),
        'pending_pm_approval': ('qc_approved', False),
        'pm_rejected': ('draft', False),
        'production_approved': ('qc_approved', False),
        'released': ('released', True),
        'in_production': ('in_production', True),
        'completed': ('completed', True),
        'closed': ('completed', True),
    }

    def _sync_linked_planning_job(status_value, transition_actor=None):
        if not job_card.planning_job_id or status_value not in planning_status_map:
            return False
        job = job_card.planning_job
        planning_status, issued_to_production = planning_status_map[status_value]
        updates = []
        if job.status != planning_status:
            job.status = planning_status
            updates.append('status')
        if getattr(job, 'issued_to_production', False) != issued_to_production:
            job.issued_to_production = issued_to_production
            updates.append('issued_to_production')
        if status_value == 'released' and job.planning_stage != PLANNING_STAGE_DONE:
            job.planning_stage = PLANNING_STAGE_DONE
            job.planning_stage_changed_at = timezone.now()
            if transition_actor:
                job.planning_stage_changed_by = transition_actor
                updates.append('planning_stage_changed_by')
            updates.extend(['planning_stage', 'planning_stage_changed_at'])
        if not updates:
            return False
        updates.append('updated_at')
        job.save(update_fields=updates)
        return True

    current_status = job_card.workflow_status
    if current_status == target_status:
        _sync_linked_planning_job(target_status, transition_actor=actor)
        return job_card

    if target_status == 'released':
        from printing_plates.services import validate_job_card_release_allowed

        validate_job_card_release_allowed(job_card)

    allowed_transitions = {
        ('draft', 'planning_approved'),
        ('pending_data', 'planning_approved'),
        ('pending_data', 'draft'),
        ('qc_rejected', 'planning_approved'),
        ('qc_rejected', 'draft'),
        ('planning_approved', 'pending_qc'),
        ('planning_approved', 'qc_approved'),
        ('planning_approved', 'qc_rejected'),
        ('planning_approved', 'draft'),
        ('pending_qc', 'qc_approved'),
        ('pending_qc', 'qc_rejected'),
        ('pending_qc', 'draft'),
        ('qc_approved', 'pending_pm_approval'),
        ('qc_approved', 'production_approved'),
        ('qc_approved', 'pm_rejected'),
        ('qc_approved', 'draft'),
        ('pending_pm_approval', 'production_approved'),
        ('pending_pm_approval', 'pm_rejected'),
        ('pending_pm_approval', 'draft'),
        ('production_approved', 'draft'),
        ('pm_rejected', 'draft'),
        ('production_approved', 'released'),
        ('released', 'in_production'),
        ('released', 'draft'),
        ('in_production', 'completed'),
        ('in_production', 'draft'),
        ('completed', 'closed'),
        ('closed', 'draft'),
        ('pm_rejected', 'qc_approved'),
        # Admin/manager "Reopen to Production" — resumes normal dispatch-%
        # auto-tracking (core.signals._sync_job_card_dispatch_status), unlike
        # ('closed', 'draft') above which resets the whole workflow.
        ('completed', 'in_production'),
        ('closed', 'in_production'),
    }
    if (current_status, target_status) not in allowed_transitions:
        raise ValidationError({'status': f'Transition not allowed from {current_status} to {target_status}.'})

    if target_status in JOB_CARD_PLANNING_APPROVAL_STATUSES and job_card.planning_missing_fields():
        missing_fields = ', '.join(job_card.planning_missing_fields())
        raise ValidationError({'status': f'Complete planning fields before moving to {target_status}: {missing_fields}.'})

    if target_status == 'completed':
        blockers = job_card_completion_blockers(job_card)
        if blockers:
            raise ValidationError({'status': '; '.join(blockers)})

    with transaction.atomic():
        before_status = current_status
        job_card.status = target_status
        job_card.save(update_fields=['status'])
        _sync_linked_planning_job(target_status, transition_actor=actor)
        log_job_card_workflow_change(
            job_card,
            actor,
            STATUS_ACTION_MAP.get(target_status, 'update'),
            reason=reason,
            before_status=before_status,
            extra_message=reason,
        )

        # Reopening a job that already has production/dispatch recorded is a
        # deliberate, permission-gated feature (planners can correct a
        # mistake even mid-production) — but the reset itself is easy to
        # miss unless someone happens to open this specific job's history.
        # Proactively tell admin/manager whenever it happens on a job that
        # already has real activity, regardless of which of the two reopen
        # paths (JobCardChangeRequest approval, or "Reopen & Apply Master
        # Sync") triggered it — both funnel through this one function.
        if (
            target_status == 'draft'
            and before_status in {'released', 'in_production', 'completed', 'closed'}
            and (job_card.total_printed_pcs > 0 or job_card.total_packed_pcs > 0)
        ):
            from django.urls import reverse

            from core.notifications import notify_roles

            try:
                notify_roles(
                    ['admin', 'manager'],
                    event_type='job_card.reopened_with_production_history',
                    title=f'{job_card.job_card_no} reopened to Draft after production had started',
                    message=(
                        f'{job_card.job_card_no} was at "{before_status}" with '
                        f'{job_card.total_printed_pcs} pcs printed, {job_card.total_packed_pcs} pcs packed, '
                        f'{job_card.total_dispatch} pcs dispatched, then reset to Draft. '
                        f'Reason: {reason or "(none given)"}.'
                    ),
                    link=reverse('planning:job_detail', args=[job_card.planning_job_id]) if job_card.planning_job_id else '',
                    entity_type='job_card',
                    entity_id=job_card.pk,
                    actor=actor,
                )
            except Exception:  # noqa: BLE001 - a notification failure must not block the reopen
                pass

        # A job card that already has printing/packing entries recorded
        # (e.g. it was reopened for a data correction — machine/pass count/
        # etc — after production had genuinely started, then walked back
        # through the full approval pipeline to 'released') would otherwise
        # sit invisible to Dispatch Entry: JOB_CARD_DISPATCHABLE_STATUSES
        # excludes 'released', despite the job having a real remaining
        # balance. Cascade straight through to 'in_production', matching
        # what already happens on a normal first-time release once a
        # Production record is saved (production/views.py, packing_entry.py
        # both call start_production() there) — here it's the reverse case,
        # where the production record already existed before 'released' was
        # (re-)reached.
        if target_status == 'released' and (job_card.total_printed_pcs > 0 or job_card.total_packed_pcs > 0):
            transition_job_card_status(
                job_card,
                'in_production',
                actor=actor,
                reason='System: production already recorded before this release — resuming In Production',
            )

    return job_card


def close_job_card_manually(job_card, actor, reason):
    """Force a job card to 'closed', walking through 'completed' first if
    needed, and record the never-dispatched remainder against the existing
    (previously unused) short_close_* fields as confirmed wastage.

    Used by the admin/manager finalization tool for job cards that will
    never hit the 95%-dispatch auto-complete threshold (a permanent
    wastage/shortfall keeps dispatch under 95% of order_qty forever — see
    core.signals._sync_job_card_dispatch_status), and for job cards that
    already auto-completed but were never closed (nothing in this app
    closes a job card automatically today). Once 'closed', the auto-signal
    never touches it again, so no extra guard flag is needed to make this
    stick.
    """
    if job_card.workflow_status == 'in_production':
        transition_job_card_status(job_card, 'completed', actor=actor, reason=reason)
    if job_card.workflow_status == 'completed':
        gap = max(job_card.order_qty - job_card.total_dispatch, 0)
        if gap:
            job_card.short_close_closed_qty = gap
            job_card.short_close_wastage_qty = gap
            job_card.short_close_closed_by = actor
            job_card.short_close_closed_at = timezone.now()
            job_card.short_close_close_reason = reason
            job_card.save(update_fields=[
                'short_close_closed_qty', 'short_close_wastage_qty',
                'short_close_closed_by', 'short_close_closed_at', 'short_close_close_reason',
            ])
        transition_job_card_status(job_card, 'closed', actor=actor, reason=reason)
    return job_card


def reopen_job_card_manually(job_card, actor, reason):
    """Reopen a manually-closed/completed job card back to 'in_production',
    clearing any short-close record set by close_job_card_manually() — the
    gap it confirmed as wastage may no longer be accurate once more
    production/dispatch activity can be logged against the job again."""
    transition_job_card_status(job_card, 'in_production', actor=actor, reason=reason)
    if job_card.short_close_closed_at:
        job_card.short_close_closed_qty = 0
        job_card.short_close_wastage_qty = 0
        job_card.short_close_closed_by = None
        job_card.short_close_closed_at = None
        job_card.short_close_close_reason = None
        job_card.save(update_fields=[
            'short_close_closed_qty', 'short_close_wastage_qty',
            'short_close_closed_by', 'short_close_closed_at', 'short_close_close_reason',
        ])
    return job_card


def reopen_job_card_for_master_sync(job_card, actor=None, reason=''):
    """Reopen a locked job card to draft so master sync can refresh sheet fields."""
    if job_card.workflow_status in JOB_CARD_PLANNING_EDITABLE_STATUSES:
        return job_card

    allowed_locked_statuses = {
        'released',
        'in_production',
        'qc_approved',
        'production_approved',
        'pending_pm_approval',
        'planning_approved',
        'pending_qc',
    }
    if job_card.workflow_status not in allowed_locked_statuses:
        raise ValidationError({
            'status': f'Job Card cannot be reopened for master sync from {job_card.workflow_status}.',
        })

    return transition_job_card_status(
        job_card,
        'draft',
        actor=actor,
        reason=reason or 'Reopened for SKU master sync',
    )


def execute_job_card_action(job_card, action, actor=None, reason=''):
    mapping = {
        'approve_planning': 'planning_approved',
        'reject_planning': 'draft',
        'approve_qc': 'qc_approved',
        'reject_qc': 'qc_rejected',
        'approve_pm': 'production_approved',
        'reject_pm': 'pm_rejected',
        'release_for_production': 'released',
        'start_production': 'in_production',
        'reopen': 'draft',
    }
    if action not in mapping:
        raise ValueError('Unknown approval transition.')
    return transition_job_card_status(job_card, mapping[action], actor=actor, reason=reason)


def enqueue_qc(job_card, actor=None, reason=''):
    return transition_job_card_status(job_card, 'pending_qc', actor=actor, reason=reason)


def submit_to_qc(job_card, actor=None, reason=''):
    if job_card.workflow_status == 'pending_qc':
        return job_card

    if job_card.workflow_status == 'planning_approved':
        return enqueue_qc(job_card, actor=actor, reason=reason)

    if job_card.workflow_status in {'draft', 'pending_data', 'qc_rejected'}:
        job_card = execute_job_card_action(job_card, 'approve_planning', actor=actor, reason=reason)
        if job_card.workflow_status == 'planning_approved':
            job_card = enqueue_qc(job_card, actor=actor, reason=reason)
        return job_card

    return job_card


def reject_planning(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'reject_planning', actor=actor, reason=reason)


def approve_qc(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'approve_qc', actor=actor, reason=reason)


def reject_qc(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'reject_qc', actor=actor, reason=reason)


def approve_pm(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'approve_pm', actor=actor, reason=reason)


def release_to_production(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'release_for_production', actor=actor, reason=reason)


def approve_card_for_merged_run(job_card, group, actor=None):
    """Production-approve a member card for a combined (smart-merge) run.

    The group's own layout approval stands in for this SKU's individual QC/PM
    gate, so this sets the card straight to 'production_approved' rather than
    walking the normal per-SKU transitions (which a draft member could never
    satisfy). It stops short of 'released' on purpose: the lead must still raise
    the combined plate (released jobs skip plate making), and the whole group is
    released together when those plates are received. Idempotent: a card already
    at or past production_approved is left as-is.
    """
    if job_card.workflow_status in {'production_approved', 'released', 'in_production', 'completed', 'closed'}:
        return job_card

    reason = f'Production-approved for combined layout {group.code} (group approval)'
    with transaction.atomic():
        before_status = job_card.workflow_status
        job_card.status = 'production_approved'
        job_card.save(update_fields=['status'])

        planning_job = job_card.planning_job
        if planning_job and planning_job.status != 'qc_approved':
            planning_job.status = 'qc_approved'
            planning_job.save(update_fields=['status', 'updated_at'])

        log_job_card_workflow_change(
            job_card, actor, 'approve_pm',
            reason=reason, before_status=before_status, extra_message=reason,
        )
    return job_card


def start_production(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'start_production', actor=actor, reason=reason)


def complete_production(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'complete', actor=actor, reason=reason)


def close_job_card(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'close', actor=actor, reason=reason)
