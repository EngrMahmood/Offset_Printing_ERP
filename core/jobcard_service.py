from __future__ import annotations

import re

from django.core.exceptions import ValidationError
from django.db import transaction

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
        'planning': {'draft', 'pending_data', 'qc_rejected'},
        'qc': {'planning_approved'},
        'production_manager': {'qc_approved'},
        'production': {'production_approved'},
    }
    return queue_map.get(queue_name, set())


def job_card_queue_queryset(queue_name):
    statuses = job_card_queue_statuses(queue_name)
    queryset = JobCard.objects.filter(is_active=True)
    if statuses:
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

    return model_class.objects.filter(name__iexact=raw_value).first()


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
            'po_date': planning_job.po_received_date,
            'PO_No': planning_job.po_number,
            'SKU': planning_job.sku,
            'material': material,
            'colour': planning_job.color_spec,
            'application': planning_job.application,
            'order_qty': int(planning_job.order_qty or 0),
            'total_impressions_required': int(planning_job.calculated_sheets_required or 0),
            'ups': int(planning_job.ups or 0) or None,
            'print_sheet_size': planning_job.print_sheet_size or '',
            'plate_set_no': planning_job.plate_set_no or '',
            'wastage': int(planning_job.wastage_sheets or 0),
            'total_sheet_quantity': planning_job.calculated_sheets_required,
            'purchase_sheet_size': planning_job.purchase_sheet_size or '',
            'purchase_sheet_ups': int(planning_job.purchase_sheet_ups or 0) or None,
            'remarks': planning_job.remarks or '',
            'destination': planning_job.destination or '',
            'machine_name': machine,
            'department': department,
            'die_cutting': planning_job.die_cutting or '',
            'total_colors': planning_job.number_of_colors,
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


def transition_job_card_status(job_card: JobCard, target_status, actor=None, reason=''):
    target_status = normalize_job_card_status(target_status, default='')
    if not target_status:
        raise ValidationError({'status': 'Unknown Job Card workflow status.'})

    current_status = job_card.workflow_status
    if current_status == target_status:
        return job_card

    allowed_transitions = {
        ('draft', 'planning_approved'),
        ('pending_data', 'planning_approved'),
        ('qc_rejected', 'planning_approved'),
        ('planning_approved', 'pending_qc'),
        ('planning_approved', 'qc_approved'),
        ('planning_approved', 'qc_rejected'),
        ('pending_qc', 'qc_approved'),
        ('pending_qc', 'qc_rejected'),
        ('qc_approved', 'pending_pm_approval'),
        ('qc_approved', 'production_approved'),
        ('qc_approved', 'pm_rejected'),
        ('pending_pm_approval', 'production_approved'),
        ('pending_pm_approval', 'pm_rejected'),
        ('production_approved', 'released'),
        ('released', 'in_production'),
        ('in_production', 'completed'),
        ('completed', 'closed'),
        ('closed', 'draft'),
        ('pm_rejected', 'qc_approved'),
    }
    if (current_status, target_status) not in allowed_transitions:
        raise ValidationError({'status': f'Transition not allowed from {current_status} to {target_status}.'})

    if target_status in JOB_CARD_PLANNING_APPROVAL_STATUSES and job_card.planning_missing_fields():
        missing_fields = ', '.join(job_card.planning_missing_fields())
        raise ValidationError({'status': f'Complete planning fields before moving to {target_status}: {missing_fields}.'})

    with transaction.atomic():
        before_status = current_status
        job_card.status = target_status
        job_card.save(update_fields=['status'])
        log_job_card_workflow_change(
            job_card,
            actor,
            STATUS_ACTION_MAP.get(target_status, 'update'),
            reason=reason,
            before_status=before_status,
            extra_message=reason,
        )

    return job_card


def execute_job_card_action(job_card, action, actor=None, reason=''):
    mapping = {
        'approve_planning': 'planning_approved',
        'reject_planning': 'qc_rejected',
        'approve_qc': 'qc_approved',
        'reject_qc': 'qc_rejected',
        'approve_pm': 'production_approved',
        'reject_pm': 'pm_rejected',
        'release_for_production': 'released',
    }
    if action not in mapping:
        raise ValueError('Unknown approval transition.')
    return transition_job_card_status(job_card, mapping[action], actor=actor, reason=reason)


def submit_to_qc(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'approve_planning', actor=actor, reason=reason)


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


def start_production(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'start_production', actor=actor, reason=reason)


def complete_production(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'complete', actor=actor, reason=reason)


def close_job_card(job_card, actor=None, reason=''):
    return execute_job_card_action(job_card, 'close', actor=actor, reason=reason)
