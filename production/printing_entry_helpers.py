"""Shared helpers for printing production entry and job card search."""

from __future__ import annotations

import math
import re

from django.db.models import Prefetch, Q, Sum

from core.machine_routing import color_class
from core.models import JOB_CARD_PRODUCTION_CONTINUE_STATUSES, JobCard, Machine, Production
from core.services import compute_planned_minutes
from production.printing_pass_helpers import (
    build_pass_tracking_info,
    effective_print_pass_number,
    get_job_card_pass_count,
    get_max_print_passes,
)
from printing_plates.services import job_is_waiting_for_plates


def get_degraded_machine_pass_hint(job_card, machine, planned_passes):
    """Suggest a pass-count override when the assigned machine is running at
    reduced colour capacity (a colour unit under maintenance).

    Returns ``None`` when the machine is at full capacity or the job's colours
    fit the working units. The supervisor always confirms the final number —
    this is only a hint, never applied automatically.
    """
    if machine is None:
        return None
    default_colors = int(getattr(machine, 'default_colors', 0) or 0)
    effective = int(getattr(machine, 'effective_colors', 0) or 0)
    if default_colors <= 0 or effective <= 0 or effective >= default_colors:
        return None
    colors_per_pass = int(color_class(job_card.colour) or 0)
    if colors_per_pass <= effective:
        return None
    factor = math.ceil(colors_per_pass / effective)
    suggested = min(get_max_print_passes(), max(int(planned_passes or 1), int(planned_passes or 1) * factor))
    return {
        'machine_name': machine.name,
        'default_colors': default_colors,
        'effective_colors': effective,
        'colors_per_pass': colors_per_pass,
        'suggested_passes': suggested,
    }


def resolve_related_machine(job_card):
    if job_card.machine_name_id:
        return job_card.machine_name

    display_name = (job_card.machine_name_display or '').strip()
    if not display_name:
        return None

    for lookup in ('iexact', 'istartswith', 'icontains'):
        machine = Machine.objects.filter(**{f'name__{lookup}': display_name}).first()
        if machine:
            return machine

    normalized = re.sub(r'[^A-Za-z0-9 ]+', ' ', display_name).strip()
    if normalized and normalized != display_name:
        machine = Machine.objects.filter(name__icontains=normalized).first()
        if machine:
            return machine
    return None


def get_effective_job_card_plan(job_card):
    planned_run = float(job_card.estimated_run_time_minutes or 0)
    planned_setup = float(job_card.estimated_setup_time_minutes or 0)
    planned_total = float(job_card.estimated_total_time_minutes or 0)
    machine_obj = resolve_related_machine(job_card)

    if planned_total <= 0 and job_card.total_impressions_required and machine_obj:
        fallback_run, fallback_setup, fallback_total = compute_planned_minutes(
            job_card.total_impressions_required,
            machine_obj,
            job_card.colour,
        )
        planned_run = float(planned_run or fallback_run or 0)
        planned_setup = float(planned_setup or fallback_setup or 0)
        planned_total = float(planned_total or fallback_total or 0)

    return {
        'machine_obj': machine_obj,
        'planned_total': planned_total,
        'planned_run': planned_run,
        'planned_setup': planned_setup,
    }


def get_remaining_planned_for_job_card(job_card, planned_total, exclude_production_id=None):
    if planned_total <= 0:
        return 0
    allocated_qs = job_card.productions.filter(is_active=True, entry_type='printing')
    if exclude_production_id:
        allocated_qs = allocated_qs.exclude(pk=exclude_production_id)
    allocated = float(allocated_qs.aggregate(total=Sum('planned_time'))['total'] or 0)
    return max(planned_total - allocated, 0)


def _printing_job_cards_related():
    """select_related/prefetch_related shared by both queryset branches below —
    eliminates the per-job-card FK queries (planning_job/material/machine_name)
    and lets JobCard.total_printed_pcs/total_production reuse a prefetched
    `productions` instead of hitting the DB again in build_printing_job_card_maps."""
    return dict(
        select_related=('planning_job', 'material', 'machine_name'),
        prefetch_related=(
            Prefetch('productions', queryset=Production.objects.filter(is_active=True)),
        ),
    )


def printing_job_cards_queryset(edit_record=None):
    related = _printing_job_cards_related()
    if edit_record:
        qs = JobCard.objects.filter(is_active=True, is_print_job=True).filter(
            Q(status__in=JOB_CARD_PRODUCTION_CONTINUE_STATUSES) | Q(pk=edit_record.job_card_id)
        ).distinct().order_by('-created_at')
    else:
        qs = JobCard.objects.filter(
            is_active=True,
            is_print_job=True,
            status__in=JOB_CARD_PRODUCTION_CONTINUE_STATUSES,
        ).order_by('-created_at')
    return qs.select_related(*related['select_related']).prefetch_related(*related['prefetch_related'])


def build_printing_job_card_maps(job_cards, edit_record=None):
    exclude_id = edit_record.pk if edit_record else None
    job_cards = list(job_cards)
    plan_map = {}
    machine_map = {}
    info_map = {}

    # Batched once for the whole page instead of one query per job card — see
    # get_remaining_planned_for_job_card(), whose exact "planned_total <= 0 -> 0,
    # else max(planned_total - allocated, 0)" logic is replicated below.
    allocated_qs = Production.objects.filter(
        job_card_id__in=[jc.id for jc in job_cards],
        is_active=True,
        entry_type='printing',
    )
    if exclude_id:
        allocated_qs = allocated_qs.exclude(pk=exclude_id)
    allocated_map = {
        row['job_card_id']: float(row['total'] or 0)
        for row in allocated_qs.values('job_card_id').annotate(total=Sum('planned_time'))
    }

    for job_card in job_cards:
        plan = get_effective_job_card_plan(job_card)
        planned_total = plan['planned_total']
        if planned_total <= 0:
            remaining_planned = 0
        else:
            remaining_planned = max(planned_total - allocated_map.get(job_card.id, 0), 0)
        job_id = str(job_card.id)
        plan_map[job_id] = {
            'planned_total': plan['planned_total'],
            'planned_setup': plan['planned_setup'],
            'planned_run': plan['planned_run'],
            'remaining_planned': remaining_planned,
        }
        resolved_machine = plan['machine_obj']
        machine_map[job_id] = {
            'machine_id': str(resolved_machine.id) if resolved_machine else (job_card.machine_name_id or ''),
            'mapped_machine_name': resolved_machine.name if resolved_machine else '',
            'job_card_machine_name': job_card.machine_name_display or '',
        }

        pass_count = get_job_card_pass_count(job_card)
        pass_tracking = build_pass_tracking_info(job_card, exclude_production_id=exclude_id)
        total_impressions_used = pass_tracking['total_impressions_used']
        allowed_impressions = pass_tracking['total_impressions_allowed']
        remaining_impressions = pass_tracking['total_impressions_remaining']

        history_qs = job_card.productions.filter(is_active=True, entry_type='printing').order_by('-date', '-created_at')[:8]
        history_data = []
        for row in history_qs:
            pass_no = effective_print_pass_number(row, pass_count)
            history_data.append({
                'date': row.date.strftime('%d-%b') if row.date else '',
                'shift': row.shift,
                'pass_label': f'Pass {pass_no}' + (' (final)' if pass_no >= pass_count else ''),
                'impressions': f'{row.impressions:,}',
                'output': f'{row.output_sheets:,}',
                'waste': f'{row.waste_sheets:,}',
                'runtime': row.run_time,
                'make_ready': f'{float(row.make_ready_time or 0):g}',
                'downtime': f'{float(row.downtime_minutes or 0):g}',
                'status': row.get_status_display(),
            })

        info_map[job_id] = {
            'job_card_no': job_card.job_card_no,
            'customer': job_card.destination or '-',
            'product': job_card.planning_job.job_name if job_card.planning_job else (job_card.SKU or '-'),
            'machine': job_card.machine_name_display or '-',
            'paper': job_card.material.name if job_card.material else '-',
            'gsm': '-',
            'colors': (job_card.colour or '').strip() or (
                str(job_card.total_colors) if job_card.total_colors else '-'
            ),
            'order_qty': f'{job_card.order_qty:,}',
            'required_sheets': f'{int(job_card.total_sheet_quantity_display or 0):,}',
            'produced_qty': f'{int(job_card.total_production_pcs or 0):,}',
            'remaining_qty': f'{max(0, job_card.order_qty - (job_card.total_production_pcs or 0)):,}',
            'remaining_display': f'{max(0, job_card.order_qty - (job_card.total_production_pcs or 0)):,}',
            'due_date': (
                job_card.planning_job.delivery_date.strftime('%Y-%m-%d')
                if job_card.planning_job and job_card.planning_job.delivery_date else '-'
            ),
            'job_type': (
                job_card.planning_job.repeat_flag
                if job_card.planning_job and getattr(job_card.planning_job, 'repeat_flag', None) else 'New Job'
            ),
            'pass_count': pass_count,
            'pass_type': f'{pass_count}-pass' if pass_count > 1 else 'Single-pass',
            'pass_override': job_card.pass_count_override,
            'pass_override_reason': job_card.pass_count_override_reason or '',
            'planned_pass_baseline': job_card.planned_pass_baseline,
            'machine_pass_hint': get_degraded_machine_pass_hint(
                job_card, resolved_machine, job_card.planned_pass_baseline
            ),
            'passes_from_planning': pass_tracking['passes_from_planning'],
            'passes_inferred': pass_tracking['passes_inferred'],
            'legacy_notice': pass_tracking['legacy_notice'],
            'suggested_pass': pass_tracking['suggested_pass'],
            'per_pass_budget': pass_tracking['per_pass_budget'],
            'per_pass_budget_display': pass_tracking['per_pass_budget_display'],
            'pass_rows': pass_tracking['pass_rows'],
            'allowed_impressions': f'{allowed_impressions:,}',
            'used_impressions': f'{total_impressions_used:,}',
            'remaining_impressions': f'{remaining_impressions:,}',
            'total_impressions_used_display': pass_tracking['total_impressions_used_display'],
            'total_impressions_allowed_display': pass_tracking['total_impressions_allowed_display'],
            'total_impressions_remaining_display': pass_tracking['total_impressions_remaining_display'],
            'sheets_used_display': pass_tracking['sheets_used_display'],
            'sheets_allowed_display': pass_tracking['sheets_allowed_display'],
            'sheets_remaining_display': pass_tracking['sheets_remaining_display'],
            'good_sheets_target_display': pass_tracking['good_sheets_target_display'],
            'waiting_for_plate': job_is_waiting_for_plates(job_card),
            'history': history_data,
        }

        # Smart layout merge context for the operator.
        merge_item = job_card.planning_job.active_merge_item if job_card.planning_job else None
        if merge_item:
            group = merge_item.merge_group
            info_map[job_id]['merge'] = {
                'code': group.code,
                'is_lead': merge_item.is_lead,
                'lead_jc': group.lead_job.jc_number if group.lead_job else '',
                'allocated_ups': merge_item.allocated_ups,
                'run_sheets': group.run_sheets,
            }
            # The lead prints the whole combined sheet — show the combined run,
            # not this SKU's standalone sheet count, so 5000 and 1429 do not clash.
            if merge_item.is_lead and group.run_sheets:
                info_map[job_id]['required_sheets'] = f'{int(group.run_sheets):,}'
                combined_impr = group.combined_impressions()
                if combined_impr:
                    info_map[job_id]['allowed_impressions'] = f'{combined_impr:,}'
                    info_map[job_id]['remaining_impressions'] = f'{max(combined_impr - total_impressions_used, 0):,}'

    return plan_map, machine_map, info_map
