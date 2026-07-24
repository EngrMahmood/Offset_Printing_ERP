"""Print pass tracking for multi-pass production entry."""

from __future__ import annotations

import re

from core.models import JobCard

MAX_PRINT_PASSES = 4
INFERENCE_RATIO_TOLERANCE = 0.18


def has_explicit_planning_passes(job_card):
    return bool(job_card.planning_job and job_card.planning_job.print_passes)


def get_sheets_per_pass_basis(job_card):
    sheets = int(job_card.total_sheets_planned or 0)
    if sheets > 0:
        return sheets
    return int(job_card.total_impressions_required or 0)


def normalize_machine_name(job_card):
    return (job_card.machine_name_display or '').strip().upper()


def get_color_count(job_card):
    from core.print_colors import print_color_total_units

    if job_card.total_colors is not None and int(job_card.total_colors) > 0:
        return int(job_card.total_colors)
    units = print_color_total_units(job_card.colour)
    return units or None


def infer_pass_count_from_one_plus_one_colour(job_card):
    """Colour pattern like 1+1 implies two passes."""
    colour = (job_card.colour or '').strip()
    if not colour:
        return None, ''
    match = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', colour)
    if match and int(match.group(2)) > 0:
        return 2, 'colour 1+1'
    return None, ''


def infer_pass_count_from_machine_rules(job_card):
    """Machine + colour shop rules — lowest legacy priority."""
    machine = normalize_machine_name(job_card)
    color_count = get_color_count(job_card)

    if 'GTO' in machine:
        if color_count in {3, 4}:
            return 2, 'GTO with 3-4 colors'
        if color_count in {1, 2}:
            return 1, 'GTO with 1-2 colors'

    if 'SM74' in machine and color_count == 4:
        return 1, 'SM74 with 4 colors'

    return None, ''


def infer_pass_count_from_impressions(job_card):
    """Infer 1/2/3 passes when planning.print_passes is missing (legacy jobs)."""
    sheets = get_sheets_per_pass_basis(job_card)
    total_impressions = int(job_card.total_impressions_required or 0)
    if sheets <= 0 or total_impressions <= 0:
        return None

    ratio = total_impressions / sheets
    if ratio <= 1 + INFERENCE_RATIO_TOLERANCE:
        return 1

    for passes in range(MAX_PRINT_PASSES, 1, -1):
        if abs(ratio - passes) <= INFERENCE_RATIO_TOLERANCE:
            return passes

    rounded = max(1, min(MAX_PRINT_PASSES, round(ratio)))
    if rounded > 1 and abs(ratio - rounded) <= INFERENCE_RATIO_TOLERANCE:
        return rounded
    return None


def resolve_pass_inference(job_card):
    """Return pass count plus how it was determined."""
    override = getattr(job_card, 'pass_count_override', None)
    if override:
        return {
            'passes': int(override),
            'source': 'override',
            'reason': 'supervisor pass-count override',
            'uses_legacy_inference': False,
        }

    if has_explicit_planning_passes(job_card):
        passes = int(job_card.planning_job.print_passes)
        return {
            'passes': passes,
            'source': 'planning',
            'reason': 'planning print_passes',
            'uses_legacy_inference': False,
        }

    impression_passes = infer_pass_count_from_impressions(job_card)
    if impression_passes and impression_passes > 1:
        return {
            'passes': impression_passes,
            'source': 'impressions',
            'reason': 'impressions vs sheets',
            'uses_legacy_inference': True,
        }

    colour_passes, colour_reason = infer_pass_count_from_one_plus_one_colour(job_card)
    if colour_passes is not None:
        return {
            'passes': colour_passes,
            'source': 'colour',
            'reason': colour_reason,
            'uses_legacy_inference': True,
        }

    machine_passes, machine_reason = infer_pass_count_from_machine_rules(job_card)
    if machine_passes is not None:
        return {
            'passes': machine_passes,
            'source': 'rule',
            'reason': machine_reason,
            'uses_legacy_inference': True,
        }

    if impression_passes == 1:
        return {
            'passes': 1,
            'source': 'impressions',
            'reason': 'impressions vs sheets',
            'uses_legacy_inference': False,
        }

    if job_card.planning_job:
        front_pass = int(job_card.planning_job.front_pass or 0)
        back_pass = int(job_card.planning_job.back_pass or 0)
        if front_pass > 0 or back_pass > 0:
            passes = max(1, front_pass + back_pass)
            return {
                'passes': passes,
                'source': 'planning_fields',
                'reason': 'front/back pass fields',
                'uses_legacy_inference': passes > 1,
            }

    return {
        'passes': 1,
        'source': 'default',
        'reason': 'single-pass default',
        'uses_legacy_inference': not has_explicit_planning_passes(job_card),
    }


def get_job_card_pass_count(job_card):
    return resolve_pass_inference(job_card)['passes']


def passes_are_inferred(job_card):
    inference = resolve_pass_inference(job_card)
    return inference['uses_legacy_inference'] and inference['passes'] > 1


def effective_print_pass_number(production_row, total_passes):
    """Resolve pass number for legacy rows without print_pass_number."""
    if production_row.print_pass_number:
        return int(production_row.print_pass_number)
    if total_passes <= 1:
        return 1
    if production_row.intermediate_pass:
        return 1
    return total_passes


def get_per_pass_impression_budget(job_card):
    total_passes = get_job_card_pass_count(job_card)
    if total_passes <= 1:
        return int(job_card.total_impressions_required or 0)

    sheets = get_sheets_per_pass_basis(job_card)
    inference = resolve_pass_inference(job_card)
    if sheets > 0 and inference['source'] in {'colour', 'rule', 'impressions'}:
        return sheets

    total_required = int(job_card.total_impressions_required or 0)
    if total_required > 0:
        return int(round(total_required / total_passes))
    return sheets


def get_pass_impression_usage(job_card, exclude_production_id=None):
    total_passes = get_job_card_pass_count(job_card)
    usage = {pass_no: 0 for pass_no in range(1, total_passes + 1)}
    qs = job_card.productions.filter(is_active=True, entry_type='printing')
    if exclude_production_id:
        qs = qs.exclude(pk=exclude_production_id)
    for row in qs:
        pass_no = effective_print_pass_number(row, total_passes)
        usage[pass_no] = usage.get(pass_no, 0) + int(row.impressions or 0)
    return usage


def get_suggested_print_pass(job_card, exclude_production_id=None):
    total_passes = get_job_card_pass_count(job_card)
    if total_passes <= 1:
        return 1
    usage = get_pass_impression_usage(job_card, exclude_production_id)
    budget = get_per_pass_impression_budget(job_card)
    for pass_no in range(1, total_passes):
        if usage.get(pass_no, 0) < budget:
            return pass_no
    return total_passes


def _build_legacy_notice(job_card, inference, total_passes, sheets_basis):
    if has_explicit_planning_passes(job_card):
        return ''

    if inference['source'] in {'rule', 'colour'}:
        if total_passes > 1:
            return (
                f'Passes inferred from shop rule ({total_passes}-pass: {inference["reason"]}). '
                'Ask planner to set No. of Passes in planning.'
            )
        return (
            f'Single-pass inferred from shop rule ({inference["reason"]}). '
            'Set No. of Passes in planning when confirmed.'
        )
    if inference['source'] == 'impressions' and total_passes > 1:
        return (
            f'Passes inferred from job data ({total_passes}-pass: '
            f'{int(job_card.total_impressions_required or 0):,} impressions vs '
            f'{sheets_basis:,} sheets). Ask planner to set No. of Passes in planning.'
        )
    return (
        'No. of Passes not set in planning — treating as single-pass. '
        'Set passes in planning if this job needs multiple passes.'
    )


def build_pass_tracking_info(job_card, exclude_production_id=None):
    inference = resolve_pass_inference(job_card)
    total_passes = inference['passes']
    budget = get_per_pass_impression_budget(job_card)
    usage = get_pass_impression_usage(job_card, exclude_production_id)
    allowed_total = int(job_card.total_impressions_allowed_with_tolerance or 0)
    total_used = sum(usage.values())
    sheets_basis = get_sheets_per_pass_basis(job_card)

    pass_rows = []
    for pass_no in range(1, total_passes + 1):
        used = usage.get(pass_no, 0)
        is_final = pass_no >= total_passes
        pass_rows.append({
            'pass_number': pass_no,
            'label': f'Pass {pass_no}' + (' (final)' if is_final else ''),
            'used': used,
            'used_display': f'{used:,}',
            'budget': budget,
            'budget_display': f'{budget:,}',
            'remaining': max(0, budget - used),
            'remaining_display': f'{max(0, budget - used):,}',
            'is_final': is_final,
        })

    legacy_notice = _build_legacy_notice(job_card, inference, total_passes, sheets_basis)

    return {
        'total_passes': total_passes,
        'passes_from_planning': has_explicit_planning_passes(job_card),
        'passes_inferred': passes_are_inferred(job_card),
        'pass_inference_source': inference['source'],
        'pass_inference_reason': inference['reason'],
        'legacy_notice': legacy_notice,
        'per_pass_budget': budget,
        'per_pass_budget_display': f'{budget:,}',
        'sheets_per_pass_basis': sheets_basis,
        'sheets_per_pass_basis_display': f'{sheets_basis:,}',
        'pass_usage': usage,
        'pass_rows': pass_rows,
        'suggested_pass': get_suggested_print_pass(job_card, exclude_production_id),
        'total_impressions_used': total_used,
        'total_impressions_allowed': allowed_total,
        'total_impressions_remaining': max(0, allowed_total - total_used),
        'total_impressions_used_display': f'{total_used:,}',
        'total_impressions_allowed_display': f'{allowed_total:,}',
        'total_impressions_remaining_display': f'{max(0, allowed_total - total_used):,}',
    }


def validate_print_pass_number(job_card, pass_number, exclude_production_id=None):
    total_passes = get_job_card_pass_count(job_card)
    if pass_number < 1 or pass_number > total_passes:
        raise ValueError(f'Print pass must be between 1 and {total_passes} for this job.')

    is_final = pass_number >= total_passes

    if pass_number > 1 or not is_final:
        usage = get_pass_impression_usage(job_card, exclude_production_id)
        if pass_number > 1:
            prior_total = sum(usage.get(p, 0) for p in range(1, pass_number))
            if prior_total <= 0:
                raise ValueError(
                    f'Log Pass {pass_number - 1} impressions before starting Pass {pass_number}.'
                )
        # Block logging a NEW entry onto a non-final pass that is already
        # complete (its per-pass budget is used up) — the operator should move
        # to the next pass. Edits (exclude_production_id set) are never blocked
        # here, so existing/legacy rows stay correctable.
        if not is_final and exclude_production_id is None:
            budget = get_per_pass_impression_budget(job_card)
            if budget > 0 and usage.get(pass_number, 0) >= budget:
                raise ValueError(
                    f'Pass {pass_number} is already complete for this job — '
                    f'select the next pass.'
                )
    return {
        'total_passes': total_passes,
        'is_final_pass': is_final,
        'intermediate_pass': not is_final,
    }
