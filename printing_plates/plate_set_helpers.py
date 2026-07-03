"""Multi-set plate quantity helpers (impressions ÷ machine plate life)."""

from __future__ import annotations

import math

from core.models import Machine

DEFAULT_PLATE_LIFE_IMPRESSIONS = 25000


def count_plate_inks(plate_color):
    return len([part.strip() for part in str(plate_color or '').split(',') if part.strip()])


def get_plate_life_impressions(machine=None):
    if machine is not None:
        life = getattr(machine, 'plate_life_impressions', None)
        if life:
            return int(life)
    return DEFAULT_PLATE_LIFE_IMPRESSIONS


def resolve_plate_machine(plate_request):
    if plate_request is None:
        return None
    if plate_request.machine_id:
        return plate_request.machine

    job_card = getattr(plate_request, 'job_card', None)
    if job_card is not None and getattr(job_card, 'machine_name_id', None):
        return job_card.machine_name

    planning_job = getattr(plate_request, 'planning_job', None)
    machine_name = (getattr(planning_job, 'machine_name', None) or '').strip()
    if machine_name:
        return Machine.objects.filter(name__iexact=machine_name, is_active=True).first()
    return None


def suggest_sets_required(impressions, plate_life=None):
    life = int(plate_life or DEFAULT_PLATE_LIFE_IMPRESSIONS) or DEFAULT_PLATE_LIFE_IMPRESSIONS
    total = int(impressions or 0)
    if total <= 0:
        return 1
    return max(1, math.ceil(total / life))


def suggest_plate_quantity(sets_required, ink_count):
    sets_value = max(1, int(sets_required or 1))
    inks = max(0, int(ink_count or 0))
    if inks <= 0:
        return None
    return sets_value * inks


def build_plate_set_suggestion(plate_request, *, plate_color=None):
    machine = resolve_plate_machine(plate_request)
    impressions = plate_request.impressions if plate_request else None
    plate_life = get_plate_life_impressions(machine)
    ink_count = count_plate_inks(plate_color if plate_color is not None else getattr(plate_request, 'plate_color', ''))
    sets_required = suggest_sets_required(impressions, plate_life)
    plate_quantity = suggest_plate_quantity(sets_required, ink_count)
    return {
        'machine': machine,
        'impressions': int(impressions or 0),
        'plate_life_impressions': plate_life,
        'ink_count': ink_count,
        'sets_required': sets_required,
        'plate_quantity': plate_quantity,
    }


def format_plate_quantity_display(plate_quantity, sets_required):
    qty = plate_quantity
    sets = sets_required
    if qty in (None, ''):
        return ''
    if sets:
        return f'{qty} ({sets} set{"s" if int(sets) != 1 else ""})'
    return str(qty)
