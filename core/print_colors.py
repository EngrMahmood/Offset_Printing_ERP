"""Production print-color master helpers (1, 2, 4, 1+1, …)."""

from __future__ import annotations

import re

from django.db.models import Q

_COLOR_PLUS_RE = re.compile(r'^(\d+)\s*\+\s*(\d+)$')
_COLOR_SINGLE_RE = re.compile(r'^(\d+)\s*(?:colou?r(?:s)?)?$', re.IGNORECASE)
_DECIMAL_COLOR_RE = re.compile(r'^(\d+)\.(\d+)$')
# Legacy sheet import mangled 4.0 → 40, 1.0 → 10, etc. (dot stripped before parsing).
_MANGLED_X0_COLOR_RE = re.compile(r'^([1-6])0$')


def get_active_print_colors():
    from core.models import PrintColor
    return PrintColor.objects.filter(is_active=True).order_by('sort_order', 'name')


def get_print_color_choices(*, include_legacy=None):
    colors = list(get_active_print_colors())
    choices = [('', 'Select Print Color')] + [(item.name, item.name) for item in colors]
    names = {item.name for item in colors}
    legacy = (include_legacy or '').strip()
    if legacy and legacy not in names:
        choices.append((legacy, f'{legacy} (legacy)'))
    return choices


def resolve_print_color_name(value):
    """Return canonical master name, or '' if not in master (unless exact legacy kept)."""
    from core.models import PrintColor

    raw = str(value or '').strip()
    if not raw:
        return ''
    match = PrintColor.objects.filter(name__iexact=raw, is_active=True).values_list('name', flat=True).first()
    return match or ''


def repair_mangled_decimal_color_spec(value):
    """Repair decimal mangling from legacy import (40 → 4+0, 10 → 1+0, …)."""
    raw = str(value or '').strip()
    if not raw:
        return ''
    match = _MANGLED_X0_COLOR_RE.fullmatch(raw)
    if match:
        return f'{int(match.group(1))}+0'
    return raw


def normalize_color_spec_value(raw_value):
    """Normalize Color column / color_spec for storage (sheet import, forms, cleanup)."""
    raw_text = str(raw_value or '').strip()
    if not raw_text:
        return ''

    repaired = repair_mangled_decimal_color_spec(raw_text)
    if repaired != raw_text:
        raw_text = repaired

    resolved = resolve_print_color_name(raw_text)
    if resolved:
        return resolved

    # Sheet stores 4+0 as Excel decimal 4.0 — must run before stripping punctuation.
    decimal_match = _DECIMAL_COLOR_RE.fullmatch(raw_text)
    if decimal_match:
        candidate = f'{int(decimal_match.group(1))}+{int(decimal_match.group(2))}'
        return resolve_print_color_name(candidate) or candidate

    lowered = raw_text.lower()
    if lowered in {'no', 'none', 'n/a', 'na', 'nil'}:
        return ''

    usable = False
    if re.search(r'color|colour|colours|colors', lowered):
        usable = True
    elif re.search(r'\d+\s*c\b', lowered) or ('c' in lowered and re.search(r'\d', lowered)):
        usable = True
    elif any(sep in lowered for sep in ['+', '/', '-']):
        usable = True
    elif raw_text.isdigit() or _DECIMAL_COLOR_RE.fullmatch(raw_text):
        usable = True

    if not usable:
        return raw_text

    normalized = lowered.replace('colours', 'color').replace('colour', 'color').replace('colors', 'color')
    normalized = normalized.replace('c/', '+').replace('c+', '+').replace('/', '+').replace('-', '+')
    normalized = re.sub(r'[^0-9\+\s]+', '', normalized).strip()
    normalized = re.sub(r'\s+', '+', normalized)
    normalized = re.sub(r'\++', '+', normalized)

    plus_match = _COLOR_PLUS_RE.fullmatch(normalized)
    if plus_match:
        candidate = f'{int(plus_match.group(1))}+{int(plus_match.group(2))}'
        return resolve_print_color_name(candidate) or candidate

    single_match = _COLOR_SINGLE_RE.fullmatch(normalized)
    if single_match:
        candidate = str(int(single_match.group(1)))
        return resolve_print_color_name(candidate) or candidate

    numbers = re.findall(r'[0-9]+', normalized)
    if len(numbers) == 1:
        candidate = str(int(numbers[0]))
        return resolve_print_color_name(candidate) or candidate
    if len(numbers) == 2:
        candidate = f'{int(numbers[0])}+{int(numbers[1])}'
        return resolve_print_color_name(candidate) or candidate

    return raw_text


def normalize_print_color_for_form(value):
    """Map legacy text (e.g. '4 color') to master name when possible, else keep legacy."""
    return normalize_color_spec_value(value)


def print_color_total_units(value):
    """Numeric colour units for setup/pass logic."""
    from core.models import PrintColor

    raw = str(value or '').strip()
    if not raw:
        return 0

    master = PrintColor.objects.filter(name__iexact=raw, is_active=True).first()
    if master:
        return int(master.total_units or 0)

    # Legacy patterns still supported for count only
    plus = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', raw)
    if plus:
        return int(plus.group(1)) + int(plus.group(2))
    single = re.fullmatch(r'(\d+)(?:\s*colou?rs?)?', raw, re.IGNORECASE)
    if single:
        return int(single.group(1))
    digits = re.findall(r'\d+', raw)
    if len(digits) == 1:
        return int(digits[0])
    if len(digits) == 2:
        return int(digits[0]) + int(digits[1])
    return 0


def apply_print_color_to_planning_job(job, color_name=None):
    """Set planning job color_spec + total_colors from print color master."""
    if job is None:
        return False
    name = resolve_print_color_name(color_name if color_name is not None else job.color_spec)
    if not name:
        # Keep legacy text but still try to derive count
        units = print_color_total_units(job.color_spec)
        if units and job.total_colors != units:
            job.total_colors = units
            return True
        return False

    units = print_color_total_units(name)
    changed = False
    if (job.color_spec or '').strip() != name:
        job.color_spec = name
        changed = True
    if job.total_colors != units:
        job.total_colors = units
        changed = True
    return changed


def apply_print_color_to_sku_recipe(recipe, color_name=None):
    """Normalize SKU recipe color_spec to master print color name."""
    if recipe is None:
        return False
    name = resolve_print_color_name(color_name if color_name is not None else recipe.color_spec)
    if not name:
        return False
    if (recipe.color_spec or '').strip() != name:
        recipe.color_spec = name
        return True
    return False


def sync_print_color_from_recipe_to_jobs(sku, color_name):
    """Push approved print color onto active planning jobs for this SKU."""
    from planning.models import PlanningJob

    name = resolve_print_color_name(color_name)
    if not name:
        return 0
    units = print_color_total_units(name)
    updated = 0
    for job in PlanningJob.objects.filter(sku__iexact=sku, is_active=True):
        fields = []
        if (job.color_spec or '').strip() != name:
            job.color_spec = name
            fields.append('color_spec')
        if job.total_colors != units:
            job.total_colors = units
            fields.append('total_colors')
        if fields:
            fields.append('updated_at')
            job.save(update_fields=fields)
            updated += 1
    return updated
