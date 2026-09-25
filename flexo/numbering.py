import re

from django.db import transaction
from django.utils import timezone

from core.models import SequenceCounter

from .models import FlexoPlanningJob

# Assumed prefix codes — the source spreadsheet's own manual numbering used
# "FS" (FL Planning's job card samples) and "FP" (SATIN's), which don't cleanly
# decode; these are readable placeholders pending confirmation with the user.
FAMILY_PREFIX = {
    'fl': 'FL',
    'satin': 'ST',
}

_JC_PATTERN = re.compile(r'^JC-[A-Z]+-\d{2}-\d{2}-(\d+)$')


def _extract_serial(jc_number):
    if not jc_number:
        return None
    match = _JC_PATTERN.match(str(jc_number).strip())
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None


def _max_existing_serial(family):
    prefix = f"JC-{FAMILY_PREFIX[family]}-"
    max_serial = 0
    for number in FlexoPlanningJob.objects.filter(
        material_family=family, jc_number__startswith=prefix
    ).values_list('jc_number', flat=True):
        serial = _extract_serial(number)
        if serial and serial > max_serial:
            max_serial = serial
    return max_serial


@transaction.atomic
def allocate_next_flexo_jc_number(material_family, for_date=None):
    """Allocate the next JC number for a Flexo family in format
    JC-<FL|ST>-MM-YY-####, with DB locking — mirrors
    core.jc_numbering.allocate_next_jc_number()'s pattern using the same
    SequenceCounter model, under a family-specific counter key."""
    if material_family not in FAMILY_PREFIX:
        raise ValueError(f"Unknown material_family: {material_family}")

    counter_key = f'flexo_jc:{material_family}'
    counter, _ = SequenceCounter.objects.select_for_update().get_or_create(
        key=counter_key,
        defaults={'last_value': 0},
    )

    max_existing = _max_existing_serial(material_family)
    if max_existing > counter.last_value:
        counter.last_value = max_existing

    counter.last_value += 1
    counter.save(update_fields=['last_value', 'updated_at'])

    date_value = for_date or timezone.localdate()
    prefix = FAMILY_PREFIX[material_family]
    return f"JC-{prefix}-{date_value:%m}-{date_value:%y}-{counter.last_value:04d}"
