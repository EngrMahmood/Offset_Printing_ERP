import re

from django.db import transaction
from django.utils import timezone

from core.models import SequenceCounter

from .models import FlexoPlanningJob

# Format: JC-FPP-MM-YY-NN, one running series that resets every year (like offset).
JC_PREFIX = 'FPP'
_JC_PATTERN = re.compile(r'^JC-' + JC_PREFIX + r'-\d{2}-(\d{2})-(\d+)$')


def _max_existing_serial(yy):
    max_serial = 0
    for number in FlexoPlanningJob.objects.filter(
        jc_number__startswith=f'JC-{JC_PREFIX}-'
    ).values_list('jc_number', flat=True):
        match = _JC_PATTERN.match(str(number or '').strip())
        if match and match.group(1) == yy:
            max_serial = max(max_serial, int(match.group(2)))
    return max_serial


@transaction.atomic
def allocate_next_flexo_jc_number(for_date=None):
    """Next Flexo JC number, with DB locking (same pattern as
    core.jc_numbering.allocate_next_jc_number) under a per-year counter key."""
    date_value = for_date or timezone.localdate()
    yy = f'{date_value:%y}'
    counter, _ = SequenceCounter.objects.select_for_update().get_or_create(
        key=f'flexo_jc:{yy}',
        defaults={'last_value': 0},
    )

    max_existing = _max_existing_serial(yy)
    if max_existing > counter.last_value:
        counter.last_value = max_existing

    counter.last_value += 1
    counter.save(update_fields=['last_value', 'updated_at'])

    return f"JC-{JC_PREFIX}-{date_value:%m}-{yy}-{counter.last_value:02d}"
