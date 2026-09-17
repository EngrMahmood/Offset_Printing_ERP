import re

from django.apps import apps
from django.db import transaction
from django.utils import timezone

from core.models import JobCard, SequenceCounter


_JC_PATTERN = re.compile(r'^JC-\d{2}-\d{2}-(?:PP-)?(\d+)(?:\.\d+)?$')


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


def _max_existing_jc_serial():
    max_serial = 0

    for number in JobCard.objects.values_list('job_card_no', flat=True):
        serial = _extract_serial(number)
        if serial and serial > max_serial:
            max_serial = serial

    try:
        PlanningJob = apps.get_model('planning', 'PlanningJob')
        for number in PlanningJob.objects.values_list('jc_number', flat=True):
            serial = _extract_serial(number)
            if serial and serial > max_serial:
                max_serial = serial
    except LookupError:
        # Planning app may not be loaded in some contexts.
        pass

    return max_serial


@transaction.atomic
def allocate_next_jc_number(for_date=None):
    """Allocate the next JC number in format JC-MM-YY-PP-#### with DB locking."""
    counter, _ = SequenceCounter.objects.select_for_update().get_or_create(
        key='jc_global',
        defaults={'last_value': 0},
    )

    max_existing = _max_existing_jc_serial()
    if max_existing > counter.last_value:
        counter.last_value = max_existing

    counter.last_value += 1
    counter.save(update_fields=['last_value', 'updated_at'])

    date_value = for_date or timezone.localdate()
    suffix = 'PP-'
    return f"JC-{date_value:%m}-{date_value:%y}-{suffix}{counter.last_value:04d}"


def _max_child_suffix(base_jc_number):
    pattern = re.compile(r'^' + re.escape(base_jc_number) + r'\.(\d+)$')
    max_suffix = 0

    for number in JobCard.objects.filter(
        job_card_no__startswith=f'{base_jc_number}.'
    ).values_list('job_card_no', flat=True):
        match = pattern.match(str(number).strip())
        if match:
            max_suffix = max(max_suffix, int(match.group(1)))

    try:
        PlanningJob = apps.get_model('planning', 'PlanningJob')
        for number in PlanningJob.objects.filter(
            jc_number__startswith=f'{base_jc_number}.'
        ).values_list('jc_number', flat=True):
            match = pattern.match(str(number).strip())
            if match:
                max_suffix = max(max_suffix, int(match.group(1)))
    except LookupError:
        pass

    return max_suffix


@transaction.atomic
def allocate_child_jc_number(base_jc_number):
    """Allocate the next dotted-suffix JC number for a sibling form under
    `base_jc_number` (e.g. 'JC-01-26-1115' -> 'JC-01-26-1115.1', then '.2', ...).

    `base_jc_number` must be the root JC number of the group — callers should
    never pass an already-suffixed number, so siblings stay one level deep
    (no 'JC-...-1115.1.1').
    """
    counter_key = f'jc_child:{base_jc_number}'
    counter, _ = SequenceCounter.objects.select_for_update().get_or_create(
        key=counter_key,
        defaults={'last_value': 0},
    )

    max_existing = _max_child_suffix(base_jc_number)
    if max_existing > counter.last_value:
        counter.last_value = max_existing

    counter.last_value += 1
    counter.save(update_fields=['last_value', 'updated_at'])

    return f'{base_jc_number}.{counter.last_value}'
