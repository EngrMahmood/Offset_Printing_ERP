import re

from django.apps import apps
from django.db import transaction
from django.utils import timezone

from core.models import JobCard, SequenceCounter


_JC_PATTERN = re.compile(r'^JC-\d{2}-\d{2}-(\d+)(?:\.\d+)?$')


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
    """Allocate the next JC number in format JC-MM-YY-#### with DB locking."""
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
    return f"JC-{date_value:%m}-{date_value:%y}-{counter.last_value:04d}"
