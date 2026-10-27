"""Keep non-offset machines (Flexo) out of offset-only screens, counts and reports.

Machine is one shared master. Offset screens that list "all machines" call
`offset_machines()` so a machine whose Machine Type is Flexo Printing never shows up
in offset dropdowns or distorts offset KPIs. Maintenance and downtime keep using the
unfiltered master on purpose, so both plants are covered there.
"""
from .models import Machine

# Machine Types that belong to the Flexo plant only.
FLEXO_ONLY_MACHINE_TYPES = ('flexo_printing',)
# Machine Types the Flexo screens may offer (printing + slitting/cutting + generic).
FLEXO_SCREEN_MACHINE_TYPES = ('flexo_printing', 'cutting', 'other', '')


def offset_machines(queryset=None):
    """Machines for offset-only use (everything except Flexo-only machines)."""
    qs = Machine.objects.all() if queryset is None else queryset
    return qs.exclude(machine_type__in=FLEXO_ONLY_MACHINE_TYPES)


def flexo_machines(queryset=None):
    """Machines the Flexo screens may offer (never offset/digital presses)."""
    qs = Machine.objects.all() if queryset is None else queryset
    return qs.filter(machine_type__in=FLEXO_SCREEN_MACHINE_TYPES)
