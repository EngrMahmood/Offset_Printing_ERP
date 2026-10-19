"""Aggregated maintenance report for a date range.

Definitions (also shown on the report page):

* In scope: active tickets reported in the range; CANCELLED tickets are excluded.
* Breakdown: a ticket classified Breakdown or one that stopped the machine.
* Downtime: machine-unavailable time from MachineDowntime intervals, clipped to
  the range. An interval still open is counted up to now and flagged "ongoing".
* MTTR: average length of the CLOSED breakdown downtime intervals that started in
  the range (so it measures time the machine was out of production, not
  technician labour). Open intervals are excluded until they close.
* Repair vs waiting time: time tickets spent In Progress versus Awaiting Parts /
  Awaiting Vendor, read from the status history.
* Repeat fault: the same machine and fault category on two or more tickets.
* PM compliance: preventive tickets raised in the range that reached Verified or
  Closed, as a share of those raised (cancelled excluded).
* MTBF is deliberately not computed: it needs the machine's operating hours, and
  the ERP does not record them reliably.

Cost figures are not part of this stage of the work.
"""
from collections import defaultdict
from datetime import datetime, time, timedelta

from django.db.models import Count, Q
from django.utils import timezone

from .models import MachineDowntime, MaintenanceActivityLog, MaintenanceRecord

WAITING = {'AWAITING_PARTS', 'AWAITING_VENDOR'}


def _aware(day, end_of_day=False):
    stamp = datetime.combine(day, time.max if end_of_day else time.min)
    return timezone.make_aware(stamp)


def build_report(date_from, date_to):
    start, end = _aware(date_from), _aware(date_to, end_of_day=True)
    now = timezone.now()

    scope = (
        MaintenanceRecord.objects.filter(is_active=True, reported_date__gte=date_from, reported_date__lte=date_to)
        .exclude(status='CANCELLED')
    )
    breakdown_q = Q(maintenance_type='BREAKDOWN') | Q(machine_stopped=True)

    machine_rows = {}
    for row in scope.values('machine_id', 'machine__name').annotate(
        tickets=Count('id'), breakdowns=Count('id', filter=breakdown_q),
    ):
        machine_rows[row['machine_id']] = {
            'machine': row['machine__name'], 'tickets': row['tickets'], 'breakdowns': row['breakdowns'],
            'downtime_hours': 0.0, 'ongoing': False, 'mttr_minutes': None, '_repairs': [],
            'total_cost': sum((r.total_cost for r in scope.filter(machine_id=row['machine_id'])), 0),
        }

    intervals = MachineDowntime.objects.filter(
        Q(end_at__isnull=True) | Q(end_at__gt=start), start_at__lte=end,
    ).select_related('machine')
    for interval in intervals:
        row = machine_rows.setdefault(interval.machine_id, {
            'machine': interval.machine.name, 'tickets': 0, 'breakdowns': 0, 'downtime_hours': 0.0,
            'ongoing': False, 'mttr_minutes': None, '_repairs': [], 'total_cost': 0,
        })
        clipped_start = max(interval.start_at, start)
        clipped_end = min(interval.end_at or now, end)
        if clipped_end > clipped_start:
            row['downtime_hours'] += (clipped_end - clipped_start).total_seconds() / 3600
        if interval.end_at is None:
            row['ongoing'] = True
        elif interval.reason == 'BREAKDOWN' and start <= interval.start_at <= end:
            row['_repairs'].append((interval.end_at - interval.start_at).total_seconds() / 60)

    all_repairs = []
    for row in machine_rows.values():
        if row['_repairs']:
            row['mttr_minutes'] = sum(row['_repairs']) / len(row['_repairs'])
            all_repairs.extend(row['_repairs'])
        row['downtime_hours'] = round(row['downtime_hours'], 1)
    by_machine = sorted(machine_rows.values(), key=lambda r: (-r['breakdowns'], -r['downtime_hours'], r['machine']))

    by_type = {
        row['maintenance_type'] or 'UNCLASSIFIED': row['n']
        for row in scope.values('maintenance_type').annotate(n=Count('id'))
    }

    repeats = list(
        scope.filter(fault_types__isnull=False)
        .values('machine__name', 'fault_types__name')
        .annotate(n=Count('id', distinct=True)).filter(n__gte=2).order_by('-n', 'machine__name')[:20]
    )

    # Repair vs waiting time, from the status history of in-scope tickets.
    in_scope_ids = list(scope.values_list('id', flat=True))
    history = defaultdict(list)
    for record_id, to_status, stamp in (
        MaintenanceActivityLog.objects.filter(record_id__in=in_scope_ids).exclude(to_status='')
        .order_by('record_id', 'created_at').values_list('record_id', 'to_status', 'created_at')
    ):
        history[record_id].append((to_status, stamp))
    repair_seconds = waiting_seconds = 0.0
    for events in history.values():
        for index, (status, stamp) in enumerate(events):
            if status != 'IN_PROGRESS' and status not in WAITING:
                continue
            until = events[index + 1][1] if index + 1 < len(events) else now
            seconds = max((until - stamp).total_seconds(), 0)
            if status == 'IN_PROGRESS':
                repair_seconds += seconds
            else:
                waiting_seconds += seconds

    pm = scope.filter(pm_plan__isnull=False)
    pm_total = pm.count()
    pm_done = pm.filter(status__in=('VERIFIED', 'CLOSED')).count()

    return {
        'total_tickets': scope.count(),
        'open_count': scope.filter(status__in=MaintenanceRecord.OPEN_STATUSES).count(),
        'closed_count': scope.filter(status='CLOSED').count(),
        'by_machine': by_machine,
        'total_downtime_hours': round(sum(r['downtime_hours'] for r in by_machine), 1),
        'overall_mttr_minutes': (sum(all_repairs) / len(all_repairs)) if all_repairs else None,
        'by_type': by_type,
        'breakdown_tickets': by_type.get('BREAKDOWN', 0),
        'preventive_tickets': by_type.get('PREVENTIVE', 0),
        'repeat_faults': repeats,
        'repair_hours': round(repair_seconds / 3600, 1),
        'waiting_hours': round(waiting_seconds / 3600, 1),
        'pm_total': pm_total,
        'pm_done': pm_done,
        'pm_open': pm.filter(status__in=MaintenanceRecord.OPEN_STATUSES).count(),
        'pm_compliance_pct': round(pm_done * 100 / pm_total, 1) if pm_total else None,
    }
