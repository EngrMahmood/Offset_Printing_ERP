"""Process-wise downtime analysis.

Every stretch of machine downtime is recorded as a sequence of phases (see
``DowntimePhase``): waiting for maintenance, repair, waiting for spare parts,
waiting for an outside vendor, testing. This module turns those phases into the
answer to "where did the downtime go, and who owns it?".

A "waiting for spare parts" phase is split further by what Supply Chain was doing,
using the dates on the linked Item Request (request raised, PR, PO, received):

    before the request was raised   -> Maintenance
    request review & approval       -> Supply Chain
    purchasing (PR to PO)           -> Supply Chain / Procurement
    supplier delivery (PO to GRN)   -> Supplier
    received, waiting to be fitted  -> Maintenance

so the procurement team can be shown exactly how many days were theirs.
"""
from collections import defaultdict
from datetime import datetime, time, timedelta

from django.db.models import Q
from django.utils import timezone

from .models import DowntimePhase, MachineDowntime, MaintenanceRecord

TEAM_MAINTENANCE = 'maintenance'
TEAM_SUPPLY = 'supply_chain'
TEAM_SUPPLIER = 'supplier'
TEAM_VENDOR = 'vendor'
TEAM_UNTRACKED = 'untracked'

TEAM_LABELS = {
    TEAM_MAINTENANCE: 'Internal maintenance team',
    TEAM_SUPPLY: 'Supply Chain / Procurement',
    TEAM_SUPPLIER: 'Supplier delivery',
    TEAM_VENDOR: 'Outside repair vendor',
    TEAM_UNTRACKED: 'Not tracked by step',
}
TEAM_ORDER = [TEAM_MAINTENANCE, TEAM_SUPPLY, TEAM_SUPPLIER, TEAM_VENDOR, TEAM_UNTRACKED]

# (step key, label, owning team) for every step shown on the downtime window.
STEPS = {
    DowntimePhase.WAIT_RESPONSE: ('Waiting for maintenance to respond', TEAM_MAINTENANCE),
    DowntimePhase.WAIT_START: ('Diagnosed - waiting to start repair', TEAM_MAINTENANCE),
    DowntimePhase.REPAIR: ('Repair in progress', TEAM_MAINTENANCE),
    DowntimePhase.TESTING: ('Testing & verification', TEAM_MAINTENANCE),
    DowntimePhase.WAIT_VENDOR: ('Waiting for outside vendor', TEAM_VENDOR),
    DowntimePhase.UNTRACKED: ('Before step tracking', TEAM_UNTRACKED),
    'PARTS_BEFORE_REQUEST': ('Parts: before the request was raised', TEAM_MAINTENANCE),
    'PARTS_APPROVAL': ('Parts: request review & approval', TEAM_SUPPLY),
    'PARTS_PURCHASING': ('Parts: purchasing (PR to PO)', TEAM_SUPPLY),
    'PARTS_DELIVERY': ('Parts: supplier delivery (PO to receipt)', TEAM_SUPPLIER),
    'PARTS_NOT_FITTED': ('Parts: received, waiting to be fitted', TEAM_MAINTENANCE),
    'PARTS_UNLINKED': ('Parts: no Supply Chain request linked', TEAM_MAINTENANCE),
}
STEP_ORDER = list(STEPS)


def fmt_duration(seconds):
    """'2d 4h', '3h 15m', '12m' - readable at a glance."""
    seconds = max(int(seconds or 0), 0)
    days, rest = divmod(seconds, 86400)
    hours, rest = divmod(rest, 3600)
    minutes = rest // 60
    if days:
        return f'{days}d {hours}h'
    if hours:
        return f'{hours}h {minutes}m'
    return f'{minutes}m'


def _to_dt(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value if timezone.is_aware(value) else timezone.make_aware(value)
    return timezone.make_aware(datetime.combine(value, time.min))


def gating_request(record):
    """The Supply Chain request the repair was actually waiting on.

    Of the parts requested for the ticket, the one received last (or not received
    yet) is what held the machine up.
    """
    best, best_key = None, None
    if record is None:
        return None
    for line in record.spare_parts.all():
        request = line.item_request
        if request is None:
            continue
        procurement = getattr(request, 'procurement', None)
        received = _to_dt(getattr(procurement, 'received_date', None)) if procurement else None
        key = received or timezone.make_aware(datetime.max.replace(year=3000))
        if best_key is None or key > best_key:
            best, best_key = request, key
    return best


def parts_stage_segments(phase, now=None):
    """Split a WAIT_PARTS phase into (step key, start, end) pieces using procurement dates."""
    now = now or timezone.now()
    phase_start, phase_end = phase.started_at, phase.ended_at or now
    request = gating_request(phase.record)
    if request is None or phase_end <= phase_start:
        return [('PARTS_UNLINKED', phase_start, phase_end)] if phase_end > phase_start else []
    procurement = getattr(request, 'procurement', None)
    starts = [
        ('PARTS_APPROVAL', _to_dt(request.request_date)),
        ('PARTS_PURCHASING', _to_dt(getattr(procurement, 'pr_date', None))),
        ('PARTS_DELIVERY', _to_dt(getattr(procurement, 'po_date', None))),
        ('PARTS_NOT_FITTED', _to_dt(getattr(procurement, 'received_date', None))),
    ]
    stages, previous = [], None
    for key, when in starts:
        if when is None:
            continue
        when = max(when, previous) if previous else when      # dates must not run backwards
        stages.append((key, when))
        previous = when
    if not stages:
        return [('PARTS_UNLINKED', phase_start, phase_end)]
    pieces = []
    if stages[0][1] > phase_start:
        pieces.append(('PARTS_BEFORE_REQUEST', phase_start, min(stages[0][1], phase_end)))
    for index, (key, when) in enumerate(stages):
        end = stages[index + 1][1] if index + 1 < len(stages) else phase_end
        start, end = max(when, phase_start), min(end, phase_end)
        if end > start:
            pieces.append((key, start, end))
    return pieces


def phase_pieces(phase, now=None):
    """(step key, start, end) pieces for one phase - WAIT_PARTS is split, others are whole."""
    now = now or timezone.now()
    if phase.category == DowntimePhase.WAIT_PARTS:
        return parts_stage_segments(phase, now)
    end = phase.ended_at or now
    return [(phase.category, phase.started_at, end)] if end > phase.started_at else []


def _clip(start, end, window_start, window_end):
    start, end = max(start, window_start), min(end, window_end)
    return (end - start).total_seconds() if end > start else 0.0


def _phases_in_window(window_start, window_end, machine_id=None):
    qs = DowntimePhase.objects.filter(started_at__lte=window_end).filter(
        Q(ended_at__isnull=True) | Q(ended_at__gt=window_start)
    ).select_related('downtime__machine', 'record').prefetch_related(
        'record__spare_parts__item_request__procurement',
    )
    if machine_id:
        qs = qs.filter(downtime__machine_id=machine_id)
    return qs


def build_downtime_analysis(date_from, date_to, machine_id=None):
    now = timezone.now()
    window_start = _to_dt(date_from)
    window_end = min(_to_dt(date_to + timedelta(days=1)), now)

    step_seconds = defaultdict(float)
    step_count = defaultdict(int)
    machine_team = defaultdict(lambda: defaultdict(float))
    machine_names = {}
    waits = []
    for phase in _phases_in_window(window_start, window_end, machine_id):
        machine = phase.downtime.machine
        machine_names[machine.pk] = machine.name
        counted = False
        for key, start, end in phase_pieces(phase, now):
            seconds = _clip(start, end, window_start, window_end)
            if seconds <= 0:
                continue
            step_seconds[key] += seconds
            counted = True
            machine_team[machine.pk][STEPS[key][1]] += seconds
        if counted:
            for key in {piece[0] for piece in phase_pieces(phase, now)}:
                step_count[key] += 1
        if phase.category in (DowntimePhase.WAIT_PARTS, DowntimePhase.WAIT_VENDOR):
            waits.append(_describe_wait(phase, now))

    total = sum(step_seconds.values())
    team_seconds = defaultdict(float)
    for key, seconds in step_seconds.items():
        team_seconds[STEPS[key][1]] += seconds

    teams = [{
        'key': team, 'label': TEAM_LABELS[team], 'seconds': team_seconds.get(team, 0.0),
        'duration': fmt_duration(team_seconds.get(team, 0.0)),
        'hours': round(team_seconds.get(team, 0.0) / 3600, 1),
        'pct': round(team_seconds.get(team, 0.0) * 100 / total, 1) if total else 0.0,
    } for team in TEAM_ORDER if team_seconds.get(team, 0) >= 60]

    steps = [{
        'key': key, 'label': STEPS[key][0], 'team': TEAM_LABELS[STEPS[key][1]], 'team_key': STEPS[key][1],
        'seconds': step_seconds[key], 'duration': fmt_duration(step_seconds[key]),
        'hours': round(step_seconds[key] / 3600, 1),
        'pct': round(step_seconds[key] * 100 / total, 1) if total else 0.0,
        'occurrences': step_count.get(key, 0),
        'average': fmt_duration(step_seconds[key] / step_count[key]) if step_count.get(key) else '-',
    } for key in STEP_ORDER if step_seconds.get(key, 0) >= 60]

    machines = []
    for machine_pk, by_team in machine_team.items():
        machine_total = sum(by_team.values())
        machines.append({
            'machine': machine_names[machine_pk], 'total': fmt_duration(machine_total), 'seconds': machine_total,
            'teams': {team: fmt_duration(by_team.get(team, 0)) for team in TEAM_ORDER},
        })
    machines.sort(key=lambda row: -row['seconds'])

    waits.sort(key=lambda row: (not row['ongoing'], -row['seconds']))
    return {
        'total_seconds': total, 'total': fmt_duration(total), 'teams': teams, 'steps': steps, 'machines': machines,
        'ongoing_waits': [w for w in waits if w['ongoing']],
        'longest_waits': sorted(waits, key=lambda row: -row['seconds'])[:15],
        'team_labels': TEAM_LABELS,
        'team_order': TEAM_ORDER,
    }


def _describe_wait(phase, now):
    """One parts / vendor wait, with what it is waiting on and where Supply Chain has got to."""
    end = phase.ended_at or now
    seconds = max((end - phase.started_at).total_seconds(), 0)
    record = phase.record
    request, supplier, stage = None, '', ''
    if phase.category == DowntimePhase.WAIT_PARTS:
        request = gating_request(record)
        pieces = parts_stage_segments(phase, now)
        if pieces:
            stage = STEPS[pieces[-1][0]][0].replace('Parts: ', '')
    else:
        job = record.service_jobs.select_related('item_request').first() if record else None
        request = job.item_request if job else None
        supplier = job.vendor if job else ''
        stage = 'with the outside vendor'
    procurement = getattr(request, 'procurement', None) if request else None
    if procurement and procurement.supplier:
        supplier = procurement.supplier
    return {
        'machine': phase.downtime.machine.name,
        'record': record,
        'kind': 'Spare parts' if phase.category == DowntimePhase.WAIT_PARTS else 'Outside vendor',
        'request_no': request.request_no if request else '',
        'request_pk': request.pk if request else None,
        'item': request.item_title if request else '',
        'request_status': request.get_status_display() if request else '',
        'supplier': supplier,
        'stage': stage,
        'started_at': phase.started_at,
        'seconds': seconds,
        'duration': fmt_duration(seconds),
        'ongoing': phase.ended_at is None,
    }


def interval_timeline(interval, now=None):
    """Compact per-step breakdown of one downtime interval for the expandable list."""
    now = now or timezone.now()
    totals = defaultdict(float)
    for phase in interval.phases.all():
        for key, start, end in phase_pieces(phase, now):
            totals[key] += max((end - start).total_seconds(), 0)
    return [
        {'label': STEPS[key][0], 'team': TEAM_LABELS[STEPS[key][1]], 'duration': fmt_duration(totals[key])}
        for key in STEP_ORDER if totals.get(key)
    ]


def recent_intervals(date_from, date_to, machine_id=None, limit=40):
    window_start = _to_dt(date_from)
    window_end = _to_dt(date_to + timedelta(days=1))
    qs = MachineDowntime.objects.filter(start_at__lte=window_end).filter(
        Q(end_at__isnull=True) | Q(end_at__gt=window_start)
    ).select_related('machine', 'record').prefetch_related(
        'phases__record__spare_parts__item_request__procurement',
    ).order_by('-start_at')
    if machine_id:
        qs = qs.filter(machine_id=machine_id)
    now = timezone.now()
    rows = []
    for interval in qs[:limit]:
        end = interval.end_at or now
        rows.append({
            'interval': interval,
            'duration': fmt_duration((end - interval.start_at).total_seconds()),
            'ongoing': interval.end_at is None,
            'steps': interval_timeline(interval, now),
        })
    return rows


def machine_overview(user_is_staff, user):
    """Machine-wise board: current state, open tickets and all tickets per machine."""
    from core.models import Machine

    now = timezone.now()
    tickets = MaintenanceRecord.objects.filter(is_active=True).select_related('machine', 'assigned_to')
    if not user_is_staff:
        tickets = tickets.filter(reported_by=user)
    by_machine = defaultdict(list)
    for ticket in tickets.order_by('-reported_date', '-id'):
        by_machine[ticket.machine_id].append(ticket)

    down = {
        interval.machine_id: interval
        for interval in MachineDowntime.objects.filter(end_at__isnull=True).prefetch_related('phases')
    }
    days_90 = _clip_hours_by_machine(now)
    rows = []
    for machine in Machine.objects.filter(is_active=True).order_by('name'):
        machine_tickets = by_machine.get(machine.pk, [])
        interval = down.get(machine.pk)
        current_step = ''
        if interval:
            open_phase = next((p for p in interval.phases.all() if p.ended_at is None), None)
            current_step = open_phase.get_category_display() if open_phase else ''
        open_tickets = [t for t in machine_tickets if t.status in MaintenanceRecord.OPEN_STATUSES]
        if not machine_tickets and not interval:
            continue                                  # keep the board to machines with maintenance activity
        rows.append({
            'machine': machine,
            'stopped': bool(interval),
            'stopped_since': interval.start_at if interval else None,
            'stopped_for': fmt_duration((now - interval.start_at).total_seconds()) if interval else '',
            'current_step': current_step,
            'open_tickets': open_tickets,
            'tickets': machine_tickets,
            'downtime_90d': fmt_duration(days_90.get(machine.pk, 0)),
        })
    rows.sort(key=lambda r: (not r['stopped'], -len(r['open_tickets']), r['machine'].name))
    return rows


def _clip_hours_by_machine(now, days=90):
    window_start = now - timedelta(days=days)
    seconds = defaultdict(float)
    for interval in MachineDowntime.objects.filter(
        Q(end_at__isnull=True) | Q(end_at__gt=window_start), start_at__lte=now,
    ):
        seconds[interval.machine_id] += _clip(interval.start_at, interval.end_at or now, window_start, now)
    return seconds
