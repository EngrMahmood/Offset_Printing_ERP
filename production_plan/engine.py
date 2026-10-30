"""The planning engine: which job runs on which machine in which shift, for the next working days.

Rules (agreed with the planner):
  * Three lanes: Black (1 and 1+1 colour), Colour on GTO-size sheets, Colour on SM 74-size sheets.
    Each lane has its own machines (PlanSettings); GTO 2C and SM 74 never run black.
  * Capacity is 40,000 impressions per machine per shift (35,000 on a Friday day shift); a machine
    in preventive maintenance or blocked by the planner has none; Sundays and listed holidays are off.
  * FIFO: jobs are taken oldest order date first (a planner-set priority goes ahead of age), until the
    lane's free capacity for the horizon is full.
  * New jobs always go in a DAY shift, never a night shift. Day shifts run from small impression
    quantities up to larger ones; night shifts run from the largest down. A job bigger than one shift
    is cut into shift-sized chunks (which naturally fall to the nights).
  * Re-planning keeps everything already in the plan, removes only the jobs that have since been
    printed, and fills the freed capacity by the same rules - so a planner's changes always survive.
"""
from __future__ import annotations

import datetime
from collections import defaultdict
from dataclasses import dataclass, field

from django.db.models import Sum
from django.utils import timezone

from core.machine_routing import color_class, parse_sheet_size_mm
from core.models import Machine, Production

LANES = ('black', 'colour_gto', 'colour_sm')
LANE_LABELS = {
    'black': 'Black (1 and 1+1 colour)',
    'colour_gto': 'Colour - GTO sheet',
    'colour_sm': 'Colour - SM 74 sheet',
}
PLATE_MAKING_STAGES = ('repeat_plate_making', 'new_plate_making')
SHIFTS = ('D', 'N')
SHIFT_LABELS = {'D': 'Day', 'N': 'Night'}


# ----------------------------------------------------------------------------- calendar

def _split_names(text):
    return [part.strip() for part in str(text or '').replace('\n', ',').split(',') if part.strip()]


def off_dates(settings_obj):
    out = set()
    for line in _split_names(settings_obj.off_dates):
        try:
            out.add(datetime.date.fromisoformat(line))
        except ValueError:
            continue
    return out


def is_working_day(day, settings_obj):
    return day.weekday() != 6 and day not in off_dates(settings_obj)        # Sunday off


def working_dates_from(start, count, settings_obj):
    """``count`` working days starting on ``start`` (or the first working day after it)."""
    out, day = [], start
    while len(out) < count:
        if is_working_day(day, settings_obj):
            out.append(day)
        day += datetime.timedelta(days=1)
    return out


def next_working_day(after, settings_obj):
    return working_dates_from(after + datetime.timedelta(days=1), 1, settings_obj)[0]


def slot_capacity(day, shift, settings_obj):
    if not is_working_day(day, settings_obj):
        return 0
    if shift == 'D' and day.weekday() == 4:                                # Friday: lunch and Jumma prayer
        return int(settings_obj.friday_day_capacity)
    return int(settings_obj.shift_capacity)


def lane_machines(settings_obj):
    """{lane: [Machine, ...]} from the configured names (unknown names are skipped)."""
    by_lower = {m.name.strip().lower(): m for m in Machine.objects.filter(is_active=True)}
    config = {
        'black': settings_obj.black_machines,
        'colour_gto': settings_obj.colour_gto_machines,
        'colour_sm': settings_obj.colour_sm_machines,
    }
    return {lane: [by_lower[n.lower()] for n in _split_names(names) if n.lower() in by_lower] for lane, names in config.items()}


def blocked_slots(machines, dates, settings_obj):
    """{(machine_id, date, shift)} that cannot run: preventive maintenance dates plus planner blocks."""
    from production_plan.models import PlanBlock

    blocked = set()
    ids = [m.id for m in machines]
    try:
        from maintenance.models import PreventiveMaintenancePlan

        for plan in PreventiveMaintenancePlan.objects.filter(is_active=True, machine_id__in=ids, next_due_at__in=dates):
            for shift in (SHIFTS if settings_obj.pm_block_shift == 'B' else (settings_obj.pm_block_shift,)):
                blocked.add((plan.machine_id, plan.next_due_at, shift))
    except Exception:                                                        # the maintenance module is optional here
        pass
    for block in PlanBlock.objects.filter(machine_id__in=ids, date__in=dates):
        for shift in (SHIFTS if block.shift == 'B' else (block.shift,)):
            blocked.add((block.machine_id, block.date, shift))
    return blocked


# ----------------------------------------------------------------------------- jobs

@dataclass
class Unit:
    job_id: int
    jc: str
    sku: str
    lane: str
    order_date: datetime.date
    priority: int
    size: int
    is_new: bool
    is_chunk: bool
    remaining: int


def order_date_of(job, today):
    card = getattr(job, 'job_card', None)
    return job.po_approval_date or (card.po_date if card is not None and card.po_date else None) or today


def lane_of(job, settings_obj):
    classes = color_class(job.color_spec or '')
    if classes == 1:
        return 'black'
    if classes <= 0:
        return None
    size = parse_sheet_size_mm(job.print_sheet_size)
    long_mm, short_mm = float(settings_obj.gto_max_sheet_long_in) * 25.4, float(settings_obj.gto_max_sheet_short_in) * 25.4
    big = bool(size and (max(size) > long_mm + 1 or min(size) > short_mm + 1))
    return 'colour_sm' if big else 'colour_gto'


def load_jobs(settings_obj, today, only_held=False):
    """Every job that could be planned, as {job_id: info}, with its remaining impressions.

    Eligible: released or in production, plates ready, not on hold, not cut-and-pack, not opted out of
    machine planning. A combined-layout follower prints with its lead, so only the lead is planned.
    ``only_held=True`` returns the jobs that would be eligible but are on hold (the plan's "On hold" window)."""
    from planning.models import MergeGroupItem, PlanningJob, MERGE_GROUP_OPEN_STATUSES
    from reports.models import MachinePlanningJcSelection

    queryset = PlanningJob.objects.filter(is_active=True, status__in=('released', 'in_production')).exclude(job_process_type='cut_and_pack')
    queryset = queryset.filter(is_on_hold=True) if only_held else queryset.exclude(is_on_hold=True)
    jobs = list(queryset.select_related('job_card', 'hold_by'))
    excluded = set(MachinePlanningJcSelection.objects.filter(is_excluded=True).values_list('jc_number', flat=True))
    open_items = MergeGroupItem.objects.filter(merge_group__status__in=MERGE_GROUP_OPEN_STATUSES).select_related('merge_group')
    followers = {i.planning_job_id for i in open_items if not i.is_lead}
    lead_group = {i.planning_job_id: i.merge_group for i in open_items if i.is_lead}
    printed = {
        row['job_card__planning_job_id']: row['i'] or 0
        for row in Production.objects.filter(
            entry_type='printing', is_active=True, merge_parent__isnull=True, job_card__planning_job__in=[j.id for j in jobs],
        ).values('job_card__planning_job_id').annotate(i=Sum('impressions'))
    }
    info = {}
    for job in jobs:
        if job.jc_number in excluded or job.id in followers or (job.planning_stage or '') in PLATE_MAKING_STAGES:
            continue
        planned = job.planned_total_impressions or ((job.print_sheets or job.actual_sheet_required or 0) * max(job.print_passes or 1, 1))
        group = lead_group.get(job.id)
        if group is not None and group.combined_impressions():
            planned = group.combined_impressions()
        remaining = max(int(planned or 0) - printed.get(job.id, 0), 0)
        lane = lane_of(job, settings_obj)
        info[job.id] = {
            'job': job, 'lane': lane, 'remaining': remaining, 'order_date': order_date_of(job, today),
            'is_new': (job.repeat_flag or '').strip().lower() == 'new', 'priority': int(job.priority or 0),
        }
    return info


def make_units(info_by_job, job_ids, settings_obj, deduct=None):
    """Shift-sized units (a job above one shift is cut into chunks) for the given jobs."""
    cap = int(settings_obj.shift_capacity)
    units = defaultdict(list)
    for job_id in job_ids:
        info = info_by_job[job_id]
        remaining = info['remaining'] - (deduct or {}).get(job_id, 0)
        if remaining <= 0 or not info['lane']:
            continue
        left, chunked = remaining, remaining > cap
        while left > 0:
            part = min(left, cap)
            units[info['lane']].append(Unit(
                job_id=job_id, jc=info['job'].jc_number, sku=info['job'].sku or '', lane=info['lane'],
                order_date=info['order_date'], priority=info['priority'], size=part, is_new=info['is_new'],
                is_chunk=chunked, remaining=remaining,
            ))
            left -= part
    return units


def fifo_key(unit):
    return (-unit.priority, unit.order_date, unit.jc)


# ----------------------------------------------------------------------------- the build

@dataclass
class Slot:
    machine: object
    date: datetime.date
    shift: str
    cap: int
    load: int = 0

    @property
    def free(self):
        return max(self.cap - self.load, 0)


@dataclass
class Build:
    dates: list
    items: list = field(default_factory=list)          # dicts ready for PlanItem(**item)
    summary: dict = field(default_factory=dict)
    diff: dict = field(default_factory=dict)


def _place(units, layers, items, lane, seq_counter):
    """Place units in order; the first layer (shift) with room takes them, balanced across machines.

    A slot only counts as having room if it can take the whole unit, or at least a fifth of a shift of
    it, so a job is never cut into slivers to squeeze out the last few impressions. What does not fit
    anywhere is returned as leftover units."""
    leftovers = []
    for unit in units:
        left = unit.size
        while left > 0:
            slot = None
            for layer in layers:
                candidates = [s for s in layer if s.free >= min(left, max(s.cap // 5, 1)) and s.free > 0]
                if not candidates:
                    continue
                fits = [s for s in candidates if s.free >= left]
                slot = min(fits, key=lambda s: s.load) if fits else max(candidates, key=lambda s: s.free)
                break
            if slot is None:
                leftovers.append(Unit(**{**unit.__dict__, 'size': left}))
                break
            take = min(left, slot.free)
            key = (slot.machine.id, slot.date, slot.shift)
            seq_counter[key] += 1
            items.append({
                'lane': lane, 'machine_id': slot.machine.id, 'plan_date': slot.date, 'shift': slot.shift,
                'sequence': seq_counter[key], 'planning_job_id': unit.job_id, 'jc_number': unit.jc, 'sku': unit.sku,
                'order_date': unit.order_date, 'qty': take, 'remaining_at_plan': unit.remaining, 'priority': unit.priority,
                'is_new': unit.is_new, 'is_chunk': unit.is_chunk, 'pinned': False, 'manual': False, 'fifo_reason': '',
            })
            slot.load += take
            left -= take
    return leftovers


def build_plan(morning, settings_obj, base_items=None, today=None):
    """Build the plan for the horizon starting at ``morning``.

    ``base_items`` are the items (dicts) of the plan being carried forward - everything in them stays
    except jobs that are now printed or no longer eligible, and the freed capacity is refilled by the rules."""
    today = today or timezone.localdate()
    dates = working_dates_from(morning, int(settings_obj.horizon_working_days), settings_obj)
    machines_by_lane = lane_machines(settings_obj)
    all_machines = [m for ms in machines_by_lane.values() for m in ms]
    blocked = blocked_slots(all_machines, dates, settings_obj)
    info = load_jobs(settings_obj, today)

    build = Build(dates=dates)
    removed_printed, removed_other, reduced = [], [], []
    carried_by_job = defaultdict(int)
    seq = defaultdict(int)
    slots_by_lane = {}

    for lane, machines in machines_by_lane.items():
        slots_by_lane[lane] = [
            Slot(machine=m, date=d, shift=sh, cap=0 if (m.id, d, sh) in blocked else slot_capacity(d, sh, settings_obj))
            for d in dates for sh in SHIFTS for m in machines
        ]
    slot_index = {(s.machine.id, s.date, s.shift): s for slots in slots_by_lane.values() for s in slots}

    from planning.models import PlanningJob

    completed_ids = set(PlanningJob.objects.filter(
        pk__in=[i['planning_job_id'] for i in (base_items or []) if i.get('planning_job_id')], status='completed',
    ).values_list('pk', flat=True))

    # 1. carry what is already planned (planner changes included); drop only what is printed or no longer eligible
    for item in sorted(base_items or [], key=lambda i: (i['plan_date'], i['shift'], i['machine_id'], i['sequence'])):
        if item['plan_date'] not in dates:
            continue
        job_info = info.get(item['planning_job_id'])
        left = (job_info['remaining'] - carried_by_job[item['planning_job_id']]) if job_info else 0
        if job_info is None or not job_info['lane']:
            (removed_printed if item['planning_job_id'] in completed_ids else removed_other).append(item['jc_number'])
            continue
        if left <= 0:
            removed_printed.append(item['jc_number'])
            continue
        qty = min(item['qty'], left)
        if qty < item['qty']:
            reduced.append((item['jc_number'], item['qty'], qty))
        slot = slot_index.get((item['machine_id'], item['plan_date'], item['shift']))
        if slot is None:                                                   # a machine that left the lane
            removed_other.append(item['jc_number'])
            continue
        if slot.cap == 0 and not (item.get('pinned') or item.get('manual')):
            continue                                                       # slot now blocked: the job returns to the pool
        carried_by_job[item['planning_job_id']] += qty
        seq[(slot.machine.id, slot.date, slot.shift)] = max(seq[(slot.machine.id, slot.date, slot.shift)], item['sequence'])
        slot.load += qty
        kept = dict(item)
        kept.update(qty=qty, remaining_at_plan=job_info['remaining'], priority=job_info['priority'], is_new=job_info['is_new'])
        build.items.append(kept)

    # 2. fill the free capacity lane by lane
    pool = make_units(info, list(info), settings_obj, deduct=carried_by_job)
    lane_stats = {}
    overflow_new = []
    for lane, slots in slots_by_lane.items():
        units = sorted(pool.get(lane, []), key=fifo_key)
        day_slots = [s for s in slots if s.shift == 'D']
        night_slots = [s for s in slots if s.shift == 'N']
        day_free, night_free = sum(s.free for s in day_slots), sum(s.free for s in night_slots)

        new_units = [u for u in units if u.is_new]
        others = [u for u in units if not u.is_new]

        forced, used = [], 0
        for u in new_units:                                               # new jobs: day shifts only, in FIFO order
            if used >= day_free:
                overflow_new.append({'jc': u.jc, 'qty': u.size, 'lane': lane})
                continue
            forced.append(u)
            used += u.size
        budget = max(day_free - used, 0)

        window, acc = [], 0
        for u in others:                                                  # FIFO window sized to the remaining capacity
            if acc >= budget + night_free:
                break
            window.append(u)
            acc += u.size
        ascending = sorted(window, key=lambda u: (u.size, fifo_key(u)))
        day_others, day_acc = [], 0
        for u in ascending:                                               # the day takes the small end
            if day_acc >= budget:
                break
            day_others.append(u)
            day_acc += u.size
        taken = {id(u) for u in day_others}
        night_units = sorted([u for u in window if id(u) not in taken], key=lambda u: (-u.size, fifo_key(u)))

        day_units = sorted(forced + day_others, key=lambda u: (u.size, fifo_key(u)))
        day_layers = [[s for s in day_slots if s.date == d] for d in dates]
        night_layers = [[s for s in night_slots if s.date == d] for d in dates]
        day_left = _place(day_units, day_layers, build.items, lane, seq)
        for left_unit in day_left:                                          # a new job that does not fit never goes to night
            if left_unit.is_new:
                overflow_new.append({'jc': left_unit.jc, 'qty': left_unit.size, 'lane': lane})
        night_units = sorted(night_units + [u for u in day_left if not u.is_new], key=lambda u: (-u.size, fifo_key(u)))
        night_left = _place(night_units, night_layers, build.items, lane, seq)

        placed_ids = {(i['planning_job_id']) for i in build.items if i['lane'] == lane}
        planned_total = sum(i['qty'] for i in build.items if i['lane'] == lane)
        total_remaining = sum(info[j]['remaining'] for j in info if info[j]['lane'] == lane)
        lane_cap_day = sum(slot_capacity(d, sh, settings_obj) for d in dates for sh in SHIFTS) * len(machines_by_lane[lane]) or 1
        per_day_cap = lane_cap_day / max(len(dates), 1)
        lane_stats[lane] = {
            'machines': [m.name for m in machines_by_lane[lane]],
            'capacity': sum(s.cap for s in slots), 'planned': planned_total,
            'jobs_planned': len(placed_ids), 'remaining_total': total_remaining,
            'unplanned': max(total_remaining - planned_total, 0),
            'backlog_days_at_target': round(total_remaining / per_day_cap, 1) if per_day_cap else None,
            'unplaced': sum(u.size for u in night_left),
        }

    build.summary = {
        'lanes': lane_stats, 'overflow_new': overflow_new,
        'blocked': sorted([f'{machine_id}|{d.isoformat()}|{sh}' for machine_id, d, sh in blocked]),
        'no_colour_data': sum(1 for v in info.values() if v['lane'] is None and v['remaining'] > 0),
    }
    build.diff = {'removed_printed': removed_printed, 'removed_other': removed_other,
                  'reduced': [{'jc': jc, 'from': a, 'to': b} for jc, a, b in reduced]}
    build.items.sort(key=lambda i: (i['plan_date'], i['shift'], i['machine_id'], i['sequence']))
    return build
