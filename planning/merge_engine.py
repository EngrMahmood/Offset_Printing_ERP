"""Smart layout merge (ganging) engine.

Pure functions that detect PlanningJobs which can share one print sheet, and
allocate ups between them. No DB writes here so the logic stays unit testable.
"""

import math
from dataclasses import dataclass
from datetime import date
from itertools import combinations


@dataclass
class MergeConfig:
    size_tolerance_mm: int = 2
    qty_tolerance_pct: float = 5.0
    max_group_size: int = 6
    delivery_window_days: int = 0  # 0 = disabled
    setup_wastage_sheets_per_colour: int = 25


def normalise_material(value):
    return ' '.join((value or '').strip().lower().split())


def size_key(job, tol_mm):
    """Bucket key for the piece size, orientation-insensitive."""
    if job.size_w_mm is None or job.size_h_mm is None:
        return None
    long_side, short_side = sorted((job.size_w_mm, job.size_h_mm), reverse=True)
    step = max(tol_mm, 1)
    return (round(long_side / step), round(short_side / step))


def bucket_signature(job, cfg):
    """Rules 1-3: piece size, then material/sheet, then colours."""
    key = size_key(job, cfg.size_tolerance_mm)
    if key is None:
        return None
    if not normalise_material(job.material):
        return None
    if not (job.print_sheet_size or '').strip():
        return None
    return (
        key,
        normalise_material(job.material),
        (job.print_sheet_size or '').strip().lower(),
        (job.purchase_sheet_size or '').strip().lower(),
        job.front_colors or 0,
        job.back_colors or 0,
        job.total_colors or 0,
        job.print_passes or 0,
    )


def merge_blockers(job, cfg=None):
    """Why this job cannot be ganged yet, in words a planner can act on.

    Empty list means the job is mergeable. Only the geometry is required — a
    missing AWC is fine, since a new SKU's artwork is drawn into the combined
    layout rather than copied from an existing file.
    """
    cfg = cfg or MergeConfig()
    reasons = []

    if job.size_w_mm is None or job.size_h_mm is None:
        reasons.append('Piece size (W×H mm) missing on SKU master')
    if not normalise_material(job.material):
        reasons.append('Material missing')
    if not (job.print_sheet_size or '').strip():
        reasons.append('Print sheet size missing')
    if not job.ups_value:
        reasons.append('UPS missing — the ups split cannot be calculated')
    if not job.total_colors:
        reasons.append('Colour count missing')
    if not job.print_passes:
        reasons.append('No. of passes missing')
    if not job.net_print_qty:
        reasons.append('No open quantity left to produce')
    return reasons


def candidate_buckets(jobs, cfg):
    buckets = {}
    for job in jobs:
        signature = bucket_signature(job, cfg)
        if signature is None:
            continue
        if not job.ups_value or not job.net_print_qty:
            continue
        buckets.setdefault(signature, []).append(job)
    return {sig: group for sig, group in buckets.items() if len(group) > 1}


def _allocate_exact_fill(jobs, quantities, sheet_ups, cfg):
    """Exhaustive search for a whole-ups split that fills every ups of the sheet.

    The proportional heuristic can miss a valid split. For each candidate run
    length, every job's ups must lie between ceil(qty / run) (no under-production)
    and floor(qty * (1 + tolerance) / run) (no excess); if the bounds can add up
    to exactly ``sheet_ups`` the shortest such run wins. Ups above each job's
    minimum are handed out until the sheet is full — i.e. other jobs' ups are
    raised to take up the slack.
    """
    tol = 1 + cfg.qty_tolerance_pct / 100.0
    total = sum(quantities)
    first = max(1, int(total / sheet_ups))
    last = int(total * tol / sheet_ups) + 2
    for run_sheets in range(first, last + 1):
        lows = [math.ceil(q / run_sheets) for q in quantities]
        highs = [int(q * tol / run_sheets + 1e-9) for q in quantities]
        if any(lo > hi for lo, hi in zip(lows, highs)):
            continue
        if not (sum(lows) <= sheet_ups <= sum(highs)):
            continue
        ups = list(lows)
        spare = sheet_ups - sum(ups)
        while spare:
            # Give the next ups to the job it over-produces least.
            index = min(
                (i for i in range(len(ups)) if ups[i] < highs[i]),
                key=lambda i: (run_sheets * (ups[i] + 1) - quantities[i]) / quantities[i],
            )
            ups[index] += 1
            spare -= 1
        overages = [(run_sheets * u - q) / q * 100.0 for u, q in zip(ups, quantities)]
        return {
            'items': [
                {
                    'job': job,
                    'allocated_ups': u,
                    'planned_produced_qty': run_sheets * u,
                    'net_qty': q,
                    'overage_pct': round(o, 2),
                }
                for job, u, q, o in zip(jobs, ups, quantities, overages)
            ],
            'run_sheets': run_sheets,
            'sheet_ups': sheet_ups,
            'unused_ups': 0,
            'worst_overage_pct': round(max(overages), 2),
        }
    return None


def allocate_ups(jobs, sheet_ups, cfg, exhaustive=False):
    """Split `sheet_ups` between jobs so nobody is under-produced.

    Always fills the whole sheet. With ``exhaustive`` it also runs an exact search
    (slower, so the combinatorial suggestion builder leaves it off) that raises
    other jobs' ups where the heuristic could not. Returns None when the
    quantities cannot be reconciled within the tolerance.
    """
    quantities = [job.net_print_qty or 0 for job in jobs]
    if len(jobs) < 2 or sheet_ups < len(jobs) or any(q <= 0 for q in quantities):
        return None

    filled = _allocate_filling_sheet(jobs, quantities, sheet_ups, cfg)
    if filled or not exhaustive:
        return filled
    return _allocate_exact_fill(jobs, quantities, sheet_ups, cfg)


def _allocate_filling_sheet(jobs, quantities, sheet_ups, cfg):

    total_qty = sum(quantities)
    allocation = [max(1, int(round(sheet_ups * q / total_qty))) for q in quantities]

    # Reconcile the rounding drift one ups at a time, always moving the ups to
    # wherever it reduces the worst over-production.
    guard = 0
    while sum(allocation) != sheet_ups and guard < sheet_ups * 4:
        guard += 1
        if sum(allocation) > sheet_ups:
            # take one from the job with the most slack (lowest qty per ups)
            candidates = [i for i, u in enumerate(allocation) if u > 1]
            if not candidates:
                return None
            index = min(candidates, key=lambda i: quantities[i] / allocation[i])
            allocation[index] -= 1
        else:
            index = max(range(len(allocation)), key=lambda i: quantities[i] / allocation[i])
            allocation[index] += 1

    if sum(allocation) != sheet_ups:
        return None

    run_sheets = max(math.ceil(q / u) for q, u in zip(quantities, allocation))

    items = []
    worst_overage = 0.0
    for job, qty, ups in zip(jobs, quantities, allocation):
        produced = run_sheets * ups
        overage = (produced - qty) / qty * 100.0
        if overage < 0:
            return None
        if overage > cfg.qty_tolerance_pct:
            return None
        worst_overage = max(worst_overage, overage)
        items.append({
            'job': job,
            'allocated_ups': ups,
            'planned_produced_qty': produced,
            'net_qty': qty,
            'overage_pct': round(overage, 2),
        })

    return {
        'items': items,
        'run_sheets': run_sheets,
        'sheet_ups': sheet_ups,
        'unused_ups': 0,
        'worst_overage_pct': round(worst_overage, 2),
    }


def compute_savings(allocation, jobs, cfg=None):
    """Savings versus printing each SKU separately.

    The real win is the make-ready that no longer happens: its setup wastage
    (sheets per colour) and its setup time. Impressions barely move when each SKU
    already fills the sheet, so it is reported but is not the headline.
    """
    cfg = cfg or MergeConfig()
    total_colors = jobs[0].total_colors or 0
    passes = jobs[0].print_passes or 1
    separate_sheets = sum(job.calculated_sheets_required or 0 for job in jobs)
    wastage = max((job.wastage_sheets or 0) for job in jobs)
    merged_sheets = allocation['run_sheets'] + wastage
    makereadies_saved = len(jobs) - 1
    mr_minutes = max((job.total_mr_time_minutes or 0) for job in jobs)
    return {
        'plates_saved': makereadies_saved * total_colors,
        'makereadies_saved': makereadies_saved,
        'setup_sheets_saved': makereadies_saved * total_colors * cfg.setup_wastage_sheets_per_colour,
        'mr_minutes_saved': makereadies_saved * mr_minutes,
        'impressions_saved': max((separate_sheets - merged_sheets) * passes, 0),
        'merged_sheets': merged_sheets,
    }


def _within_delivery_window(jobs, cfg):
    if not cfg.delivery_window_days:
        return True
    dates = [job.delivery_date for job in jobs if job.delivery_date]
    if len(dates) < len(jobs):
        return True
    return (max(dates) - min(dates)).days <= cfg.delivery_window_days


# Exhaustive combinations(bucket, size) is C(n, size) — fine for the handful of
# jobs a signature usually collects, but a bucket that grows into the dozens
# (a common size/material/colour combo just keeps matching new jobs) makes this
# combinatorially explosive: C(58, 6) alone is 40M+ subsets, which is what made
# /planning/merge/ hang. Cap how many jobs any one combinations() call ever sees
# by chunking an oversized bucket into windows (sorted by delivery date, so jobs
# likely to actually ship together stay grouped) instead of searching the whole
# bucket at once. Buckets at or under this size (the overwhelming majority) are
# untouched — one window covers the whole bucket, same as before.
MAX_COMBINATION_WINDOW = 18


def _bucket_windows(bucket, window_size):
    if len(bucket) <= window_size:
        yield bucket
        return
    ordered = sorted(bucket, key=lambda job: job.delivery_date or date.max)
    for start in range(0, len(ordered), window_size):
        yield ordered[start:start + window_size]


def build_suggestions(jobs, cfg=None):
    """Return non-overlapping merge suggestions, best savings first."""
    cfg = cfg or MergeConfig()
    suggestions = []

    for signature, bucket in candidate_buckets(jobs, cfg).items():
        sheet_ups = bucket[0].ups_value
        for window in _bucket_windows(bucket, MAX_COMBINATION_WINDOW):
            max_size = min(cfg.max_group_size, len(window))
            for size in range(max_size, 1, -1):
                for subset in combinations(window, size):
                    if not _within_delivery_window(subset, cfg):
                        continue
                    allocation = allocate_ups(list(subset), sheet_ups, cfg)
                    if not allocation:
                        continue
                    suggestion = {
                        'signature': signature,
                        'jobs': list(subset),
                        'job_ids': sorted(job.id for job in subset),
                        'allocation': allocation,
                        'savings': compute_savings(allocation, list(subset), cfg),
                    }
                    suggestions.append(suggestion)

    suggestions.sort(
        key=lambda s: (
            s['savings']['setup_sheets_saved'],
            s['savings']['plates_saved'],
            -s['allocation']['worst_overage_pct'],
        ),
        reverse=True,
    )

    chosen = []
    used = set()
    for suggestion in suggestions:
        if used.intersection(suggestion['job_ids']):
            continue
        used.update(suggestion['job_ids'])
        chosen.append(suggestion)
    return chosen
