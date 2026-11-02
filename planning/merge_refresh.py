"""Bring an open combined-layout group back in line with its jobs' current quantities.

A merge group freezes each member's net quantity (order qty less carried-forward stock) when it is accepted.
If a member's order or stock then changes - for example a missing earlier entry is entered and the carried
forward stock is cleared - the group's run sheets and quantities are out of date. The ups split is part of
the approved layout, so it is kept; only the net quantities, the run length and the planned quantities move.
"""
from __future__ import annotations

import math

from django.db import transaction

from planning.models import MERGE_GROUP_OPEN_STATUSES, MergeGroup


def printing_started(group):
    from core.models import Production

    job_ids = list(group.items.values_list('planning_job_id', flat=True))
    return Production.objects.filter(
        entry_type='printing', is_active=True, job_card__planning_job_id__in=job_ids,
    ).exists()


def current_numbers(group):
    """What the group should say today, keeping the approved ups: {'run_sheets', 'items': [...]} or None."""
    items = list(group.items.select_related('planning_job'))
    nets = {item.id: int(item.planning_job.net_print_qty or 0) for item in items}
    if not items or any(q <= 0 or not item.allocated_ups for item, q in ((i, nets[i.id]) for i in items)):
        return None
    run_sheets = max(math.ceil(nets[item.id] / item.allocated_ups) for item in items)
    rows = []
    for item in items:
        net = nets[item.id]
        planned = run_sheets * item.allocated_ups
        rows.append({
            'item': item, 'net_qty': net, 'planned_produced_qty': planned,
            'overage_pct': round((planned - net) / net * 100.0, 2),
        })
    return {'run_sheets': run_sheets, 'items': rows}


def stale_changes(group):
    """Human-readable differences between the stored group and today's job quantities (empty = up to date)."""
    if group.status not in MERGE_GROUP_OPEN_STATUSES or printing_started(group):
        return []                                  # once printing has started the group is history, not out of date
    fresh = current_numbers(group)
    if not fresh:
        return []
    changes = []
    if (group.run_sheets or 0) != fresh['run_sheets']:
        changes.append(f"Run sheets {group.run_sheets or 0:,} -> {fresh['run_sheets']:,}")
    for row in fresh['items']:
        if (row['item'].net_qty or 0) != row['net_qty']:
            changes.append(f"{row['item'].planning_job.jc_number}: net quantity {row['item'].net_qty or 0:,} -> {row['net_qty']:,}")
    return changes


def refresh_group(group):
    """Update the group from its jobs. Returns (status, message): 'updated', 'unchanged', 'locked' or 'skipped'."""
    if group.status not in MERGE_GROUP_OPEN_STATUSES:
        return 'skipped', f'{group.code} is not open.'
    if printing_started(group):
        return 'locked', f'{group.code} has printing entries, so its quantities are no longer changed.'
    changes = stale_changes(group)
    if not changes:
        return 'unchanged', f'{group.code} already matches its jobs.'
    fresh = current_numbers(group)
    with transaction.atomic():
        group.run_sheets = fresh['run_sheets']
        group.save(update_fields=['run_sheets'])
        for row in fresh['items']:
            item = row['item']
            item.net_qty = row['net_qty']
            item.planned_produced_qty = row['planned_produced_qty']
            item.overage_pct = row['overage_pct']
            item.save(update_fields=['net_qty', 'planned_produced_qty', 'overage_pct'])
    worst = max(r['overage_pct'] for r in fresh['items'])
    warn = f' Over-production of {worst:g}% on the combined sheet - check the ups split.' if worst > 10 else ''
    return 'updated', f"{group.code}: {'; '.join(changes)}.{warn}"


def refresh_open_groups():
    return [(group, *refresh_group(group)) for group in MergeGroup.objects.filter(status__in=MERGE_GROUP_OPEN_STATUSES).order_by('code')]
