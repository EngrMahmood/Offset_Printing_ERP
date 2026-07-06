"""Analyze repeat/new and pending SKU impact on live DB data."""
import os
import sys

sys.path.insert(0, os.getcwd())
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')

import django

django.setup()

from collections import Counter, defaultdict

from django.db.models import Count

from planning.models import PlanningJob, PoDocument, SkuRecipe
from workflow.services import (
    _annotate_items_with_recipe,
    _build_recipe_map,
    _collect_pending_sku_rows,
    _po_payload_items,
    _sku_key,
)

ADVANCED_STATUSES = {
    'pending_qc',
    'qc_approved',
    'released',
    'in_production',
    'completed',
    'planning_approved',
}
ADVANCED_STAGES = {
    'plate_received',
    'new_plate_making',
    'repeat_plate_making',
    'planning_done',
    'in_production',
}


def recipe_is_bulk_like(recipe):
    if not recipe:
        return False
    filled = sum(
        1
        for field in ('material', 'color_spec', 'ups', 'print_sheet_size', 'job_name')
        if str(getattr(recipe, field, '') or '').strip()
    )
    return filled >= 3


def yesterday_repeat(sku, recipe_map):
    key = _sku_key(sku)
    return bool(key and key in recipe_map)


def production_repeat(sku, exclude_po=None):
    key = _sku_key(sku)
    if not key:
        return False
    qs = PlanningJob.objects.filter(sku__iexact=sku, is_active=True)
    if exclude_po:
        qs = qs.exclude(po_number__iexact=exclude_po)
    for job in qs.only('status', 'planning_stage'):
        status = (job.status or '').lower()
        stage = (job.planning_stage or '').strip()
        if status in ADVANCED_STATUSES or stage in ADVANCED_STAGES:
            return True
    return False


def legacy_repeat(sku, recipe_map):
    recipe = recipe_map.get(_sku_key(sku))
    return recipe_is_bulk_like(recipe)


def concurrent_po_repeat(sku, po_number, po_created_at, open_po_skus):
    key = _sku_key(sku)
    if not key:
        return False
    for other_po, created_at, other_sku in open_po_skus:
        if other_sku != key:
            continue
        if other_po == _normalize_po(po_number):
            continue
        if created_at < po_created_at:
            return True
    return False


def _normalize_po(po):
    return (po or '').strip().upper()


def proposed_repeat(sku, po_number, po_created_at, recipe_map, open_po_skus):
    if legacy_repeat(sku, recipe_map):
        return 'legacy'
    if production_repeat(sku, exclude_po=po_number):
        return 'prior_production'
    if concurrent_po_repeat(sku, po_number, po_created_at, open_po_skus):
        return 'concurrent_po'
    return 'new'


def main():
    print('=== DATASET OVERVIEW ===')
    active_jobs = PlanningJob.objects.filter(is_active=True)
    print('PlanningJob total:', PlanningJob.objects.count())
    print('PlanningJob active:', active_jobs.count())
    print('SkuRecipe total:', SkuRecipe.objects.count())
    print('PoDocument:', PoDocument.objects.exclude(extracted_payload__isnull=True).count())

    print('\n=== JOB STATUS (active) ===')
    for row in active_jobs.values('status').annotate(c=Count('id')).order_by('-c'):
        print(f"  {row['status'] or '(blank)'}: {row['c']}")

    print('\n=== REPEAT_FLAG (active) ===')
    for row in active_jobs.values('repeat_flag').annotate(c=Count('id')).order_by('-c'):
        print(f"  {row['repeat_flag'] or '(blank)'}: {row['c']}")

    print('\n=== RECIPE STATUS ===')
    for row in SkuRecipe.objects.values('master_data_status').annotate(c=Count('id')).order_by('-c'):
        print(f"  {row['master_data_status']}: {row['c']}")

    all_recipes = list(SkuRecipe.objects.all())
    recipe_by_key = {r.sku.upper(): r for r in all_recipes}
    bulk_like = sum(1 for r in all_recipes if recipe_is_bulk_like(r))
    print(f'\nRecipes likely from sheet bulk (3+ core fields): {bulk_like}/{len(all_recipes)}')

    # Build open PO sku index
    open_po_skus = []
    po_docs = PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('created_at', 'id')
    for doc in po_docs:
        payload = doc.extracted_payload or {}
        po_number = payload.get('po_number') or ''
        items = _po_payload_items(payload)
        for item in items:
            sku = (item.get('sku') or '').strip()
            if sku:
                open_po_skus.append((_normalize_po(po_number), doc.created_at, _sku_key(sku)))

    print('\n=== CLASSIFICATION COMPARISON ON ACTIVE JOBS ===')
    counters = Counter()
    stored_mismatch = []
    advanced_impact = []

    for job in active_jobs.iterator():
        recipe_map = {_sku_key(job.sku): recipe_by_key.get(_sku_key(job.sku))}
        if recipe_by_key.get(_sku_key(job.sku)):
            recipe_map[_sku_key(job.sku)] = recipe_by_key[_sku_key(job.sku)]

        y = yesterday_repeat(job.sku, {_sku_key(job.sku): recipe_by_key.get(_sku_key(job.sku))})
        p_reason = proposed_repeat(job.sku, job.po_number, None, recipe_map, open_po_skus)
        p = p_reason != 'new'
        stored = (job.repeat_flag or '').strip().lower() == 'repeat'

        key = (
            f"yesterday={'Repeat' if y else 'New'}",
            f"proposed={'Repeat' if p else 'New'}({p_reason})",
            f"stored={'Repeat' if stored else 'New'}",
        )
        counters[key] += 1

        if stored != p:
            stored_mismatch.append((job.po_number, job.sku, job.repeat_flag, p_reason, job.status, job.planning_stage))

        adv = (job.status or '').lower() in ADVANCED_STATUSES or (job.planning_stage or '') in ADVANCED_STAGES
        if adv and y != p:
            advanced_impact.append((job.po_number, job.sku, y, p_reason, job.repeat_flag, job.status, job.planning_stage))

    print('Top classification tuples (count):')
    for key, count in counters.most_common(12):
        print(f'  {count:4d} | {key[0]} | {key[1]} | {key[2]}')

    print(f'\nActive jobs where STORED repeat_flag != proposed: {len(stored_mismatch)}')
    for row in stored_mismatch[:15]:
        print(' ', row)

    print(f'\nAdvanced-phase jobs where yesterday != proposed: {len(advanced_impact)}')
    for row in advanced_impact[:15]:
        print(' ', row)

    print('\n=== CONCURRENT PO: SAME SKU ON MULTIPLE OPEN POS ===')
    sku_pos = defaultdict(list)
    for po, created_at, sku_key in open_po_skus:
        sku_pos[sku_key].append((po, created_at))
    concurrent = {sku: pos for sku, pos in sku_pos.items() if len({p for p, _ in pos}) > 1}
    print(f'SKUs on multiple POs in DB: {len(concurrent)}')
    for sku, pos in list(sorted(concurrent.items(), key=lambda x: -len(x[1])))[:12]:
        unique_pos = sorted({p for p, _ in pos})
        print(f'  {sku}: {len(unique_pos)} POs -> {unique_pos[:4]}')

    print('\n=== PENDING SKU QUEUE COMPARISON ===')
    docs = list(PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('-created_at')[:400])
    current_rows = _collect_pending_sku_rows(docs)
    current_skus = {_sku_key(r['sku']) for r in current_rows}

    yesterday_pending = set()
    for doc in docs:
        payload = doc.extracted_payload or {}
        items = _po_payload_items(payload)
        if not items:
            continue
        recipe_map = _build_recipe_map(items)
        _, _, _, missing = _annotate_items_with_recipe(items, recipe_map)
        for sku in missing:
            yesterday_pending.add(_sku_key(sku))

  # proposed pending = unapproved master
    proposed_pending = set()
    for doc in docs:
        payload = doc.extracted_payload or {}
        items = _po_payload_items(payload)
        recipe_map = _build_recipe_map(items)
        ignored = {_sku_key(s) for s in (payload.get('new_skus_ignored') or []) if s}
        seen = set()
        for item in items:
            sku = (item.get('sku') or '').strip()
            key = _sku_key(sku)
            if not key or key in ignored or key in seen:
                continue
            seen.add(key)
            recipe = recipe_map.get(key)
            approved = bool(recipe and recipe.master_data_status == 'approved')
            if not approved:
                proposed_pending.add(key)

    print(f'Yesterday pending (no recipe): {len(yesterday_pending)}')
    print(f'Current pending (unapproved, fixed): {len(current_skus)}')
    print(f'Proposed pending (unapproved): {len(proposed_pending)}')
    only_current = current_skus - yesterday_pending
    only_yesterday = yesterday_pending - current_skus
    print(f'In current but not yesterday: {len(only_current)}')
    for sku in list(sorted(only_current))[:10]:
        r = recipe_by_key.get(sku)
        print(f'  {sku} -> recipe={r.master_data_status if r else "missing"}')
    print(f'In yesterday but not current: {len(only_yesterday)}')
    for sku in list(sorted(only_yesterday))[:10]:
        print(f'  {sku}')

    print('\n=== JOBS IN ADVANCED PHASE WITH DRAFT/UNAPPROVED RECIPE ===')
    count = 0
    for job in active_jobs.iterator():
        recipe = recipe_by_key.get(_sku_key(job.sku))
        if recipe and recipe.master_data_status != 'approved':
            adv = (job.status or '').lower() in ADVANCED_STATUSES or (job.planning_stage or '') in ADVANCED_STAGES
            if adv or (job.planning_stage or '') == 'plate_received':
                count += 1
                if count <= 20:
                    print(
                        f"  {job.po_number} | {job.sku[:50]} | recipe={recipe.master_data_status}"
                        f" | flag={job.repeat_flag} | status={job.status} | stage={job.planning_stage}"
                    )
    print(f'Total such jobs: {count}')


if __name__ == '__main__':
    main()
