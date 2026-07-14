from __future__ import annotations

import math
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

from django.db.models import Case, IntegerField, Prefetch, Sum, Value, When

from core.models import Dispatch, Material, Production
from planning.models import PLANNING_STATUS_ALIASES, PlanningJob, SkuRecipe

from .models import (
    RawMaterialSku,
    StockTransaction,
    display_material_name,
    normalize_material_name,
    normalize_purchase_sheet_size,
)

COMPLETED_STATUSES = {'completed', 'closed'}


def normalize_planning_status(raw_value):
    raw = (raw_value or '').strip().lower()
    return PLANNING_STATUS_ALIASES.get(raw, raw or 'draft')


def is_completed_planning_job(planning_job):
    return normalize_planning_status(planning_job.status) in COMPLETED_STATUSES


def _get_job_card(planning_job):
    try:
        return planning_job.job_card
    except Exception:
        return None


def _recipe_for_job(planning_job, recipe_by_sku):
    sku = (planning_job.sku or '').strip().lower()
    if not sku:
        return None
    return recipe_by_sku.get(sku)


def _purchase_sheet_size_for_job(planning_job, recipe=None):
    size = (planning_job.purchase_sheet_size or '').strip()
    if size:
        return normalize_purchase_sheet_size(size)
    if recipe and (recipe.purchase_sheet_size or '').strip():
        return normalize_purchase_sheet_size(recipe.purchase_sheet_size)
    return ''


def _material_name_for_job(planning_job, recipe=None):
    name = (planning_job.material or '').strip()
    if name:
        return name
    if recipe and (recipe.material or '').strip():
        return recipe.material.strip()
    return ''


def _purchase_sheet_ups_for_job(planning_job, recipe=None):
    if planning_job.purchase_sheet_ups is not None:
        return planning_job.purchase_sheet_ups
    if recipe and recipe.purchase_sheet_ups is not None:
        return recipe.purchase_sheet_ups
    return None


def _process_type_for_job(planning_job, recipe=None):
    if recipe and (recipe.job_process_type or '').strip():
        return recipe.job_process_type.strip()
    return (planning_job.job_process_type or 'print_and_pack').strip() or 'print_and_pack'


def _purchase_sheets_for_job(planning_job, recipe=None):
    """Mirror PlanningJob.purchase_sheet_required_display without N+1 recipe hits."""
    net_qty = planning_job.net_print_qty
    ups_value = planning_job.ups
    if ups_value is None and recipe and recipe.ups is not None:
        ups_value = recipe.ups

    sheets_required = None
    if net_qty is not None and ups_value:
        sheets_required = math.ceil(net_qty / ups_value) + (planning_job.wastage_sheets or 0)
    elif planning_job.print_sheets is not None:
        sheets_required = planning_job.print_sheets + (planning_job.wastage_sheets or 0)
    elif planning_job.actual_sheet_required is not None:
        sheets_required = planning_job.actual_sheet_required

    purchase_sheet_ups = _purchase_sheet_ups_for_job(planning_job, recipe)
    if sheets_required is not None and purchase_sheet_ups:
        return math.ceil(sheets_required / purchase_sheet_ups)
    return planning_job.purchase_sheet_required


def _prefetched_list(instance, related_name):
    cache = getattr(instance, '_prefetched_objects_cache', None) or {}
    if related_name in cache:
        return list(cache[related_name])
    return None


def _job_card_activity(job_card):
    """Packed / dispatched / print consumption from prefetched relations when possible."""
    if not job_card:
        return {
            'packed_pcs': 0,
            'dispatched_pcs': 0,
            'consumed_print_sheets': 0,
            'has_print_activity': False,
        }

    productions = _prefetched_list(job_card, 'productions')
    if productions is None:
        productions = list(job_card.productions.filter(is_active=True))

    packed_pcs = 0
    consumed_print_sheets = 0
    has_print = False
    for production in productions:
        if not production.is_active:
            continue
        if production.entry_type == 'packing':
            packed_pcs += int(production.packing_qty or 0)
        elif production.entry_type == 'printing':
            has_print = True
            consumed_print_sheets += int(production.output_sheets or 0) + int(production.waste_sheets or 0)

    dispatches = _prefetched_list(job_card, 'dispatch_set')
    if dispatches is None:
        dispatched_pcs = int(
            job_card.dispatch_set.filter(is_active=True).aggregate(total=Sum('dispatch_qty'))['total'] or 0
        )
    else:
        dispatched_pcs = sum(
            int(row.dispatch_qty or 0)
            for row in dispatches
            if row.is_active
        )

    return {
        'packed_pcs': packed_pcs,
        'dispatched_pcs': dispatched_pcs,
        'consumed_print_sheets': consumed_print_sheets,
        'has_print_activity': has_print,
    }


def _get_dispatched_pcs(planning_job, job_card):
    if job_card:
        return _job_card_activity(job_card)['dispatched_pcs']
    if not planning_job.pk:
        return 0
    runs = _prefetched_list(planning_job, 'dispatch_runs')
    if runs is None:
        total = planning_job.dispatch_runs.aggregate(total=Sum('delivered_qty'))['total']
        return int(total or 0)
    return sum(int(run.delivered_qty or 0) for run in runs)


def _get_packed_pcs(job_card):
    return _job_card_activity(job_card)['packed_pcs']


def _get_consumed_print_sheets(job_card):
    return _job_card_activity(job_card)['consumed_print_sheets']


def _has_print_activity(job_card):
    return _job_card_activity(job_card)['has_print_activity']


def _proportional_remaining(purchase_sheets, order_qty, remaining_pcs):
    if purchase_sheets is None:
        return None
    purchase_sheets = int(purchase_sheets)
    if not order_qty or order_qty <= 0:
        return purchase_sheets
    remaining_pcs = max(0, int(remaining_pcs))
    value = Decimal(purchase_sheets) * Decimal(remaining_pcs) / Decimal(order_qty)
    return int(value.quantize(Decimal('1'), rounding=ROUND_HALF_UP))


def compute_job_demand_sheets(planning_job, job_card=None, *, recipe=None, activity=None):
    """Compute remaining purchase-sheet demand for one planning job."""
    purchase_sheets = _purchase_sheets_for_job(planning_job, recipe)
    order_qty = int(planning_job.order_qty or 0)
    process_type = _process_type_for_job(planning_job, recipe)
    is_cut_pack = process_type == 'cut_and_pack'

    if activity is None:
        activity = _job_card_activity(job_card)

    packed_pcs = int(activity.get('packed_pcs') or 0)
    if job_card:
        dispatched_pcs = int(activity.get('dispatched_pcs') or 0)
    else:
        dispatched_pcs = _get_dispatched_pcs(planning_job, None)

    has_print = bool(activity.get('has_print_activity')) if not is_cut_pack else False
    consumed_print_sheets = int(activity.get('consumed_print_sheets') or 0) if has_print else 0
    purchase_sheet_ups = _purchase_sheet_ups_for_job(planning_job, recipe)

    result = {
        'purchase_sheets_planning': purchase_sheets,
        'order_qty': order_qty,
        'packed_pcs': packed_pcs,
        'dispatched_pcs': dispatched_pcs,
        'consumed_print_sheets': consumed_print_sheets,
        'is_cut_and_pack': is_cut_pack,
        'has_print_activity': has_print,
        'remaining_from_print': None,
        'remaining_from_pack': None,
        'remaining_from_dispatch': None,
        'job_demand_sheets': None,
        'is_incomplete': purchase_sheets is None,
    }

    if purchase_sheets is None:
        return result

    purchase_sheets = int(purchase_sheets)

    if not has_print and packed_pcs == 0 and dispatched_pcs == 0:
        result['job_demand_sheets'] = purchase_sheets
        return result

    if is_cut_pack:
        fulfilled_pcs = max(packed_pcs, dispatched_pcs)
        remaining_pcs = max(0, order_qty - fulfilled_pcs)
        result['job_demand_sheets'] = _proportional_remaining(
            purchase_sheets, order_qty, remaining_pcs
        )
        if packed_pcs > 0:
            result['remaining_from_pack'] = _proportional_remaining(
                purchase_sheets, order_qty, max(0, order_qty - packed_pcs)
            )
        if dispatched_pcs > 0:
            result['remaining_from_dispatch'] = _proportional_remaining(
                purchase_sheets, order_qty, max(0, order_qty - dispatched_pcs)
            )
        return result

    signals = []
    if has_print:
        if purchase_sheet_ups:
            consumed_purchase = math.ceil(consumed_print_sheets / int(purchase_sheet_ups))
        else:
            consumed_purchase = consumed_print_sheets
        remaining_from_print = max(0, purchase_sheets - consumed_purchase)
        result['remaining_from_print'] = remaining_from_print
        signals.append(remaining_from_print)

    if packed_pcs > 0:
        remaining_from_pack = _proportional_remaining(
            purchase_sheets, order_qty, max(0, order_qty - packed_pcs)
        )
        result['remaining_from_pack'] = remaining_from_pack
        signals.append(remaining_from_pack)

    if dispatched_pcs > 0:
        remaining_from_dispatch = _proportional_remaining(
            purchase_sheets, order_qty, max(0, order_qty - dispatched_pcs)
        )
        result['remaining_from_dispatch'] = remaining_from_dispatch
        signals.append(remaining_from_dispatch)

    result['job_demand_sheets'] = min(signals) if signals else purchase_sheets
    return result


def _preferred_material(existing, candidate):
    """Prefer mixed-case master names over ALL-CAPS / all-lowercase duplicates."""
    if existing is None:
        return candidate
    existing_name = existing.name or ''
    candidate_name = candidate.name or ''
    if existing_name.isupper() and not candidate_name.isupper():
        return candidate
    if existing_name.islower() and not candidate_name.islower():
        return candidate
    return existing


def _build_materials_by_name():
    materials_by_name = {}
    for material in Material.objects.all():
        key = normalize_material_name(material.name)
        if not key:
            continue
        materials_by_name[key] = _preferred_material(materials_by_name.get(key), material)
    return materials_by_name


def resolve_material_for_job(planning_job, job_card=None, *, materials_by_name=None, recipe=None):
    material_name = ''
    if job_card and job_card.material_id:
        material_name = job_card.material.name or ''
    if not material_name:
        material_name = _material_name_for_job(planning_job, recipe)
    key = normalize_material_name(material_name)
    if not key:
        return None
    if materials_by_name is not None:
        return materials_by_name.get(key)
    return Material.objects.filter(name__iexact=material_name.strip()).first()


def build_job_demand_row(
    planning_job,
    job_card=None,
    *,
    raw_skus_by_key=None,
    materials_by_name=None,
    recipe_by_sku=None,
):
    if job_card is None:
        job_card = _get_job_card(planning_job)

    recipe = _recipe_for_job(planning_job, recipe_by_sku or {})
    activity = _job_card_activity(job_card)
    demand = compute_job_demand_sheets(
        planning_job,
        job_card,
        recipe=recipe,
        activity=activity,
    )

    raw_material_name = ''
    if job_card and job_card.material_id:
        raw_material_name = job_card.material.name or ''
    if not raw_material_name:
        raw_material_name = _material_name_for_job(planning_job, recipe)

    material = resolve_material_for_job(
        planning_job,
        job_card,
        materials_by_name=materials_by_name,
        recipe=recipe,
    )
    purchase_sheet_size = _purchase_sheet_size_for_job(planning_job, recipe)
    material_key = normalize_material_name(
        material.name if material else raw_material_name
    )
    raw_sku = None
    if material_key and purchase_sheet_size and raw_skus_by_key is not None:
        raw_sku = raw_skus_by_key.get((material_key, purchase_sheet_size.lower()))

    if raw_sku:
        display_name = raw_sku.material.name
        material = raw_sku.material
    elif material:
        display_name = material.name
    else:
        display_name = display_material_name(raw_material_name) or '—'

    process_type = _process_type_for_job(planning_job, recipe)
    process_labels = dict(PlanningJob.JOB_PROCESS_TYPE_CHOICES)

    # Expose UPS factors so the template can render unit-conversion hints
    purchase_sheet_ups = _purchase_sheet_ups_for_job(planning_job, recipe)
    ups_value = planning_job.ups
    if ups_value is None and recipe and recipe.ups is not None:
        ups_value = recipe.ups

    # Pre-compute Purchase-Sheet equivalents for display (avoids template arithmetic)
    # consumed_print_sheets → PS equivalent: print_sheets / purchase_sheet_ups
    consumed_print_sheets = int(demand.get('consumed_print_sheets') or 0)
    packed_pcs = int(demand.get('packed_pcs') or 0)
    dispatched_pcs = int(demand.get('dispatched_pcs') or 0)
    order_qty = int(demand.get('order_qty') or 0)
    purchase_sheets_planning = int(demand.get('purchase_sheets_planning') or 0) if demand.get('purchase_sheets_planning') is not None else None

    # pcs_per_purchase_sheet: how many pieces fit in one purchase sheet
    pcs_per_purchase_sheet = None
    if ups_value and purchase_sheet_ups:
        pcs_per_purchase_sheet = int(ups_value) * int(purchase_sheet_ups)
    elif ups_value:
        pcs_per_purchase_sheet = int(ups_value)

    def _ps_equiv_print(sheets):
        """Convert print sheets → purchase sheets."""
        if not sheets:
            return None
        if purchase_sheet_ups and int(purchase_sheet_ups) > 0:
            return math.ceil(sheets / int(purchase_sheet_ups))
        return sheets

    def _ps_equiv_pcs(pcs):
        """Convert pieces → purchase sheets (proportional, rounded)."""
        if not pcs or not purchase_sheets_planning or not order_qty:
            return None
        val = Decimal(purchase_sheets_planning) * Decimal(pcs) / Decimal(order_qty)
        return int(val.quantize(Decimal('1'), rounding=ROUND_HALF_UP))

    return {
        'planning_job': planning_job,
        'job_card': job_card,
        'jc_number': planning_job.jc_number,
        'status': planning_job.workflow_status,
        'status_label': planning_job.workflow_status_label,
        'process_type': process_type,
        'process_label': process_labels.get(process_type, process_type),
        'material_name': display_name,
        'material_key': material_key,
        'purchase_sheet_size': purchase_sheet_size or '—',
        'material': material,
        'raw_material_sku': raw_sku,
        'supply_chain_item': raw_sku,
        'is_mapped': bool(raw_sku),
        'purchase_sheet_ups': purchase_sheet_ups,           # 1 purchase sheet = N print sheets
        'print_ups': ups_value,                              # 1 print sheet = N pcs
        'pcs_per_purchase_sheet': pcs_per_purchase_sheet,  # convenience: pcs in 1 purchase sheet
        # PS-equivalent of raw values (for parenthetical display in template)
        'consumed_print_sheets_ps': _ps_equiv_print(consumed_print_sheets),
        'packed_pcs_ps': _ps_equiv_pcs(packed_pcs),
        'dispatched_pcs_ps': _ps_equiv_pcs(dispatched_pcs),
        **demand,
    }


def parse_gap_filters(request_get):
    return {
        'plan_month': (request_get.get('plan_month') or '').strip(),
        'status': (request_get.get('status') or '').strip(),
        'process_type': (request_get.get('process_type') or '').strip(),
        'material_q': (request_get.get('material_q') or '').strip().lower(),
        'shortages_only': request_get.get('shortages_only') == '1',
        'exclude_zero_demand': request_get.get('exclude_zero_demand') == '1',
    }


def _planning_jobs_queryset(filters=None):
    qs = (
        PlanningJob.objects
        .exclude(status__in=COMPLETED_STATUSES)
        .select_related('job_card', 'job_card__material')
        .prefetch_related(
            Prefetch(
                'job_card__productions',
                queryset=Production.objects.filter(is_active=True).only(
                    'id',
                    'job_card_id',
                    'entry_type',
                    'is_active',
                    'output_sheets',
                    'waste_sheets',
                    'packing_qty',
                ),
            ),
            Prefetch(
                'job_card__dispatch_set',
                queryset=Dispatch.objects.filter(is_active=True).only(
                    'id',
                    'job_card_id',
                    'is_active',
                    'dispatch_qty',
                ),
            ),
            'dispatch_runs',
        )
        .order_by('jc_number')
    )

    if not filters:
        return qs

    if filters.get('plan_month'):
        qs = qs.filter(plan_month__iexact=filters['plan_month'])
    if filters.get('status'):
        qs = qs.filter(status=filters['status'])

    return qs


def _matches_job_filters(row, filters):
    if filters.get('process_type') and row['process_type'] != filters['process_type']:
        return False
    if filters.get('material_q'):
        needle = filters['material_q']
        haystack = f"{row['material_name']} {row.get('purchase_sheet_size', '')}".lower()
        if needle not in haystack:
            return False
    if filters.get('exclude_zero_demand') and int(row.get('job_demand_sheets') or 0) == 0:
        return False
    return True


def _gap_status(gap, is_mapped):
    if not is_mapped:
        return 'unmapped'
    if gap > 0:
        return 'shortage'
    if gap < 0:
        return 'surplus'
    return 'balanced'


def _build_recipe_by_sku(planning_jobs):
    wanted = {(job.sku or '').strip().lower() for job in planning_jobs if (job.sku or '').strip()}
    if not wanted:
        return {}
    recipe_by_sku = {}
    for recipe in SkuRecipe.objects.order_by('-updated_at').only(
        'sku',
        'material',
        'purchase_sheet_size',
        'purchase_sheet_ups',
        'job_process_type',
        'ups',
    ):
        key = (recipe.sku or '').strip().lower()
        if key in wanted and key not in recipe_by_sku:
            recipe_by_sku[key] = recipe
    return recipe_by_sku


def _on_hand_by_sku(sku_ids):
    result = {sku_id: 0 for sku_id in sku_ids}
    if not sku_ids:
        return result

    rows = (
        StockTransaction.objects
        .filter(raw_material_sku_id__in=sku_ids)
        .values('raw_material_sku_id')
        .annotate(
            opening=Sum(
                Case(
                    When(transaction_type='OPENING', then='sheet_qty_pcs'),
                    default=Value(0),
                    output_field=IntegerField(),
                )
            ),
            receiving=Sum(
                Case(
                    When(transaction_type='RECEIVING', then='sheet_qty_pcs'),
                    default=Value(0),
                    output_field=IntegerField(),
                )
            ),
            issuance=Sum(
                Case(
                    When(transaction_type='ISSUANCE', then='sheet_qty_pcs'),
                    default=Value(0),
                    output_field=IntegerField(),
                )
            ),
            adjustment=Sum(
                Case(
                    When(transaction_type='ADJUSTMENT', then='sheet_qty_pcs'),
                    default=Value(0),
                    output_field=IntegerField(),
                )
            ),
        )
    )
    for row in rows:
        result[row['raw_material_sku_id']] = (
            (row['opening'] or 0)
            + (row['receiving'] or 0)
            - (row['issuance'] or 0)
            + (row['adjustment'] or 0)
        )
    return result


def build_demand_gap_report(filters=None):
    filters = filters or {}
    planning_jobs = list(_planning_jobs_queryset(filters))
    recipe_by_sku = _build_recipe_by_sku(planning_jobs)
    materials_by_name = _build_materials_by_name()

    raw_skus = list(RawMaterialSku.objects.select_related('material').filter(is_active=True))
    raw_skus_by_key = {}
    for sku in raw_skus:
        material_key = normalize_material_name(sku.material.name)
        size_key = normalize_purchase_sheet_size(sku.purchase_sheet_size).lower()
        if material_key and size_key:
            # Prefer first SKU; later duplicates with same normalized key are ignored
            raw_skus_by_key.setdefault((material_key, size_key), sku)
    on_hand_by_sku = _on_hand_by_sku([sku.pk for sku in raw_skus])

    job_rows = []
    for planning_job in planning_jobs:
        row = build_job_demand_row(
            planning_job,
            raw_skus_by_key=raw_skus_by_key,
            materials_by_name=materials_by_name,
            recipe_by_sku=recipe_by_sku,
        )
        if row['is_incomplete'] or row['job_demand_sheets'] is None:
            continue
        if not _matches_job_filters(row, filters):
            continue
        job_rows.append(row)

    material_buckets = defaultdict(lambda: {
        'material': None,
        'material_name': '',
        'purchase_sheet_size': '',
        'raw_material_sku': None,
        'supply_chain_item': None,
        'on_hand': 0,
        'total_demand': 0,
        'planning_full_demand': 0,
        'print_job_count': 0,
        'cut_pack_job_count': 0,
        'job_count': 0,
        'jobs': [],
        'is_mapped': False,
    })

    unmapped_bucket = {
        'material_name': 'Unmapped material + purchase size',
        'purchase_sheet_size': '',
        'raw_material_sku': None,
        'supply_chain_item': None,
        'on_hand': None,
        'total_demand': 0,
        'planning_full_demand': 0,
        'print_job_count': 0,
        'cut_pack_job_count': 0,
        'job_count': 0,
        'jobs': [],
        'is_mapped': False,
        'gap': None,
        'gap_status': 'unmapped',
    }

    for row in job_rows:
        demand = int(row['job_demand_sheets'] or 0)
        planning_full = int(row['purchase_sheets_planning'] or 0)
        size_key = normalize_purchase_sheet_size(row.get('purchase_sheet_size')).lower()
        material_key = row.get('material_key') or normalize_material_name(row.get('material_name'))
        if row['raw_material_sku']:
            bucket_key = ('sku', row['raw_material_sku'].pk)
        elif material_key and size_key and size_key != '—':
            bucket_key = ('unmapped', material_key, size_key)
        else:
            bucket_key = None

        if bucket_key is None:
            bucket = unmapped_bucket
        elif bucket_key[0] == 'unmapped':
            bucket = material_buckets[bucket_key]
            if bucket['material'] is None:
                bucket['material'] = row['material']
                bucket['material_name'] = row['material_name']
                bucket['purchase_sheet_size'] = size_key or row['purchase_sheet_size']
        else:
            bucket = material_buckets[bucket_key]
            if bucket['material'] is None:
                bucket['material'] = row['material']
                bucket['material_name'] = row['material_name']
                bucket['purchase_sheet_size'] = (
                    normalize_purchase_sheet_size(row['raw_material_sku'].purchase_sheet_size)
                    or size_key
                    or row['purchase_sheet_size']
                )
                bucket['raw_material_sku'] = row['raw_material_sku']
                bucket['supply_chain_item'] = row['raw_material_sku']
                bucket['is_mapped'] = True
                bucket['on_hand'] = on_hand_by_sku.get(row['raw_material_sku'].pk, 0)

        bucket['total_demand'] += demand
        bucket['planning_full_demand'] += planning_full
        bucket['job_count'] += 1
        if row['is_cut_and_pack']:
            bucket['cut_pack_job_count'] += 1
        else:
            bucket['print_job_count'] += 1
        bucket['jobs'].append(row)

    material_rows = []
    for bucket in material_buckets.values():
        if bucket['is_mapped']:
            gap = bucket['total_demand'] - bucket['on_hand']
            bucket['gap'] = gap
            bucket['gap_status'] = _gap_status(gap, True)
        else:
            bucket['gap'] = None
            bucket['gap_status'] = 'unmapped'
        material_rows.append(bucket)

    unmapped_bucket['gap'] = None
    if unmapped_bucket['job_count']:
        material_rows.append(unmapped_bucket)

    material_rows.sort(key=lambda item: (
        0 if item.get('gap_status') == 'shortage' else 1,
        -(item.get('gap') or 0),
        item.get('material_name') or '',
        item.get('purchase_sheet_size') or '',
    ))

    if filters.get('shortages_only'):
        material_rows = [
            row for row in material_rows
            if row.get('gap_status') == 'shortage'
        ]

    summary = {
        'total_jobs': len(job_rows),
        'materials_tracked': sum(1 for row in material_rows if row.get('is_mapped')),
        'shortage_count': sum(1 for row in material_rows if row.get('gap_status') == 'shortage'),
        'total_shortage_sheets': sum(
            max(row.get('gap') or 0, 0)
            for row in material_rows
            if row.get('gap_status') == 'shortage'
        ),
        'unmapped_job_count': unmapped_bucket['job_count'],
        'unmapped_demand_sheets': unmapped_bucket['total_demand'],
        'incomplete_excluded': True,
    }

    return {
        'material_rows': material_rows,
        'job_rows': job_rows,
        'summary': summary,
        'filters': filters,
    }


def available_plan_months():
    return list(
        PlanningJob.objects
        .exclude(plan_month='')
        .exclude(status__in=COMPLETED_STATUSES)
        .values_list('plan_month', flat=True)
        .distinct()
        .order_by('plan_month')
    )
