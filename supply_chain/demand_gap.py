from __future__ import annotations

import math
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

from django.db.models import Prefetch, Sum

from core.jobcard_service import _resolve_by_name
from core.models import Material, Production
from planning.models import PLANNING_STATUS_ALIASES, PlanningJob

from .models import RawMaterialSku, normalize_purchase_sheet_size
from .raw_material_sku import resolve_raw_material_sku_for_planning_job
from .services import build_dashboard_data

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


def _get_dispatched_pcs(planning_job, job_card):
    if job_card:
        return int(job_card.total_dispatch or 0)
    if not planning_job.pk:
        return 0
    total = planning_job.dispatch_runs.aggregate(total=Sum('delivered_qty'))['total']
    return int(total or 0)


def _get_packed_pcs(job_card):
    if not job_card:
        return 0
    return int(job_card.total_packed_pcs or 0)


def _get_consumed_print_sheets(job_card):
    if not job_card:
        return 0
    totals = job_card.printing_productions.aggregate(
        total_output=Sum('output_sheets'),
        total_waste=Sum('waste_sheets'),
    )
    return int(totals['total_output'] or 0) + int(totals['total_waste'] or 0)


def _has_print_activity(job_card):
    if not job_card:
        return False
    return job_card.printing_productions.exists()


def _proportional_remaining(purchase_sheets, order_qty, remaining_pcs):
    if purchase_sheets is None:
        return None
    purchase_sheets = int(purchase_sheets)
    if not order_qty or order_qty <= 0:
        return purchase_sheets
    remaining_pcs = max(0, int(remaining_pcs))
    value = Decimal(purchase_sheets) * Decimal(remaining_pcs) / Decimal(order_qty)
    return int(value.quantize(Decimal('1'), rounding=ROUND_HALF_UP))


def compute_job_demand_sheets(planning_job, job_card=None):
    """Compute remaining purchase-sheet demand for one planning job."""
    purchase_sheets = planning_job.purchase_sheet_required_display
    order_qty = int(planning_job.order_qty or 0)
    is_cut_pack = planning_job.is_cut_and_pack()

    packed_pcs = _get_packed_pcs(job_card)
    dispatched_pcs = _get_dispatched_pcs(planning_job, job_card)
    has_print = _has_print_activity(job_card) if not is_cut_pack else False
    consumed_print_sheets = _get_consumed_print_sheets(job_card) if has_print else 0
    purchase_sheet_ups = planning_job.purchase_sheet_ups_display

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


def resolve_material_for_job(planning_job, job_card=None):
    if job_card and job_card.material_id:
        return job_card.material
    material_name = planning_job.material_display or (planning_job.material or '').strip()
    return _resolve_by_name(Material, material_name)


def build_job_demand_row(planning_job, job_card=None, raw_skus_by_key=None):
    if job_card is None:
        job_card = _get_job_card(planning_job)

    demand = compute_job_demand_sheets(planning_job, job_card)
    raw_sku, material, purchase_sheet_size = resolve_raw_material_sku_for_planning_job(planning_job, job_card)
    if raw_sku is None and raw_skus_by_key and material and purchase_sheet_size:
        raw_sku = raw_skus_by_key.get((material.pk, purchase_sheet_size.lower()))

    return {
        'planning_job': planning_job,
        'job_card': job_card,
        'jc_number': planning_job.jc_number,
        'status': planning_job.workflow_status,
        'status_label': planning_job.workflow_status_label,
        'process_type': planning_job.job_process_type_display,
        'process_label': planning_job.job_process_type_label,
        'material_name': material.name if material else (planning_job.material_display or planning_job.material or '—'),
        'purchase_sheet_size': purchase_sheet_size or (planning_job.purchase_sheet_size_display or '—'),
        'material': material,
        'raw_material_sku': raw_sku,
        'supply_chain_item': raw_sku,
        'is_mapped': bool(raw_sku),
        **demand,
    }


def parse_gap_filters(request_get):
    return {
        'plan_month': (request_get.get('plan_month') or '').strip(),
        'status': (request_get.get('status') or '').strip(),
        'process_type': (request_get.get('process_type') or '').strip(),
        'material_q': (request_get.get('material_q') or '').strip().lower(),
        'shortages_only': request_get.get('shortages_only') == '1',
    }


def _planning_jobs_queryset(filters=None):
    qs = (
        PlanningJob.objects
        .exclude(status__in=COMPLETED_STATUSES)
        .select_related('job_card', 'job_card__material')
        .prefetch_related(
            Prefetch(
                'job_card__productions',
                queryset=Production.objects.filter(is_active=True),
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
    return True


def _gap_status(gap, is_mapped):
    if not is_mapped:
        return 'unmapped'
    if gap > 0:
        return 'shortage'
    if gap < 0:
        return 'surplus'
    return 'balanced'


def build_demand_gap_report(filters=None):
    filters = filters or {}
    raw_skus = RawMaterialSku.objects.select_related('material').filter(is_active=True)
    raw_skus_by_key = {
        (sku.material_id, normalize_purchase_sheet_size(sku.purchase_sheet_size).lower()): sku
        for sku in raw_skus
    }
    on_hand_by_sku = {
        row['item'].pk: row['closing']
        for row in build_dashboard_data(raw_skus)
    }

    job_rows = []
    for planning_job in _planning_jobs_queryset(filters):
        row = build_job_demand_row(
            planning_job,
            raw_skus_by_key=raw_skus_by_key,
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
        if row['raw_material_sku']:
            bucket_key = row['raw_material_sku'].pk
        elif row['material'] and row.get('purchase_sheet_size') and row['purchase_sheet_size'] != '—':
            bucket_key = ('unmapped', row['material'].pk, row['purchase_sheet_size'].lower())
        else:
            bucket_key = None

        if bucket_key is None or (isinstance(bucket_key, tuple) and bucket_key[0] == 'unmapped'):
            if isinstance(bucket_key, tuple):
                partial = material_buckets[bucket_key]
                if partial['material'] is None:
                    partial['material'] = row['material']
                    partial['material_name'] = row['material_name']
                    partial['purchase_sheet_size'] = row['purchase_sheet_size']
                bucket = partial
            else:
                bucket = unmapped_bucket
        else:
            bucket = material_buckets[bucket_key]
            if bucket['material'] is None:
                bucket['material'] = row['material']
                bucket['material_name'] = row['material_name']
                bucket['purchase_sheet_size'] = row['purchase_sheet_size']
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
