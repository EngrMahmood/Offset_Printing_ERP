"""Explode approved SkuBoms against planned jobs into a raw material requirement report.

For each PlanningJob, the job's SkuRecipe is reached the same way the rest of
the ERP does it — the cached, case-insensitive `sku_recipe` lookup on
PlanningJob (planning/models.py) — then its approved current SkuBom's lines
are scaled by the job's own driver values and summed per RawItem.
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from .models import RawItem, SkuBom

DRIVER_FIELD_BY_BASIS = {
    'per_1000_pcs': ('order_qty', Decimal(1000)),
    'per_1000_sheets': ('print_sheets', Decimal(1000)),
    'per_1000_impressions': ('planned_total_impressions', Decimal(1000)),
}


def _driver_qty_for_line(line, job):
    basis_code = line.source_template_line.quantity_basis.code if line.source_template_line_id else None
    if basis_code in DRIVER_FIELD_BY_BASIS:
        field_name, divisor = DRIVER_FIELD_BY_BASIS[basis_code]
        raw_value = getattr(job, field_name, None) or 0
        return Decimal(raw_value) / divisor
    if basis_code == 'per_job':
        return Decimal(1)
    if basis_code == 'per_plate_set':
        return Decimal(1) if (job.plate_set_no or '').strip() else Decimal(0)
    # Fallback: scale by order quantity per 1000 pieces — the most common case.
    return Decimal(job.order_qty or 0) / Decimal(1000)


def explode_requirements(jobs_queryset):
    """Returns a list of dicts, one per RawItem, aggregated across the given jobs.

    Keys: raw_item, required_qty, on_hand, safety_stock, shortfall, unit_cost,
    extended_cost, lead_time_days, suggested_order_qty, job_count.
    """
    required_by_item = defaultdict(Decimal)
    jobs_by_item = defaultdict(int)

    boms_by_sku_id = {}

    for job in jobs_queryset:
        recipe = job.sku_recipe
        if recipe is None:
            continue
        bom = boms_by_sku_id.get(recipe.pk)
        if bom is None and recipe.pk not in boms_by_sku_id:
            bom = SkuBom.objects.filter(
                sku_recipe=recipe, is_current=True, status='approved', is_active=True,
            ).prefetch_related('lines__raw_item', 'lines__source_template_line__quantity_basis').first()
            boms_by_sku_id[recipe.pk] = bom
        if bom is None:
            continue

        for line in bom.lines.all():
            qty_per_1000 = line.gross_qty  # already includes wastage, "per FG unit" scale factor handled below
            driver_units = _driver_qty_for_line(line, job)
            required_by_item[line.raw_item_id] += qty_per_1000 * driver_units
            jobs_by_item[line.raw_item_id] += 1

    if not required_by_item:
        return []

    from supply_chain.demand_gap import _on_hand_by_sku

    raw_items = {r.pk: r for r in RawItem.objects.filter(pk__in=required_by_item.keys()).select_related('category', 'uom')}
    bridged_rm_ids = {
        raw_item.raw_material_sku_id: raw_item.pk
        for raw_item in raw_items.values() if raw_item.raw_material_sku_id
    }
    on_hand_by_rm_sku = _on_hand_by_sku(list(bridged_rm_ids.keys())) if bridged_rm_ids else {}

    rows = []
    for item_id, required_qty in required_by_item.items():
        raw_item = raw_items.get(item_id)
        if raw_item is None:
            continue
        on_hand = None
        if raw_item.raw_material_sku_id:
            on_hand = Decimal(on_hand_by_rm_sku.get(raw_item.raw_material_sku_id, 0) or 0)
        shortfall = None
        if on_hand is not None:
            shortfall = max(Decimal(0), required_qty + raw_item.safety_stock - on_hand)
        extended_cost = (required_qty * raw_item.unit_cost)
        suggested_order_qty = shortfall
        if suggested_order_qty is not None and raw_item.pack_size and raw_item.pack_size > 0:
            packs = (suggested_order_qty / raw_item.pack_size).to_integral_value(rounding='ROUND_CEILING')
            suggested_order_qty = packs * raw_item.pack_size
        if suggested_order_qty is not None and raw_item.moq and suggested_order_qty > 0:
            suggested_order_qty = max(suggested_order_qty, raw_item.moq)

        rows.append({
            'raw_item': raw_item,
            'required_qty': required_qty,
            'on_hand': on_hand,
            'safety_stock': raw_item.safety_stock,
            'shortfall': shortfall,
            'unit_cost': raw_item.unit_cost,
            'extended_cost': extended_cost,
            'lead_time_days': raw_item.lead_time_days,
            'suggested_order_qty': suggested_order_qty,
            'job_count': jobs_by_item[item_id],
        })

    rows.sort(key=lambda r: (r['shortfall'] is None, -(r['shortfall'] or 0)))
    return rows


def raise_item_requests_from_shortfalls(shortfall_rows, user, request_type, department):
    """Create one supply_chain.ItemRequest per shortfall row via the existing
    request/numbering path, so IR-ID sequencing and the approval chain stay
    identical to every other item request in the system.

    ``shortfall_rows`` is a list of the dicts explode_requirements() returns
    (or a subset the caller selected). Returns the created ItemRequest list.
    """
    from django.db import transaction

    from supply_chain.item_request_service import generate_request_no
    from supply_chain.models import ItemRequest

    created = []
    with transaction.atomic():
        for row in shortfall_rows:
            raw_item = row['raw_item']
            qty = row.get('suggested_order_qty') or row.get('shortfall') or Decimal(0)
            if qty <= 0:
                continue
            item_request = ItemRequest.objects.create(
                request_type=request_type,
                item_title=f'{raw_item.name} ({raw_item.item_code})',
                uom=raw_item.uom.code,
                specifications=raw_item.specification or raw_item.name,
                required_quantity=qty,
                department=department,
                existing_sku=raw_item.raw_material_sku,
                estimated_unit_price=raw_item.unit_cost,
                raised_by=user,
            )
            generate_request_no(item_request)
            created.append(item_request)
    return created
