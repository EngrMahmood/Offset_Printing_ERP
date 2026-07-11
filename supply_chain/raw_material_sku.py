from __future__ import annotations

from decimal import Decimal, InvalidOperation

from django.db import transaction

from core.jobcard_service import _resolve_by_name
from core.models import Material

from .models import RawMaterialSku, normalize_purchase_sheet_size


def resolve_raw_material_sku(material, purchase_sheet_size, *, active_only=True):
    if not material:
        return None
    normalized_size = normalize_purchase_sheet_size(purchase_sheet_size)
    if not normalized_size:
        return None

    qs = RawMaterialSku.objects.filter(material_id=material.pk)
    if active_only:
        qs = qs.filter(is_active=True)

    for candidate in qs.select_related('material'):
        if normalize_purchase_sheet_size(candidate.purchase_sheet_size).lower() == normalized_size.lower():
            return candidate
    return None


def resolve_raw_material_sku_for_planning_job(planning_job, job_card=None):
    material = None
    if job_card and job_card.material_id:
        material = job_card.material
    if material is None:
        material_name = planning_job.material_display or (planning_job.material or '').strip()
        material = _resolve_by_name(Material, material_name)

    purchase_sheet_size = planning_job.purchase_sheet_size_display or planning_job.purchase_sheet_size or ''
    return resolve_raw_material_sku(material, purchase_sheet_size), material, normalize_purchase_sheet_size(purchase_sheet_size)


def resolve_raw_material_sku_for_job_card(job_card):
    if not job_card or not job_card.material_id:
        return None
    purchase_sheet_size = job_card.purchase_sheet_size or ''
    if not purchase_sheet_size and job_card.planning_job_id:
        purchase_sheet_size = job_card.planning_job.purchase_sheet_size_display or ''
    return resolve_raw_material_sku(job_card.material, purchase_sheet_size)


def _parse_bool(value, default=True):
    text = str(value or '').strip().lower()
    if not text:
        return default
    return text in {'1', 'true', 'yes', 'y', 'active'}


def _parse_decimal(value, default=Decimal('0')):
    text = str(value or '').strip()
    if not text:
        return default
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return default


def _parse_int(value, default=0):
    text = str(value or '').strip()
    if not text:
        return default
    try:
        return int(float(text))
    except (TypeError, ValueError):
        return default


@transaction.atomic
def upsert_raw_material_sku_row(row_data):
    sku = (row_data.get('sku') or '').strip()
    material_name = (row_data.get('material_name') or '').strip()
    purchase_sheet_size = normalize_purchase_sheet_size(row_data.get('purchase_sheet_size'))
    errors = []

    if not sku:
        errors.append('Raw Material SKU is required.')
    if not material_name:
        errors.append('Material Name is required.')
    if not purchase_sheet_size:
        errors.append('Purchase Sheet Size is required.')
    if errors:
        return None, errors, False

    material = Material.objects.filter(name__iexact=material_name).first()
    if material is None:
        material = Material.objects.create(name=material_name)

    defaults = {
        'sku': sku,
        'uom': (row_data.get('uom') or 'Sheets').strip() or 'Sheets',
        'sheet_packing_pcs': _parse_int(row_data.get('sheet_packing_pcs'), 1),
        'unit_cost': _parse_decimal(row_data.get('unit_cost')),
        'safety_stock': _parse_int(row_data.get('safety_stock'), 0),
        'max_stock_level': _parse_int(row_data.get('max_stock_level'), 10000),
        'lead_time_days': _parse_int(row_data.get('lead_time_days'), 1),
        'is_active': _parse_bool(row_data.get('is_active'), True),
    }

    existing_by_pair = None
    for candidate in RawMaterialSku.objects.filter(material=material):
        if normalize_purchase_sheet_size(candidate.purchase_sheet_size).lower() == purchase_sheet_size.lower():
            existing_by_pair = candidate
            break
    existing_by_sku = RawMaterialSku.objects.filter(sku__iexact=sku).first()

    if existing_by_sku and existing_by_pair and existing_by_sku.pk != existing_by_pair.pk:
        return None, [f'SKU "{sku}" and material/size pair point to different records.'], False

    target = existing_by_pair or existing_by_sku
    created = target is None
    if target:
        for field, value in defaults.items():
            setattr(target, field, value)
        target.material = material
        target.purchase_sheet_size = purchase_sheet_size
        target.save()
        return target, [], created

    obj = RawMaterialSku.objects.create(
        material=material,
        purchase_sheet_size=purchase_sheet_size,
        **defaults,
    )
    return obj, [], True


def import_raw_material_skus(rows):
    created = 0
    updated = 0
    skipped = 0
    errors = []

    for index, row in enumerate(rows, start=2):
        obj, row_errors, was_created = upsert_raw_material_sku_row(row)
        if row_errors:
            skipped += 1
            errors.append(f'Row {index}: {"; ".join(row_errors)}')
            continue
        if obj is None:
            skipped += 1
            continue
        if was_created:
            created += 1
        else:
            updated += 1

    return {
        'created': created,
        'updated': updated,
        'skipped': skipped,
        'errors': errors,
    }


def list_material_choices():
    return list(Material.objects.order_by('name').values_list('name', flat=True))
