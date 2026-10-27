"""Flexo PO/WO intake: parse an uploaded PO/WO PDF and turn its lines into
Flexo planning jobs. PDF parsing reuses planning.po_extractor read-only (the
offset module is not modified); everything after parsing is Flexo-specific."""
import re
from datetime import date
from decimal import Decimal, InvalidOperation

from django.core.exceptions import ValidationError
from django.db import transaction

from core.models import Department

from . import services
from .models import FlexoPlanningJob, FlexoPoDocument, FlexoSkuRecipe

_PIECE_UNITS = {'PIECE', 'PCS', 'PC', 'NOS', 'EA', 'EACH', 'UNIT'}
_ROLL_UNITS = {'ROLL', 'RL'}


def sku_key(value):
    return re.sub(r'[^A-Z0-9]', '', str(value or '').upper())


def map_unit(raw):
    unit = str(raw or '').strip().upper().rstrip('.')
    if unit in _ROLL_UNITS:
        return 'roll'
    if unit in _PIECE_UNITS:
        return 'piece'
    return None


def _money(value):
    """PO amounts arrive as floats (1.24); keep them as clean 4-dp Decimals."""
    if value in (None, ''):
        return None
    try:
        return Decimal(str(value)).quantize(Decimal('0.0001'))
    except InvalidOperation:
        return None


def _parse_iso(value):
    try:
        return date.fromisoformat(str(value)[:10])
    except (TypeError, ValueError):
        return None


def extract_pdf(file_obj):
    from planning.po_extractor import extract_po_from_pdf
    return extract_po_from_pdf(file_obj)


@transaction.atomic
def save_extracted_document(payload, *, actor):
    po_number = str((payload or {}).get('po_number') or '').strip()
    if not po_number:
        raise ValidationError('No PO/WO number could be read from this PDF.')
    if not (payload.get('items') or []):
        raise ValidationError('No line items could be read from this PDF.')
    doc, _ = FlexoPoDocument.objects.update_or_create(
        po_number=po_number,
        defaults={'document_type': payload.get('document_type') or '', 'payload': payload, 'uploaded_by': actor},
    )
    return doc


def _recipe_lookup():
    """normalized sku/job-name -> recipe, approved ones winning over the rest."""
    lookup = {}
    rank = {'approved': 0, 'reviewed': 1, 'pending_review': 2, 'draft': 3}
    for recipe in sorted(FlexoSkuRecipe.objects.filter(is_active=True), key=lambda r: rank.get(r.master_data_status, 9)):
        for key in (sku_key(recipe.sku), sku_key(recipe.job_name)):
            if key and key not in lookup:
                lookup[key] = recipe
    return lookup


def analyze_lines(doc):
    """One dict per PO line: matched SKU master, state, unit and any existing job."""
    lookup = _recipe_lookup()
    po_number = doc.po_number
    rows = []
    for index, item in enumerate(doc.items):
        recipe = lookup.get(sku_key(item.get('sku'))) or lookup.get(sku_key(item.get('job_name')))
        unit = map_unit(item.get('unit'))
        existing = None
        if recipe:
            existing = FlexoPlanningJob.objects.filter(po_number=po_number, sku_recipe=recipe).first()
        if existing:
            state = 'has_job'
        elif recipe is None:
            state = 'missing'
        elif not recipe.is_approved:
            state = 'unapproved'
        elif unit is None:
            state = 'bad_unit'
        else:
            state = 'ready'
        rows.append({
            'index': index, 'item': item, 'recipe': recipe, 'state': state,
            'unit': unit, 'existing_job': existing,
            'qty': int(round(float(item.get('quantity') or 0))),
        })
    return rows


def _destination(payload):
    raw = str(payload.get('delivery_location') or '').strip()
    return re.sub(r'^UTOPIA INDUSTRIES\s+', '', raw, flags=re.I)[:100]


def _refresh_status(doc, rows_after):
    if all(r['state'] == 'has_job' for r in rows_after):
        doc.status = 'processed'
        doc.save(update_fields=['status', 'updated_at'])


@transaction.atomic
def create_jobs_from_document(doc, line_indexes, *, actor):
    rows = {r['index']: r for r in analyze_lines(doc)}
    payload = doc.payload or {}
    department = Department.objects.filter(name__iexact=str(payload.get('department') or '').strip()).first()
    received = _parse_iso(payload.get('approval_date')) or _parse_iso(payload.get('po_date'))
    created, skipped = [], []
    for index in line_indexes:
        row = rows.get(index)
        if not row or row['state'] != 'ready':
            skipped.append(index)
            continue
        item = row['item']
        created.append(services.create_job_from_recipe(
            row['recipe'], actor=actor, po_number=doc.po_number, po_received_date=received,
            delivery_date=_parse_iso(item.get('delivery_date')), order_qty=row['qty'],
            destination=_destination(payload), unit_type=row['unit'],
            unit_cost=_money(item.get('unit_cost')), department=department,
            remarks=str(item.get('remarks') or '')[:255],
        ))
    _refresh_status(doc, analyze_lines(doc))
    return created, skipped


@transaction.atomic
def create_draft_recipes(doc, *, actor):
    """Create Draft SKU masters for PO lines whose SKU is not in the master yet."""
    created = []
    for row in analyze_lines(doc):
        if row['state'] != 'missing':
            continue
        item = row['item']
        sku = str(item.get('sku') or '').strip()
        if not sku or FlexoSkuRecipe.objects.filter(sku__iexact=sku).exists():
            continue
        created.append(FlexoSkuRecipe.objects.create(
            sku=sku[:255], job_name=str(item.get('job_name') or sku)[:255],
            unit_type=row['unit'] or 'piece', remarks=str(item.get('remarks') or '')[:255],
            default_unit_cost=_money(item.get('unit_cost')), created_by=actor,
        ))
    return created
