import csv
import base64
import io
import json
import re
from difflib import SequenceMatcher
from datetime import datetime
from decimal import Decimal, InvalidOperation
from urllib.parse import urlencode

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Count, Q, Sum
from django.http import HttpResponse
from django.utils import timezone
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse

from core.jc_numbering import allocate_next_jc_number
from core.views import permission_required
from .forms import PlanningJobEditForm, SkuRecipeForm
from .models import PlanningDispatchRun, PlanningJob, PlanningPrintRun, PoDocument, SkuRecipe
from .po_extractor import extract_po_from_pdf


PLANNING_STATUSES = [
    ('draft', 'Draft'),
    ('reviewed', 'QC Approved'),
    ('approved', 'Production Manager Approved'),
    ('closed', 'Closed'),
]
PLANNING_STATUS_SET = {value for value, _ in PLANNING_STATUSES}
NEW_SKU_REQUIREMENT_NOTE = 'NEW SKU: Shade matching and setup verification required before production run.'
COST_MISMATCH_NOTE_PREFIX = 'COST ALERT:'
DEPARTMENT_MISMATCH_NOTE_PREFIX = 'DEPARTMENT ALERT:'
SKU_MASTER_APPROVAL_REQUIRED_FIELDS = [
    ('job_name', 'Job Name'),
    ('material', 'Material'),
    ('color_spec', 'Color'),
    ('application', 'Application'),
    ('machine_name', 'Machine'),
    ('print_sheet_size', 'Print Sheet'),
    ('purchase_sheet_size', 'Purchase Sheet'),
    ('ups', 'UPS'),
    ('purchase_material', 'Purchase Material Origin'),
]

_COLOR_PLUS_RE = re.compile(r'^(\d+)\s*\+\s*(\d+)$')
_COLOR_SINGLE_RE = re.compile(r'^(\d+)\s*(?:colou?r(?:s)?)?$', re.IGNORECASE)


def build_planning_readme_text():
    return """Offset ERP - Planning Module Easy Guide

Last Updated: 2026-04-19

=============================
1) MASTER SKU (STEP 1)
=============================
Purpose:
- Keep approved SKU master data ready before routing PO jobs.

How to use:
- Create or bulk upload SKU master data.
- Save as Draft first.
- Move Draft -> Reviewed -> Approved.
- Only approved recipes are used as final master data.

Required fields for approval:
- Job Name, Material, Color, Application, Machine
- Print Sheet, Purchase Sheet, UPS, Purchase Material Origin

=============================
2) PO INTAKE (STEP 2)
=============================
Purpose:
- Upload PO and split lines into Repeat and New.

Routing rule:
- Repeat lines -> Planning Jobs
- New lines -> Pending SKU Master Data

Important notes:
- Duplicate SKU lines in one PO are merged.
- Qty display is normalized (trailing decimals removed).

=============================
3) PENDING NEW SKU (STEP 3)
=============================
Purpose:
- Complete missing master data for new SKU lines from PO.

Rules:
- Job Name comes from PO and is not manual.
- Department and unit cost can prefill from PO when available.
- Application must be one of: UV, Lamination Gloss, Lamination Matt, NO.
- Purchase Material Origin must be Local or Imported.

Approval path:
- Save Draft -> Send For Approval -> Approved
- Approved new SKU records auto-sync to planning jobs.

=============================
4) PLANNING JOBS (STEP 4)
=============================
Purpose:
- Create/update, review, and manage production planning jobs.

Input source:
- Repeat jobs from PO intake
- Approved new SKUs after master approval

Operational controls:
- Filter by PO/SKU/status/department/machine/date.
- Bulk status update for selected rows.
- Open detail, edit, print A4 job card.

=============================
5) APPROVAL QUEUE (STEP 5)
=============================
Purpose:
- Release jobs through QC and Production Manager checkpoints.

Status transitions:
- Draft -> QC Approved
- QC Approved -> Production Manager Approved

After approval:
- Print job card and run shop-floor execution flow.

=============================
6) SHOP FLOOR EXECUTION (STEP 6)
=============================
Purpose:
- Use QR scan and A4 card for execution traceability.

Use:
- Open job via scan.
- Track run/dispatch logs from planning-linked records.

=============================
7) CHANGE & MISMATCH ALERTS
=============================
Repeat route behavior:
- If PO cost differs from master cost, a COST ALERT note is attached.
- If PO department differs from master department, a DEPARTMENT ALERT note is attached.

New route behavior:
- New SKU records must complete master approval before planning sync.

=============================
8) DAILY DISCIPLINE
=============================
- Keep master records clean and approved.
- Route every PO through intake queue.
- Resolve pending SKUs same day.
- Complete approvals before production release.
"""


@login_required
@permission_required('can_edit_jobcard')
def planning_readme(request):
    return render(
        request,
        'planning/planning_readme.html',
        {'generated_on': timezone.now()},
    )


@login_required
@permission_required('can_edit_jobcard')
def download_planning_readme(request):
    content = build_planning_readme_text()
    response = HttpResponse(content, content_type='text/plain; charset=utf-8')
    response['Content-Disposition'] = 'attachment; filename="planning_workflow_guide.txt"'
    return response


def _clean_number(raw_value):
    if raw_value is None:
        return None
    text = str(raw_value).strip().replace(',', '')
    if not text:
        return None
    return text


def _to_int(raw_value):
    cleaned = _clean_number(raw_value)
    if cleaned is None:
        return None
    try:
        return int(float(cleaned))
    except ValueError:
        return None


def _to_decimal(raw_value):
    cleaned = _clean_number(raw_value)
    if cleaned is None:
        return None
    try:
        return Decimal(cleaned)
    except (InvalidOperation, ValueError):
        return None


def _format_display_qty(raw_value):
    """Render quantities without trailing decimals (e.g., 1000.0 -> 1000)."""
    value = _to_decimal(raw_value)
    if value is None:
        return raw_value if raw_value not in (None, '') else '-'

    if value == value.to_integral_value():
        return str(int(value))

    normalized = value.normalize()
    text = format(normalized, 'f').rstrip('0').rstrip('.')
    return text or '0'


def _normalize_color_spec_input(raw_value):
    value = (raw_value or '').strip()
    if not value:
        return ''

    plus_match = _COLOR_PLUS_RE.fullmatch(value)
    if plus_match:
        return f"{int(plus_match.group(1))}+{int(plus_match.group(2))}"

    single_match = _COLOR_SINGLE_RE.fullmatch(value)
    if single_match:
        return f"{int(single_match.group(1))} color"

    return value


def _normalize_application_input(raw_value):
    value = (raw_value or '').strip()
    if not value:
        return ''
    lowered = value.lower()
    if lowered in {'uv', 'uv coating'}:
        return 'UV'
    if lowered in {'lamination gloss', 'gloss lamination', 'gloss'}:
        return 'Lamination Gloss'
    if lowered in {'lamination matt', 'matt lamination', 'matte lamination', 'matt', 'matte'}:
        return 'Lamination Matt'
    if lowered in {'no', 'none', 'n/a', 'na', 'nil'}:
        return 'NO'
    return value


def _append_unique_note_line(base_text, line):
    text = str(base_text or '').strip()
    line = str(line or '').strip()
    if not line:
        return text

    lines = [part.strip() for part in text.splitlines() if part.strip()]
    if line in lines:
        return '\n'.join(lines)
    return '\n'.join(lines + [line]) if lines else line


def _build_cost_mismatch_note(master_cost, po_cost):
    master = _to_decimal(master_cost)
    po = _to_decimal(po_cost)
    if master is None or po is None:
        return ''
    if master == po:
        return ''
    return f"{COST_MISMATCH_NOTE_PREFIX} PO unit cost {po} differs from master default {master}. PO cost is applied to this job."


def _build_department_mismatch_note(master_department, po_department):
    master = (master_department or '').strip()
    po = (po_department or '').strip()
    if not master or not po:
        return ''
    if master.lower() == po.lower():
        return ''
    return f"{DEPARTMENT_MISMATCH_NOTE_PREFIX} PO department '{po}' differs from master department '{master}'."


def _to_date(raw_value):
    if not raw_value:
        return None
    text = str(raw_value).strip()
    if not text:
        return None

    for fmt in ('%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _parse_iso_date(raw_value):
    if not raw_value:
        return None
    try:
        return datetime.strptime(str(raw_value).strip(), '%Y-%m-%d').date()
    except ValueError:
        return None


def _normalize_status(raw_value, default='draft'):
    value = (raw_value or '').strip().lower()
    if value in {'open', 'pending'}:
        return 'draft'
    if value in PLANNING_STATUS_SET:
        return value
    return default


def _parse_date_filter(raw_value):
    value = (raw_value or '').strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, '%Y-%m-%d').date()
    except ValueError:
        return None


def _build_qr_image_base64(data):
    if not data:
        return ''
    try:
        import qrcode
    except ImportError:
        return ''

    qr = qrcode.QRCode(border=2, box_size=3)
    qr.add_data(data)
    qr.make(fit=True)
    image = qr.make_image(fill_color='black', back_color='white')
    buffer = io.BytesIO()
    image.save(buffer, format='PNG')
    return base64.b64encode(buffer.getvalue()).decode('ascii')


def _sku_key(sku):
    return (sku or '').strip().upper()


def _missing_required_master_fields(recipe, fallback_job_name=''):
    missing = []
    if not recipe:
        fallback = (fallback_job_name or '').strip()
        return [
            label
            for field, label in SKU_MASTER_APPROVAL_REQUIRED_FIELDS
            if not (field == 'job_name' and fallback)
        ]

    for field, label in SKU_MASTER_APPROVAL_REQUIRED_FIELDS:
        value = getattr(recipe, field, None)
        if isinstance(value, str):
            if not value.strip():
                missing.append(label)
        elif value is None:
            missing.append(label)
    return missing


def _sync_new_sku_requirement(existing_requirement, is_new):
    """Ensure NEW SKU requirement note exists only for New jobs."""
    lines = [line.strip() for line in str(existing_requirement or '').splitlines() if line.strip()]
    filtered_lines = [line for line in lines if line != NEW_SKU_REQUIREMENT_NOTE]

    if is_new:
        return '\n'.join([NEW_SKU_REQUIREMENT_NOTE] + filtered_lines)
    return '\n'.join(filtered_lines)


def _build_recipe_map(items):
    sku_values = sorted({_sku_key(item.get('sku')) for item in items if item.get('sku')})
    if not sku_values:
        return {}

    recipe_query = Q()
    for sku in sku_values:
        recipe_query |= Q(sku__iexact=sku)

    recipes = SkuRecipe.objects.filter(recipe_query, master_data_status='approved')
    return {recipe.sku.upper(): recipe for recipe in recipes}


def _to_optional_positive_int(raw_value):
    value = _to_int(raw_value)
    if value is None:
        return None
    return value if value >= 0 else None


def _to_optional_decimal(raw_value):
    value = _to_decimal(raw_value)
    if value is None:
        return None
    return value if value >= 0 else None


def _sanitize_po_payload_items(payload):
    """Normalize payload items for workflow screens.

    Applies SKU-level deduplication and respects expected line count when available
    to avoid noisy extra rows from fallback parsers.
    """
    items, _ = _deduplicate_po_items_by_sku((payload or {}).get('items', []))

    # Merge OCR-near-duplicate SKUs when qty/date match and text is almost identical.
    consolidated = []
    for item in items:
        sku = (item.get('sku') or '').strip()
        qty = _to_int(item.get('quantity'))
        ddate = (item.get('delivery_date') or '').strip()
        sku_norm = ''.join(ch for ch in sku.upper() if ch.isalnum())
        merged = False
        for existing in consolidated:
            ex_sku = (existing.get('sku') or '').strip()
            ex_qty = _to_int(existing.get('quantity'))
            ex_ddate = (existing.get('delivery_date') or '').strip()
            ex_norm = ''.join(ch for ch in ex_sku.upper() if ch.isalnum())
            similar = SequenceMatcher(a=sku_norm, b=ex_norm).ratio() >= 0.985
            if similar and qty == ex_qty and ddate == ex_ddate:
                merged = True
                break
        if not merged:
            consolidated.append(item)
    items = consolidated

    expected_line_count = _to_int((payload or {}).get('expected_line_count'))
    if expected_line_count and expected_line_count > 0 and len(items) > expected_line_count:
        items = items[:expected_line_count]
    return items


def _annotate_items_with_recipe(items, recipe_map):
    annotated = []
    repeat_count = 0
    new_count = 0
    missing_skus = []

    for item in items:
        sku = (item.get('sku') or '').strip()
        key = _sku_key(sku)
        has_recipe = bool(key and key in recipe_map)
        item_copy = dict(item)
        item_copy['is_repeat'] = has_recipe
        item_copy['recipe_status'] = 'Repeat' if has_recipe else 'New'
        annotated.append(item_copy)

        if has_recipe:
            repeat_count += 1
        else:
            new_count += 1
            if sku:
                missing_skus.append(sku)

    return annotated, repeat_count, new_count, sorted(set(missing_skus))


def _deduplicate_po_items_by_sku(items):
    """Ensure one row per SKU in a PO payload by merging duplicate SKU lines."""
    merged = {}
    order = []
    duplicate_skus = set()

    for item in items:
        item_copy = dict(item)
        sku = (item_copy.get('sku') or '').strip()
        sku_key = _sku_key(sku)
        if not sku_key:
            continue

        if sku_key not in merged:
            merged[sku_key] = item_copy
            order.append(sku_key)
            continue

        duplicate_skus.add(sku)
        existing = merged[sku_key]

        # Never add quantities from duplicate lines; keep first captured qty.
        # This avoids qty inflation when the same PO is uploaded multiple times.
        existing_qty = _to_int(existing.get('quantity'))
        current_qty = _to_int(item_copy.get('quantity'))
        if existing_qty is None and current_qty is not None:
            existing['quantity'] = current_qty

        # Fill empty values from later duplicate rows when useful.
        for field in ['job_name', 'delivery_date', 'unit', 'unit_cost', 'net_total']:
            if not existing.get(field) and item_copy.get(field):
                existing[field] = item_copy.get(field)

    deduped = [merged[key] for key in order]
    for idx, item in enumerate(deduped, start=1):
        item['line_no'] = idx
    return deduped, sorted(duplicate_skus)


def _history_repeat_new_counts(items):
    """Classify Repeat/New from historical PlanningJob SKU existence."""
    sku_keys = {_sku_key(item.get('sku')) for item in items if item.get('sku')}
    existing_any_jobs_skus = set()
    if sku_keys:
        sku_any_query = Q()
        for sku_key in sku_keys:
            sku_any_query |= Q(sku__iexact=sku_key)
        existing_any_jobs_skus = {
            _sku_key(sku)
            for sku in PlanningJob.objects.filter(sku_any_query).values_list('sku', flat=True)
            if sku
        }

    seen_skus_in_payload = set()
    repeat_count = 0
    new_count = 0
    for item in items:
        sku_key = _sku_key(item.get('sku'))
        is_new = bool(
            sku_key
            and sku_key not in existing_any_jobs_skus
            and sku_key not in seen_skus_in_payload
        )
        if is_new:
            new_count += 1
        elif sku_key:
            repeat_count += 1
        if sku_key:
            seen_skus_in_payload.add(sku_key)

    return repeat_count, new_count


def _sync_repeat_jobs_from_po(po_doc, actor=None):
    """Create or update draft planning jobs for repeat SKUs from one PO document."""
    payload = po_doc.extracted_payload or {}
    items, _ = _deduplicate_po_items_by_sku(payload.get('items', []))
    po_number = (payload.get('po_number') or '').strip()
    po_date = _parse_iso_date(payload.get('po_date'))
    delivery_location = payload.get('delivery_location', '')
    department = payload.get('department', '')

    if not items:
        return {'created': 0, 'updated': 0, 'locked': 0, 'missing_recipe': 0}

    item_sku_keys = {_sku_key(item.get('sku')) for item in items if item.get('sku')}
    existing_any_jobs_skus = set()
    if item_sku_keys:
        sku_any_query = Q()
        for sku_key in item_sku_keys:
            sku_any_query |= Q(sku__iexact=sku_key)
        existing_any_jobs_skus = {
            _sku_key(sku)
            for sku in PlanningJob.objects.filter(sku_any_query).values_list('sku', flat=True)
            if sku
        }

    recipe_map = _build_recipe_map(items)
    existing_jobs_by_sku = {}
    if po_number and item_sku_keys:
        existing_jobs = PlanningJob.objects.filter(po_number=po_number).order_by('-updated_at', '-id')
        for job in existing_jobs:
            key = _sku_key(job.sku)
            if key in item_sku_keys and key not in existing_jobs_by_sku:
                existing_jobs_by_sku[key] = job

    created_count = 0
    updated_count = 0
    locked_count = 0
    missing_recipe_count = 0

    for item in items:
        sku = (item.get('sku') or '').strip()
        sku_key = _sku_key(sku)
        if not sku_key:
            continue

        # Repeat means this SKU already exists in historical planning jobs.
        if sku_key not in existing_any_jobs_skus:
            continue

        recipe = recipe_map.get(sku_key)
        if not recipe:
            missing_recipe_count += 1
            continue

        existing_job = existing_jobs_by_sku.get(sku_key)
        if existing_job and _normalize_status(existing_job.status) == 'approved':
            locked_count += 1
            continue

        delivery_date = _parse_iso_date(item.get('delivery_date'))
        plan_date = delivery_date or po_date
        qty = item.get('quantity')
        order_qty = int(qty) if qty is not None else None
        unit_cost_val = item.get('unit_cost')
        unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None
        jc_number = existing_job.jc_number if existing_job else allocate_next_jc_number(plan_date)

        defaults = {
            'po_number': po_number,
            'sku': sku,
            'job_name': recipe.job_name or (item.get('job_name') or '').strip() or sku,
            'order_qty': order_qty,
            'department': recipe.department or department,
            'destination': delivery_location,
            'unit_cost': unit_cost_dec if unit_cost_dec is not None else recipe.default_unit_cost,
            'status': 'draft',
            'repeat_flag': 'Repeat',
            'requirement': _append_unique_note_line(
                _append_unique_note_line(
                    _sync_new_sku_requirement(existing_job.requirement if existing_job else '', False),
                    _build_cost_mismatch_note(recipe.default_unit_cost, unit_cost_dec),
                ),
                _build_department_mismatch_note(recipe.department, department),
            ),
            'material': recipe.material,
            'color_spec': recipe.color_spec,
            'application': recipe.application,
            'size_w_mm': recipe.size_w_mm,
            'size_h_mm': recipe.size_h_mm,
            'ups': recipe.ups,
            'print_sheet_size': recipe.print_sheet_size,
            'purchase_sheet_size': recipe.purchase_sheet_size,
            'purchase_sheet_ups': recipe.purchase_sheet_ups,
            'purchase_material': recipe.purchase_material,
            'machine_name': recipe.machine_name,
            'daily_demand': recipe.daily_demand,
            'awc_no': recipe.awc_no,
            'plate_set_no': recipe.plate_set_no,
            'die_cutting': recipe.die_cutting,
        }
        if plan_date:
            defaults['plan_date'] = plan_date
        if actor:
            defaults['created_by'] = actor

        job_obj, created = PlanningJob.objects.update_or_create(
            jc_number=jc_number,
            defaults=defaults,
        )
        if created:
            created_count += 1
        else:
            updated_count += 1
        existing_jobs_by_sku[sku_key] = job_obj

    payload['repeat_jobs_synced'] = True
    payload['repeat_jobs_created_count'] = created_count
    payload['repeat_jobs_updated_count'] = updated_count
    payload['repeat_jobs_locked_count'] = locked_count
    payload['repeat_jobs_missing_recipe_count'] = missing_recipe_count
    po_doc.extracted_payload = payload
    po_doc.save(update_fields=['extracted_payload'])

    return {
        'created': created_count,
        'updated': updated_count,
        'locked': locked_count,
        'missing_recipe': missing_recipe_count,
    }


def _sync_new_jobs_for_approved_sku(sku, actor=None):
    """After SKU master approval, push matching new-job PO lines into Planning Jobs."""
    sku_key = _sku_key(sku)
    if not sku_key:
        return {'created': 0, 'updated': 0, 'locked': 0, 'sent': 0}

    recipe = SkuRecipe.objects.filter(sku__iexact=sku, master_data_status='approved').first()
    if not recipe:
        return {'created': 0, 'updated': 0, 'locked': 0, 'sent': 0}

    existing_any_jobs_skus = {
        _sku_key(value)
        for value in PlanningJob.objects.values_list('sku', flat=True)
        if value
    }

    created_count = 0
    updated_count = 0
    locked_count = 0
    sent_count = 0

    po_docs = PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('created_at', 'id')
    for po_doc in po_docs:
        payload = po_doc.extracted_payload or {}
        items, _ = _deduplicate_po_items_by_sku(payload.get('items', []))
        target_item = None
        for item in items:
            if _sku_key(item.get('sku')) == sku_key:
                target_item = item
                break

        if not target_item:
            continue

        po_number = (payload.get('po_number') or '').strip()
        if not po_number:
            continue

        existing_job = PlanningJob.objects.filter(po_number=po_number, sku__iexact=sku).order_by('-updated_at', '-id').first()
        if existing_job and _normalize_status(existing_job.status) == 'approved':
            locked_count += 1
            continue

        delivery_date = _parse_iso_date(target_item.get('delivery_date'))
        po_date = _parse_iso_date(payload.get('po_date'))
        plan_date = delivery_date or po_date
        qty = target_item.get('quantity')
        order_qty = int(qty) if qty is not None else None
        unit_cost_val = target_item.get('unit_cost')
        unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None
        is_first_production = sku_key not in existing_any_jobs_skus
        jc_number = existing_job.jc_number if existing_job else allocate_next_jc_number(plan_date)
        current_requirement = existing_job.requirement if existing_job else ''

        defaults = {
            'po_number': po_number,
            'sku': sku,
            'job_name': recipe.job_name or (target_item.get('job_name') or '').strip() or sku,
            'order_qty': order_qty,
            'department': recipe.department or (payload.get('department') or ''),
            'destination': payload.get('delivery_location') or '',
            'unit_cost': unit_cost_dec if unit_cost_dec is not None else recipe.default_unit_cost,
            'status': 'draft',
            'repeat_flag': 'New' if is_first_production else 'Repeat',
            'requirement': _sync_new_sku_requirement(current_requirement, is_first_production),
            'material': recipe.material,
            'color_spec': recipe.color_spec,
            'application': recipe.application,
            'size_w_mm': recipe.size_w_mm,
            'size_h_mm': recipe.size_h_mm,
            'ups': recipe.ups,
            'print_sheet_size': recipe.print_sheet_size,
            'purchase_sheet_size': recipe.purchase_sheet_size,
            'purchase_sheet_ups': recipe.purchase_sheet_ups,
            'purchase_material': recipe.purchase_material,
            'machine_name': recipe.machine_name,
            'daily_demand': recipe.daily_demand,
            'awc_no': recipe.awc_no,
            'plate_set_no': recipe.plate_set_no,
            'die_cutting': recipe.die_cutting,
        }

        if not is_first_production:
            defaults['requirement'] = _append_unique_note_line(
                _append_unique_note_line(
                    defaults['requirement'],
                    _build_cost_mismatch_note(recipe.default_unit_cost, unit_cost_dec),
                ),
                _build_department_mismatch_note(recipe.department, payload.get('department') or ''),
            )
        if plan_date:
            defaults['plan_date'] = plan_date
        if actor and not existing_job:
            defaults['created_by'] = actor

        job_obj, created = PlanningJob.objects.update_or_create(
            jc_number=jc_number,
            defaults=defaults,
        )
        if created:
            created_count += 1
        else:
            updated_count += 1

        existing_any_jobs_skus.add(sku_key)
        sent_count += 1

        sent_to_planning = set(payload.get('new_skus_sent_to_planning') or [])
        sent_to_planning.add(sku)
        payload['new_skus_sent_to_planning'] = sorted(sent_to_planning)
        po_doc.extracted_payload = payload
        po_doc.save(update_fields=['extracted_payload'])

    return {
        'created': created_count,
        'updated': updated_count,
        'locked': locked_count,
        'sent': sent_count,
    }


def _merge_po_items_for_existing_po(existing_items, incoming_items):
    """Merge incoming PO lines into existing PO lines without creating duplicates."""
    existing_by_sku = {}
    merged_items = []

    for item in existing_items:
        sku = (item.get('sku') or '').strip()
        sku_key = _sku_key(sku)
        if not sku_key or sku_key in existing_by_sku:
            continue
        item_copy = dict(item)
        existing_by_sku[sku_key] = item_copy
        merged_items.append(item_copy)

    added_skus = []
    updated_skus = []
    ignored_lines = []

    for item in incoming_items:
        sku = (item.get('sku') or '').strip()
        sku_key = _sku_key(sku)
        if not sku_key:
            continue

        incoming_qty = _to_int(item.get('quantity'))
        existing_item = existing_by_sku.get(sku_key)

        if existing_item is None:
            item_copy = dict(item)
            merged_items.append(item_copy)
            existing_by_sku[sku_key] = item_copy
            added_skus.append(sku)
            continue

        existing_qty = _to_int(existing_item.get('quantity'))
        if existing_qty == incoming_qty:
            ignored_lines.append({'sku': sku, 'qty': incoming_qty})
            continue

        # Same SKU but changed qty/fields: treat as correction, not duplicate row.
        for field, value in item.items():
            if value not in (None, ''):
                existing_item[field] = value
        updated_skus.append(sku)

    for idx, item in enumerate(merged_items, start=1):
        item['line_no'] = idx

    return merged_items, sorted(set(added_skus)), sorted(set(updated_skus)), ignored_lines


@login_required
def planning_welcome(request):
    profile = getattr(request.user, 'profile', None)
    user_role = 'unassigned'
    can_edit_jobcard = False
    can_view_reports = False
    can_manage_masters = False

    if profile is not None:
        user_role = (profile.role or 'unassigned').strip().lower()
        can_edit_jobcard = bool(profile.can_edit_jobcard())
        can_view_reports = bool(profile.can_view_reports())
        can_manage_masters = bool(profile.can_manage_masters())

    context = {
        'user_role': user_role,
        'can_edit_jobcard': can_edit_jobcard,
        'can_view_reports': can_view_reports,
        'can_manage_masters': can_manage_masters,
    }
    return render(request, 'planning/planning_welcome.html', context)


@login_required
@permission_required('can_edit_jobcard')
def planning_home(request):
    queryset = PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs').all()

    if request.method == 'POST' and request.POST.get('action') == 'bulk_update_status':
        selected_ids = []
        for raw_id in request.POST.getlist('selected_ids'):
            try:
                selected_ids.append(int(raw_id))
            except (TypeError, ValueError):
                continue

        target_status = _normalize_status(request.POST.get('target_status'), default='')
        if target_status not in PLANNING_STATUS_SET:
            messages.error(request, 'Please select a valid target status for bulk update.')
            return redirect('planning:jobs')

        if not selected_ids:
            messages.error(request, 'Select at least one planning row for bulk update.')
            return redirect('planning:jobs')

        updated = 0
        skipped_locked = 0
        for job in PlanningJob.objects.filter(id__in=selected_ids):
            current_status = _normalize_status(job.status)
            if current_status == 'approved' and target_status not in {'approved', 'reviewed'}:
                skipped_locked += 1
                continue
            if current_status == target_status:
                continue

            job.status = target_status
            if target_status == 'approved':
                job.issued_to_production = True
            elif current_status == 'approved' and target_status == 'reviewed':
                job.issued_to_production = False
            job.save(update_fields=['status', 'issued_to_production', 'updated_at'])
            updated += 1

        messages.success(
            request,
            f'Bulk status update complete. Updated {updated}, locked-skip {skipped_locked}.',
        )
        return redirect('planning:jobs')

    q = (request.GET.get('q') or '').strip()
    status_filter = _normalize_status(request.GET.get('status'), default='')
    department_filter = (request.GET.get('department') or '').strip()
    machine_filter = (request.GET.get('machine') or '').strip()
    from_date = _parse_date_filter(request.GET.get('from_date'))
    to_date = _parse_date_filter(request.GET.get('to_date'))

    if q:
        queryset = queryset.filter(
            Q(jc_number__icontains=q)
            | Q(po_number__icontains=q)
            | Q(sku__icontains=q)
            | Q(job_name__icontains=q)
        )
    if status_filter:
        queryset = queryset.filter(status__iexact=status_filter)
    if department_filter:
        queryset = queryset.filter(department__icontains=department_filter)
    if machine_filter:
        queryset = queryset.filter(machine_name__icontains=machine_filter)
    if from_date:
        queryset = queryset.filter(plan_date__gte=from_date)
    if to_date:
        queryset = queryset.filter(plan_date__lte=to_date)

    status_rows = (
        queryset.values('status')
        .annotate(total=Count('id'))
        .order_by('status')
    )
    status_counts = {
        _normalize_status(row['status']): row['total']
        for row in status_rows
    }

    paginator = Paginator(queryset, 50)
    page_number = request.GET.get('page')
    jobs = paginator.get_page(page_number)
    return render(
        request,
        'planning/planning_home.html',
        {
            'jobs': jobs,
            'status_counts': status_counts,
            'status_choices': PLANNING_STATUSES,
            'filters': {
                'q': q,
                'status': status_filter,
                'department': department_filter,
                'machine': machine_filter,
                'from_date': request.GET.get('from_date', ''),
                'to_date': request.GET.get('to_date', ''),
            },
        },
    )


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def import_planning_sheet(request):
    if request.method == 'POST':
        upload = request.FILES.get('sheet_file')
        if not upload:
            messages.error(request, 'Please choose a CSV file first.')
            return redirect('planning:import_sheet')

        if not upload.name.lower().endswith('.csv'):
            messages.error(request, 'Only CSV file is supported in this first import phase.')
            return redirect('planning:import_sheet')

        decoded = upload.read().decode('utf-8-sig', errors='ignore')
        rows = list(csv.reader(io.StringIO(decoded)))
        header_index = None
        for idx, candidate in enumerate(rows[:8]):
            normalized = {str(col).strip().lower() for col in candidate}
            if 'jc' in normalized and 'job name' in normalized:
                header_index = idx
                break

        if header_index is None:
            messages.error(request, 'Could not detect a valid header row (expected JC and Job Name columns).')
            return redirect('planning:import_sheet')

        header = rows[header_index]
        data_rows = rows[header_index + 1 :]

        imported_count = 0
        updated_count = 0

        for raw_row in data_rows:
            row = {
                header[i]: raw_row[i] if i < len(raw_row) else ''
                for i in range(len(header))
            }
            jc_number = (row.get('JC') or '').strip()
            if not jc_number:
                continue

            defaults = {
                'plan_month': (row.get('Month') or '').strip(),
                'plan_date': _to_date(row.get('Date')),
                'po_number': (row.get('Po') or '').strip(),
                'sku': (row.get('SKU') or '').strip(),
                'job_name': (row.get('Job Name') or '').strip(),
                'repeat_flag': (row.get('Repeat') or '').strip(),
                'material': (row.get('Material') or '').strip(),
                'color_spec': (row.get('Color') or '').strip(),
                'application': (row.get('Application') or '').strip(),
                'size_w_mm': _to_decimal(row.get('Size W mm')),
                'size_h_mm': _to_decimal(row.get('Size H mm')),
                'size_w_inch': _to_decimal(row.get('Size W Inch')),
                'size_h_inch': _to_decimal(row.get('Size H Inch')),
                'order_qty': _to_int(row.get('Order Qty')),
                'print_pcs': _to_int(row.get('Print Pcs')),
                'ups': _to_int(row.get('Ups')),
                'print_sheet_size': (row.get('Print Sheet Size') or '').strip(),
                'print_sheets': _to_int(row.get('Print Sheets')),
                'wastage_sheets': _to_int(row.get('Wastage')),
                'actual_sheet_required': _to_int(row.get('Actual Sheet require')),
                'purchase_sheet_size': (row.get('Purchase Sheet Size') or '').strip(),
                'purchase_sheet_ups': _to_int(row.get('Purchase Sheet ups')),
                'purchase_sheet_required': _to_int(row.get('Purchase Sheet require')),
                'pkt_value': _to_decimal(row.get('PKT')),
                'remarks': (row.get('Remarks  ') or '').strip(),
                'requirement': (row.get('Requirement') or '').strip(),
                'front_colors': _to_int(row.get('No. of Clrs Front')),
                'back_colors': _to_int(row.get('No. Of Clrs Back')),
                'total_colors': _to_int(row.get('Total Crls')),
                'total_mr_time_minutes': _to_int(row.get('Total M/R Time (15m/clr)')),
                'front_pass': _to_int(row.get('Front Pass')),
                'back_pass': _to_int(row.get('Back Pass')),
                'planned_total_impressions': _to_int(row.get('Total Impressions')),
                'mi_quantity': _to_int(row.get('MI Quantity 5')),
                'mi_balance': _to_int(row.get('MI Balance')),
                'remaining_sheet': _to_int(row.get('Remaining sheet')),
                'status': (row.get('status') or '').strip(),
                'pr_reference': (row.get('PR') or '').strip(),
                'rejected_qty': _to_int(row.get('Rejected')),
                'balance_qty': _to_int(row.get('Balance')),
                'destination': (row.get('Destination') or '').strip(),
                'unit_cost': _to_decimal(row.get('Cost')),
                'stock_bag': _to_decimal(row.get('Stock Bag')),
                'machine_name': (row.get('Machine Name') or '').strip(),
                'purchase_material': (row.get('Purchase Material') or '').strip(),
                'stock_qty': _to_decimal(row.get('Stock')),
                'daily_demand': _to_decimal(row.get('Daily Demand')),
                'department': (row.get('Department') or '').strip(),
                'plate_set_no': (row.get('Plate Set No') or '').strip(),
                'awc_no': (row.get('AWC No.') or '').strip(),
                'aging_days': _to_int(row.get('Aging')),
                'die_cutting': (row.get('Die cutting') or '').strip(),
            }

            job, created = PlanningJob.objects.update_or_create(
                jc_number=jc_number,
                defaults=defaults,
            )

            if created:
                imported_count += 1
            else:
                updated_count += 1

            job.print_runs.all().delete()
            print_rows = []
            for i in range(1, 6):
                print_date = _to_date(row.get(f'Print Date {i}'))
                print_qty = _to_int(row.get(f'Print Qty {i}'))
                wastage_qty = _to_int(row.get(f'Wastage {i}'))
                if print_date or print_qty or wastage_qty:
                    print_rows.append(
                        PlanningPrintRun(
                            planning_job=job,
                            run_index=i,
                            print_date=print_date,
                            print_qty=print_qty,
                            wastage_qty=wastage_qty,
                        )
                    )
            if print_rows:
                PlanningPrintRun.objects.bulk_create(print_rows)

            job.dispatch_runs.all().delete()
            dispatch_rows = []
            for i in range(1, 7):
                idx = f'{i:02d}'
                delivery_date = _to_date(row.get(f'Date Delivery {idx}'))
                dc_no = (row.get(f'DC {idx}') or '').strip()
                delivered_qty = _to_int(row.get(f'Delivered Quantity {idx}'))
                if delivery_date or dc_no or delivered_qty:
                    dispatch_rows.append(
                        PlanningDispatchRun(
                            planning_job=job,
                            dispatch_index=i,
                            delivery_date=delivery_date,
                            dc_no=dc_no,
                            delivered_qty=delivered_qty,
                        )
                    )
            if dispatch_rows:
                PlanningDispatchRun.objects.bulk_create(dispatch_rows)

        messages.success(
            request,
            f'Import completed. New jobs: {imported_count}, updated jobs: {updated_count}.',
        )
        return redirect('planning:jobs')

    return render(request, 'planning/planning_import.html')


@login_required
@permission_required('can_edit_jobcard')
def planning_job_detail(request, job_id):
    job = get_object_or_404(
        PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs'),
        id=job_id,
    )
    is_repeat_with_changes = (
        (job.repeat_flag or '').lower() == 'repeat'
        and job.has_edits_since_creation
        and job.edited_fields_list
    )
    return render(
        request,
        'planning/planning_job_detail.html',
        {
            'job': job,
            'status_now': _normalize_status(job.status),
            'is_repeat_with_changes': is_repeat_with_changes,
            'changed_fields': job.edited_fields_list or [],
            'last_edited_by': job.last_edited_by,
            'last_edited_at': job.last_edited_at,
        },
    )


@login_required
@permission_required('can_edit_jobcard')
def planning_job_edit(request, job_id):
    job = get_object_or_404(PlanningJob, id=job_id)
    current_status = _normalize_status(job.status)

    if current_status == 'approved':
        messages.error(request, 'Approved records are locked. Unlock to Reviewed before editing.')
        return redirect('planning:job_detail', job_id=job.id)

    if request.method == 'POST':
        form = PlanningJobEditForm(request.POST, instance=job)
        if form.is_valid():
            edited = form.save(commit=False)
            edited.status = _normalize_status(edited.status)
            edited.job_card_version = (job.job_card_version or 1) + 1
            
            # Detect changes for repeat jobs
            if (job.repeat_flag or '').lower() == 'repeat':
                changed_fields = []
                edit_fields = ['plan_date', 'po_number', 'sku', 'job_name', 'material', 'color_spec', 'application',
                               'order_qty', 'print_sheets', 'machine_name', 'department', 'destination', 'unit_cost',
                               'daily_demand', 'remarks', 'requirement', 'status', 'print_sheet_size',
                               'purchase_sheet_size', 'ups']
                for field in edit_fields:
                    old_val = getattr(job, field, None)
                    new_val = getattr(edited, field, None)
                    if str(old_val) != str(new_val):
                        changed_fields.append(field)
                
                if changed_fields:
                    edited.has_edits_since_creation = True
                    edited.edited_fields_list = changed_fields
                    edited.last_edited_by = request.user
                    edited.last_edited_at = timezone.now()
            
            edited.save()
            messages.success(request, f'Planning job {edited.jc_number} updated.')
            if edited.has_edits_since_creation and (edited.repeat_flag or '').lower() == 'repeat':
                messages.info(request, f'Changes detected and flagged for production team: {', '.join(edited.edited_fields_list)}')
            return redirect('planning:job_detail', job_id=edited.id)
    else:
        form = PlanningJobEditForm(instance=job)

    return render(request, 'planning/planning_job_edit.html', {'job': job, 'form': form})


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def planning_job_status_update(request, job_id):
    if request.method != 'POST':
        return redirect('planning:job_detail', job_id=job_id)

    job = get_object_or_404(PlanningJob, id=job_id)
    current_status = _normalize_status(job.status)
    transition = (request.POST.get('transition') or '').strip()

    transitions = {
        'submit_review': ('draft', 'reviewed'),
        'approve': ('reviewed', 'approved'),
        'unlock': ('approved', 'reviewed'),
        'mark_closed': (None, 'closed'),
        'reopen': ('closed', 'draft'),
    }
    if transition not in transitions:
        messages.error(request, 'Unknown status transition request.')
        return redirect('planning:job_detail', job_id=job.id)

    required_from, target_status = transitions[transition]
    if required_from and current_status != required_from:
        messages.error(request, f'Transition not allowed from {current_status} to {target_status}.')
        return redirect('planning:job_detail', job_id=job.id)

    if current_status == target_status:
        messages.info(request, f'Job already in {target_status} status.')
        return redirect('planning:job_detail', job_id=job.id)

    job.status = target_status
    if target_status == 'approved':
        job.issued_to_production = True
    if transition == 'unlock':
        job.issued_to_production = False
    job.save(update_fields=['status', 'issued_to_production', 'updated_at'])

    messages.success(request, f'Job status updated: {current_status} -> {target_status}.')
    return redirect('planning:job_detail', job_id=job.id)


@login_required
@permission_required('can_edit_jobcard')
def planning_job_card_print(request, job_id):
    job = get_object_or_404(
        PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs'),
        id=job_id,
    )
    scan_url = request.build_absolute_uri(reverse('planning:scan_open', args=[job.jc_number]))
    is_repeat_with_changes = (
        (job.repeat_flag or '').lower() == 'repeat'
        and job.has_edits_since_creation
        and job.edited_fields_list
    )
    context = {
        'job': job,
        'now_ts': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'status_now': _normalize_status(job.status),
        'scan_url': scan_url,
        'qr_code_b64': _build_qr_image_base64(scan_url),
        'is_repeat_with_changes': is_repeat_with_changes,
        'changed_fields': job.edited_fields_list or [],
        'last_edited_by': job.last_edited_by,
        'last_edited_at': job.last_edited_at,
    }
    return render(request, 'planning/planning_job_card_print.html', context)


@login_required
@permission_required('can_edit_jobcard')
def planning_report(request):
    queryset = PlanningJob.objects.all()

    from_date = _parse_date_filter(request.GET.get('from_date'))
    to_date = _parse_date_filter(request.GET.get('to_date'))
    if from_date:
        queryset = queryset.filter(plan_date__gte=from_date)
    if to_date:
        queryset = queryset.filter(plan_date__lte=to_date)

    totals = queryset.aggregate(
        total_jobs=Count('id'),
        total_order_qty=Sum('order_qty'),
        approved_jobs=Count('id', filter=Q(status__iexact='approved')),
        closed_jobs=Count('id', filter=Q(status__iexact='closed')),
    )

    by_status = (
        queryset.values('status')
        .annotate(total=Count('id'), order_qty=Sum('order_qty'))
        .order_by('status')
    )
    by_department = (
        queryset.values('department')
        .annotate(total=Count('id'), order_qty=Sum('order_qty'))
        .order_by('-total', 'department')[:20]
    )
    by_machine = (
        queryset.values('machine_name')
        .annotate(total=Count('id'), order_qty=Sum('order_qty'))
        .order_by('-total', 'machine_name')[:20]
    )

    context = {
        'totals': totals,
        'by_status': by_status,
        'by_department': by_department,
        'by_machine': by_machine,
        'filters': {
            'from_date': request.GET.get('from_date', ''),
            'to_date': request.GET.get('to_date', ''),
        },
    }
    return render(request, 'planning/planning_report.html', context)


@login_required
@permission_required('can_edit_jobcard')
def planning_scan(request):
    if request.method == 'POST':
        raw_code = (request.POST.get('scan_code') or '').strip()
        if not raw_code:
            messages.error(request, 'Scan code cannot be empty.')
            return redirect('planning:scan')

        # QR may contain full URL, plain JC number, or prefixed JC field.
        parsed = raw_code
        if '/scan/open/' in parsed:
            parsed = parsed.rsplit('/scan/open/', 1)[-1]
        if '?' in parsed:
            parsed = parsed.split('?', 1)[0]
        parsed = parsed.replace('JC:', '').strip().strip('/')

        job = PlanningJob.objects.filter(jc_number__iexact=parsed).order_by('-id').first()
        if not job:
            messages.error(request, f'No planning job found for code: {parsed}')
            return redirect('planning:scan')

        return redirect('planning:job_detail', job_id=job.id)

    return render(request, 'planning/planning_scan.html')


@login_required
@permission_required('can_edit_jobcard')
def planning_scan_open(request, jc_number):
    job = PlanningJob.objects.filter(jc_number__iexact=(jc_number or '').strip()).order_by('-id').first()
    if not job:
        messages.error(request, f'No planning job found for code: {jc_number}')
        return redirect('planning:scan')
    return redirect('planning:job_detail', job_id=job.id)


@login_required
def po_debug_extract(request):
    """Debug view: upload PDF and see raw text + table rows + per-strategy parse results."""
    import json as _json
    context = {}
    if request.method == 'POST':
        pdf_file = request.FILES.get('po_pdf')
        if pdf_file:
            try:
                import pdfplumber
                full_text = ''
                table_blobs = []
                table_rows = []
                pdf_file.seek(0)
                with pdfplumber.open(pdf_file) as pdf:
                    for page in pdf.pages:
                        page_text = page.extract_text(x_tolerance=3, y_tolerance=3) or ''
                        full_text += page_text + '\n'
                        for table in (page.extract_tables() or []):
                            for row in table or []:
                                parts = [str(col).strip() for col in (row or []) if str(col).strip()]
                                if parts:
                                    table_blobs.append(' '.join(parts))
                                    table_rows.append(parts)

                from .po_extractor import (
                    _build_sku_jobname_map,
                    _detect_expected_line_count,
                    _extract_items_flexible,
                    _extract_items_from_table_blobs,
                    _extract_items_from_table_rows,
                    _extract_items_from_text_windows,
                    _extract_items_strict,
                )
                sku_map = _build_sku_jobname_map(full_text, table_blobs)
                expected = _detect_expected_line_count(full_text, table_rows)
                strict = _extract_items_strict(full_text, sku_map)
                flexible = _extract_items_flexible(full_text, sku_map)
                from_rows = _extract_items_from_table_rows(table_rows, sku_map)
                from_blobs = _extract_items_from_table_blobs(table_blobs, sku_map)
                from_windows = _extract_items_from_text_windows(full_text, sku_map)

                context = {
                    'full_text': full_text,
                    'table_rows': _json.dumps(table_rows, indent=2),
                    'table_blobs': _json.dumps(table_blobs, indent=2),
                    'expected': expected,
                    'strict': _json.dumps(strict, indent=2),
                    'flexible': _json.dumps(flexible, indent=2),
                    'from_rows': _json.dumps(from_rows, indent=2),
                    'from_blobs': _json.dumps(from_blobs, indent=2),
                    'from_windows': _json.dumps(from_windows, indent=2),
                    'strict_count': len(strict),
                    'flexible_count': len(flexible),
                    'from_rows_count': len(from_rows),
                    'from_blobs_count': len(from_blobs),
                    'from_windows_count': len(from_windows),
                }
            except Exception as exc:
                context = {'error': str(exc)}
    return render(request, 'planning/po_debug.html', context)


@login_required
@permission_required('can_edit_jobcard')
def sku_recipes_list(request):
    """List all SKU recipes with search; handles delete via POST."""
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        recipe_id = request.POST.get('recipe_id')

        if action == 'delete':
            try:
                SkuRecipe.objects.filter(id=int(recipe_id)).delete()
                messages.success(request, 'SKU Recipe deleted.')
            except (TypeError, ValueError):
                messages.error(request, 'Invalid recipe ID.')
            return redirect('planning:sku_recipes')

        if action in {'submit_review', 'approve', 'back_to_draft'}:
            try:
                recipe = SkuRecipe.objects.get(id=int(recipe_id))
            except (TypeError, ValueError, SkuRecipe.DoesNotExist):
                messages.error(request, 'Invalid recipe ID.')
                return redirect('planning:sku_recipes')

            current_status = (recipe.master_data_status or 'draft').lower()

            if action == 'submit_review':
                if current_status != 'draft':
                    messages.error(request, f'SKU {recipe.sku} can only move to Reviewed from Draft.')
                    return redirect('planning:sku_recipes')
                recipe.master_data_status = 'reviewed'
                recipe.reviewed_by = request.user
                recipe.reviewed_at = timezone.now()
                recipe.approved_by = None
                recipe.approved_at = None
                recipe.save(update_fields=['master_data_status', 'reviewed_by', 'reviewed_at', 'approved_by', 'approved_at', 'updated_at'])
                messages.success(request, f'SKU {recipe.sku} moved to Reviewed.')
                return redirect('planning:sku_recipes')

            if action == 'approve':
                if current_status != 'reviewed':
                    messages.error(request, f'SKU {recipe.sku} can only be Approved from Reviewed status.')
                    return redirect('planning:sku_recipes')
                missing_required = _missing_required_master_fields(recipe, recipe.job_name)
                if missing_required:
                    messages.error(
                        request,
                        f'SKU {recipe.sku} cannot be approved. Missing required master data: {", ".join(missing_required)}.',
                    )
                    return redirect('planning:sku_recipes')
                recipe.master_data_status = 'approved'
                recipe.approved_by = request.user
                recipe.approved_at = timezone.now()
                recipe.save(update_fields=['master_data_status', 'approved_by', 'approved_at', 'updated_at'])
                messages.success(request, f'SKU {recipe.sku} approved.')
                return redirect('planning:sku_recipes')

            if action == 'back_to_draft':
                if current_status == 'draft':
                    messages.info(request, f'SKU {recipe.sku} is already in Draft.')
                    return redirect('planning:sku_recipes')
                recipe.master_data_status = 'draft'
                recipe.reviewed_by = None
                recipe.reviewed_at = None
                recipe.approved_by = None
                recipe.approved_at = None
                recipe.save(update_fields=['master_data_status', 'reviewed_by', 'reviewed_at', 'approved_by', 'approved_at', 'updated_at'])
                messages.success(request, f'SKU {recipe.sku} moved back to Draft.')
                return redirect('planning:sku_recipes')

    q = (request.GET.get('q') or '').strip()
    qs = SkuRecipe.objects.all()
    if q:
        qs = qs.filter(
            Q(sku__icontains=q)
            | Q(job_name__icontains=q)
            | Q(material__icontains=q)
            | Q(machine_name__icontains=q)
            | Q(department__icontains=q)
        )
    paginator = Paginator(qs, 50)
    recipes = paginator.get_page(request.GET.get('page'))
    return render(request, 'planning/sku_recipes.html', {'recipes': recipes, 'q': q})


@login_required
@permission_required('can_edit_jobcard')
def sku_recipe_edit(request, recipe_id=None):
    """Create or edit a single SKU recipe."""
    if recipe_id:
        recipe = get_object_or_404(SkuRecipe, id=recipe_id)
        page_title = f'Edit SKU Recipe — {recipe.sku}'
    else:
        recipe = None
        page_title = 'Add New SKU Recipe'

    if request.method == 'POST':
        form = SkuRecipeForm(request.POST, instance=recipe)
        if form.is_valid():
            obj = form.save(commit=False)
            if not recipe_id:
                obj.created_by = request.user
            # Stage-1 workflow: every new or changed recipe starts in Draft and requires approval.
            obj.master_data_status = 'draft'
            obj.reviewed_by = None
            obj.reviewed_at = None
            obj.approved_by = None
            obj.approved_at = None
            obj.save()
            messages.success(request, f'SKU Recipe "{obj.sku}" saved as Draft. Submit for approval from SKU Recipe Master.')
            return redirect('planning:sku_recipes')
    else:
        form = SkuRecipeForm(instance=recipe)

    return render(request, 'planning/sku_recipe_edit.html', {'form': form, 'recipe': recipe, 'page_title': page_title})


@login_required
@permission_required('can_edit_jobcard')
def sku_recipe_bulk_upload(request):
    """Bulk upload SKU recipes from CSV/XLSX into Draft status for approval workflow."""
    if request.method == 'POST':
        upload_file = request.FILES.get('upload_file')
        if not upload_file:
            messages.error(request, 'Please choose a CSV or XLSX file to upload.')
            return redirect('planning:sku_recipe_bulk_upload')

        name = (upload_file.name or '').lower()
        rows = []

        try:
            if name.endswith('.csv'):
                decoded = upload_file.read().decode('utf-8-sig')
                reader = csv.DictReader(io.StringIO(decoded))
                rows = list(reader)
            elif name.endswith('.xlsx'):
                try:
                    import openpyxl
                except ImportError:
                    messages.error(request, 'openpyxl is required for XLSX upload.')
                    return redirect('planning:sku_recipe_bulk_upload')
                wb = openpyxl.load_workbook(upload_file, data_only=True)
                ws = wb.active
                header = [str(c).strip() if c is not None else '' for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])]
                for values in ws.iter_rows(min_row=2, values_only=True):
                    row = {}
                    for idx, key in enumerate(header):
                        if key:
                            row[key] = values[idx] if idx < len(values) else None
                    rows.append(row)
            else:
                messages.error(request, 'Unsupported file type. Please upload CSV or XLSX.')
                return redirect('planning:sku_recipe_bulk_upload')
        except Exception as exc:
            messages.error(request, f'Could not read upload file: {exc}')
            return redirect('planning:sku_recipe_bulk_upload')

        if not rows:
            messages.error(request, 'No rows found in upload file.')
            return redirect('planning:sku_recipe_bulk_upload')

        field_map = {
            'sku': ['sku', 'SKU'],
            'job_name': ['job_name', 'Job Name'],
            'material': ['material', 'Material'],
            'color_spec': ['color_spec', 'Color', 'Colour'],
            'application': ['application', 'Application'],
            'size_w_mm': ['size_w_mm', 'Size W (mm)'],
            'size_h_mm': ['size_h_mm', 'Size H (mm)'],
            'ups': ['ups', 'UPS'],
            'print_sheet_size': ['print_sheet_size', 'Print Sheet'],
            'purchase_sheet_size': ['purchase_sheet_size', 'Purchase Sheet'],
            'purchase_sheet_ups': ['purchase_sheet_ups', 'Purchase Sheet Ups'],
            'purchase_material': ['purchase_material', 'Purchase Material Origin'],
            'machine_name': ['machine_name', 'Machine'],
            'department': ['department', 'Department'],
            'default_unit_cost': ['default_unit_cost', 'Unit Cost'],
            'daily_demand': ['daily_demand', 'Daily Demand'],
            'awc_no': ['awc_no', 'AWC No'],
            'plate_set_no': ['plate_set_no', 'Plate Set No'],
            'die_cutting': ['die_cutting', 'Die'],
            'notes': ['notes', 'Notes'],
        }

        def pick_value(source, aliases):
            for key in aliases:
                if key in source and source.get(key) not in (None, ''):
                    return source.get(key)
            return ''

        created = 0
        updated = 0
        failed = 0
        sample_errors = []

        for idx, source in enumerate(rows, start=2):
            payload = {}
            for field_name, aliases in field_map.items():
                value = pick_value(source, aliases)
                payload[field_name] = '' if value is None else str(value).strip()

            if not payload['sku']:
                continue

            existing = SkuRecipe.objects.filter(sku__iexact=payload['sku']).first()
            form = SkuRecipeForm(payload, instance=existing)
            if not form.is_valid():
                failed += 1
                if len(sample_errors) < 8:
                    error_text = '; '.join(
                        f"{name}: {', '.join([str(msg) for msg in msgs])}"
                        for name, msgs in form.errors.items()
                    )
                    sample_errors.append(f'Row {idx} ({payload["sku"]}): {error_text}')
                continue

            obj = form.save(commit=False)
            if not existing:
                obj.created_by = request.user
            obj.master_data_status = 'draft'
            obj.reviewed_by = None
            obj.reviewed_at = None
            obj.approved_by = None
            obj.approved_at = None
            obj.save()

            if existing:
                updated += 1
            else:
                created += 1

        if created or updated:
            messages.success(
                request,
                f'Bulk upload complete. Draft recipes created {created}, updated {updated}, failed {failed}.',
            )
        if failed and sample_errors:
            messages.error(request, 'Sample row errors: ' + ' | '.join(sample_errors))

        return redirect('planning:sku_recipes')

    return render(request, 'planning/sku_recipe_bulk_upload.html')


def _collect_pending_sku_rows(po_docs):
    """Build pending SKU rows from PO documents where SKU recipe is missing."""
    rows = []
    for po_doc in po_docs:
        payload = po_doc.extracted_payload or {}
        items = _sanitize_po_payload_items(payload)
        if not items:
            continue

        recipe_map = _build_recipe_map(items)
        _, _, _, missing_skus = _annotate_items_with_recipe(items, recipe_map)
        if not missing_skus:
            continue

        item_map = {}
        for item in items:
            key = _sku_key(item.get('sku'))
            if key and key not in item_map:
                item_map[key] = item

        po_number = payload.get('po_number') or '-'
        for sku in missing_skus:
            item = item_map.get(_sku_key(sku), {})
            rows.append(
                {
                    'po_doc_id': po_doc.id,
                    'po_number': po_number,
                    'sku': sku,
                    'job_name': (item.get('job_name') or '').strip() or sku,
                    'qty': _format_display_qty(item.get('quantity')),
                    'delivery_date': item.get('delivery_date') or '-',
                    'uploaded_at': po_doc.created_at,
                }
            )

    return rows


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def pending_skus(request):
    """Central queue of SKUs that are still missing in SKU Recipe master data."""
    if request.method == 'POST':
        action = (request.POST.get('action') or 'save').strip()
        sku = (request.POST.get('sku') or '').strip()
        po_doc_id = request.POST.get('po_doc_id')
        return_po = (request.POST.get('return_po') or '').strip()
        return_q = (request.POST.get('return_q') or '').strip()

        def _redirect_pending():
            params = {}
            if return_po:
                params['po'] = return_po
            if return_q:
                params['q'] = return_q
            url = reverse('planning:pending_skus')
            return redirect(f'{url}?{urlencode(params)}' if params else url)

        if not sku:
            messages.error(request, 'SKU is required.')
            return _redirect_pending()

        if action in {'submit_review', 'approve', 'back_to_draft'}:
            recipe = SkuRecipe.objects.filter(sku__iexact=sku).first()
            if not recipe:
                messages.error(request, f'No SKU recipe found for {sku}. Save recipe data first.')
                return _redirect_pending()

            current_status = (recipe.master_data_status or 'draft').lower()

            if action == 'submit_review':
                if current_status != 'draft':
                    messages.error(request, f'SKU {sku} can only move to Reviewed from Draft.')
                    return _redirect_pending()
                recipe.master_data_status = 'reviewed'
                recipe.reviewed_by = request.user
                recipe.reviewed_at = timezone.now()
                recipe.approved_by = None
                recipe.approved_at = None
                recipe.save(update_fields=['master_data_status', 'reviewed_by', 'reviewed_at', 'approved_by', 'approved_at', 'updated_at'])
                messages.success(request, f'SKU {sku} moved to Reviewed.')
                return _redirect_pending()

            if action == 'approve':
                if current_status != 'reviewed':
                    messages.error(request, f'SKU {sku} can only be Approved from Reviewed status.')
                    return _redirect_pending()
                missing_required = _missing_required_master_fields(recipe)
                if missing_required:
                    messages.error(
                        request,
                        f'SKU {sku} cannot be approved. Missing required master data: {", ".join(missing_required)}.',
                    )
                    return _redirect_pending()
                recipe.master_data_status = 'approved'
                recipe.approved_by = request.user
                recipe.approved_at = timezone.now()
                recipe.save(update_fields=['master_data_status', 'approved_by', 'approved_at', 'updated_at'])
                sync_result = _sync_new_jobs_for_approved_sku(sku, actor=request.user)
                messages.success(
                    request,
                    f'SKU {sku} approved for master data usage. Sent to Planning: {sync_result["sent"]} PO line(s), created {sync_result["created"]}, updated {sync_result["updated"]}, locked {sync_result["locked"]}.',
                )
                return _redirect_pending()

            if action == 'back_to_draft':
                if current_status == 'draft':
                    messages.info(request, f'SKU {sku} is already in Draft.')
                    return _redirect_pending()
                recipe.master_data_status = 'draft'
                recipe.reviewed_by = None
                recipe.reviewed_at = None
                recipe.approved_by = None
                recipe.approved_at = None
                recipe.save(update_fields=['master_data_status', 'reviewed_by', 'reviewed_at', 'approved_by', 'approved_at', 'updated_at'])
                messages.success(request, f'SKU {sku} moved back to Draft.')
                return _redirect_pending()

        job_name = (request.POST.get('job_name') or '').strip()
        material = (request.POST.get('material') or '').strip()
        color_spec = (request.POST.get('color_spec') or '').strip()
        application = (request.POST.get('application') or '').strip()
        machine_name = (request.POST.get('machine_name') or '').strip()
        department = (request.POST.get('department') or '').strip()
        print_sheet_size = (request.POST.get('print_sheet_size') or '').strip()
        purchase_sheet_size = (request.POST.get('purchase_sheet_size') or '').strip()
        purchase_sheet_ups = _to_optional_positive_int(request.POST.get('purchase_sheet_ups'))
        ups = _to_optional_positive_int(request.POST.get('ups'))
        purchase_material = (request.POST.get('purchase_material') or '').strip()
        daily_demand = _to_optional_decimal(request.POST.get('daily_demand'))
        awc_no = (request.POST.get('awc_no') or '').strip()
        plate_set_no = (request.POST.get('plate_set_no') or '').strip()
        die_cutting = (request.POST.get('die_cutting') or '').strip()

        unit_cost_raw = (request.POST.get('default_unit_cost') or '').strip()
        unit_cost = None
        if unit_cost_raw:
            try:
                unit_cost = Decimal(unit_cost_raw)
            except InvalidOperation:
                unit_cost = None

        if not job_name and not material and not machine_name:
            messages.error(request, 'Please enter at least Job Name, Material, or Machine before saving.')
            return _redirect_pending()

        SkuRecipe.objects.update_or_create(
            sku=sku,
            defaults={
                'job_name': job_name,
                'material': material,
                'color_spec': color_spec,
                'application': application,
                'machine_name': machine_name,
                'department': department,
                'print_sheet_size': print_sheet_size,
                'purchase_sheet_size': purchase_sheet_size,
                'purchase_sheet_ups': purchase_sheet_ups,
                'ups': ups,
                'purchase_material': purchase_material,
                'default_unit_cost': unit_cost,
                'daily_demand': daily_demand,
                'awc_no': awc_no,
                'plate_set_no': plate_set_no,
                'die_cutting': die_cutting,
                'created_by': request.user,
                'master_data_status': 'draft',
                'reviewed_by': None,
                'reviewed_at': None,
                'approved_by': None,
                'approved_at': None,
            },
        )

        if po_doc_id:
            try:
                po_doc = PoDocument.objects.filter(id=int(po_doc_id)).first()
            except (TypeError, ValueError):
                po_doc = None
            if po_doc:
                payload = po_doc.extracted_payload or {}
                configured = set(payload.get('new_skus_configured') or [])
                configured.add(sku)
                payload['new_skus_configured'] = sorted(configured)
                po_doc.extracted_payload = payload
                po_doc.save(update_fields=['extracted_payload'])

        messages.success(request, f'SKU recipe saved for {sku}.')
        return _redirect_pending()

    po_filter = (request.GET.get('po') or '').strip()
    q = (request.GET.get('q') or '').strip()

    po_docs = PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('-created_at')[:400]
    deduped_docs = []
    seen_po_numbers = set()
    for doc in po_docs:
        payload = doc.extracted_payload or {}
        po_number = (payload.get('po_number') or '').strip().upper()
        if po_number:
            if po_number in seen_po_numbers:
                continue
            seen_po_numbers.add(po_number)
        deduped_docs.append(doc)
    po_docs = deduped_docs[:200]
    all_pending_rows = _collect_pending_sku_rows(po_docs)

    po_summary_map = {}
    for row in all_pending_rows:
        po_key = row.get('po_number') or '-'
        current = po_summary_map.get(po_key)
        if not current:
            po_summary_map[po_key] = {
                'po_number': po_key,
                'count': 1,
                'po_doc_id': row.get('po_doc_id'),
            }
        else:
            current['count'] += 1

    pending_rows = all_pending_rows
    if po_filter:
        pending_rows = [row for row in pending_rows if (row.get('po_number') or '') == po_filter]
    if q:
        q_upper = q.upper()
        pending_rows = [
            row
            for row in pending_rows
            if q_upper in (row.get('sku') or '').upper()
            or q_upper in (row.get('po_number') or '').upper()
            or q_upper in (row.get('job_name') or '').upper()
        ]

    sku_values = sorted({row['sku'] for row in pending_rows if row.get('sku')})
    recipes_by_sku = {}
    if sku_values:
        recipe_query = Q()
        for sku in sku_values:
            recipe_query |= Q(sku__iexact=sku)
        recipes = SkuRecipe.objects.filter(recipe_query)
        recipes_by_sku = {recipe.sku.upper(): recipe for recipe in recipes}

    for row in pending_rows:
        recipe = recipes_by_sku.get(_sku_key(row.get('sku')))
        row['recipe'] = recipe
        row['recipe_status'] = recipe.master_data_status if recipe else 'missing'
        row['missing_required_fields'] = _missing_required_master_fields(recipe, row.get('job_name') or '')

    pending_rows.sort(key=lambda row: (row['po_number'], row['sku']))

    context = {
        'pending_rows': pending_rows,
        'pending_count': len(pending_rows),
        'po_summary': sorted(po_summary_map.values(), key=lambda x: x['po_number']),
        'po_filter': po_filter,
        'q': q,
    }
    return render(request, 'planning/pending_skus.html', context)


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def pending_sku_master_entry(request):
    """Open a focused form for one pending SKU and send it through master-data approval flow."""
    sku = (request.GET.get('sku') or request.POST.get('sku') or '').strip()
    po_doc_id_raw = request.GET.get('po_doc_id') or request.POST.get('po_doc_id')
    return_po = (request.GET.get('return_po') or request.POST.get('return_po') or '').strip()
    return_q = (request.GET.get('return_q') or request.POST.get('return_q') or '').strip()

    def _redirect_pending():
        params = {}
        if return_po:
            params['po'] = return_po
        if return_q:
            params['q'] = return_q
        url = reverse('planning:pending_skus')
        return redirect(f'{url}?{urlencode(params)}' if params else url)

    try:
        po_doc_id = int(po_doc_id_raw)
    except (TypeError, ValueError):
        po_doc_id = None

    if not sku or not po_doc_id:
        messages.error(request, 'Missing SKU or PO reference for master-data entry.')
        return _redirect_pending()

    po_doc = PoDocument.objects.filter(id=po_doc_id).first()
    if not po_doc:
        messages.error(request, 'PO document was not found.')
        return _redirect_pending()

    payload = po_doc.extracted_payload or {}
    po_number = payload.get('po_number') or '-'
    items = _sanitize_po_payload_items(payload)

    suggested_item = None
    sku_key = _sku_key(sku)
    for item in items:
        if _sku_key(item.get('sku')) == sku_key:
            suggested_item = item
            break
    suggested_item = suggested_item or {}
    po_job_name = (suggested_item.get('job_name') or '').strip() or sku
    po_department = (payload.get('department') or '').strip()
    po_unit_cost = _to_decimal(suggested_item.get('unit_cost'))
    po_color_spec = _normalize_color_spec_input(
        suggested_item.get('color_spec') or suggested_item.get('colour') or suggested_item.get('color') or ''
    )
    po_application = _normalize_application_input(suggested_item.get('application') or payload.get('application') or '')

    recipe = SkuRecipe.objects.filter(sku__iexact=sku).first()

    if request.method == 'POST':
        posted = request.POST.copy()
        # Job name is sourced from PO parsing; keep it authoritative and non-editable.
        posted['job_name'] = po_job_name
        posted['sku'] = sku
        if not (posted.get('department') or '').strip() and po_department:
            posted['department'] = po_department
        if not (posted.get('default_unit_cost') or '').strip() and po_unit_cost is not None:
            posted['default_unit_cost'] = str(po_unit_cost)
        if not (posted.get('color_spec') or '').strip() and po_color_spec:
            posted['color_spec'] = po_color_spec
        if not (posted.get('application') or '').strip() and po_application:
            posted['application'] = po_application
        form = SkuRecipeForm(posted, instance=recipe)
        if form.is_valid():
            action = (request.POST.get('action') or 'save_draft').strip()
            obj = form.save(commit=False)
            obj.sku = sku
            obj.job_name = po_job_name
            if not recipe:
                obj.created_by = request.user

            if action == 'submit_review':
                missing_required = _missing_required_master_fields(obj)
                if missing_required:
                    messages.error(
                        request,
                        f'SKU {sku} cannot be sent for approval. Missing required data: {", ".join(missing_required)}.',
                    )
                else:
                    obj.master_data_status = 'reviewed'
                    obj.reviewed_by = request.user
                    obj.reviewed_at = timezone.now()
                    obj.approved_by = None
                    obj.approved_at = None
                    obj.save()

                    configured = set(payload.get('new_skus_configured') or [])
                    configured.add(sku)
                    payload['new_skus_configured'] = sorted(configured)
                    po_doc.extracted_payload = payload
                    po_doc.save(update_fields=['extracted_payload'])

                    messages.success(request, f'SKU {sku} submitted for approval review.')
                    return _redirect_pending()
            else:
                obj.master_data_status = 'draft'
                obj.reviewed_by = None
                obj.reviewed_at = None
                obj.approved_by = None
                obj.approved_at = None
                obj.save()

                configured = set(payload.get('new_skus_configured') or [])
                configured.add(sku)
                payload['new_skus_configured'] = sorted(configured)
                po_doc.extracted_payload = payload
                po_doc.save(update_fields=['extracted_payload'])

                messages.success(request, f'SKU {sku} saved as Draft.')
                return _redirect_pending()
    else:
        initial = {
            'sku': sku,
            'job_name': po_job_name,
            'department': (recipe.department if recipe else '') or po_department,
            'default_unit_cost': (recipe.default_unit_cost if recipe else None) or po_unit_cost,
            'color_spec': (recipe.color_spec if recipe else '') or po_color_spec,
            'application': (recipe.application if recipe else '') or po_application,
        }
        form = SkuRecipeForm(instance=recipe, initial=initial)

    form.fields['sku'].widget.attrs['readonly'] = True

    current_recipe = recipe
    if request.method == 'POST' and form.is_valid() and 'obj' in locals():
        current_recipe = obj

    mismatch_alerts = []
    if current_recipe:
        cost_alert = _build_cost_mismatch_note(current_recipe.default_unit_cost, po_unit_cost)
        dept_alert = _build_department_mismatch_note(current_recipe.department, po_department)
        if cost_alert:
            mismatch_alerts.append(cost_alert)
        if dept_alert:
            mismatch_alerts.append(dept_alert)

    context = {
        'form': form,
        'sku': sku,
        'po_doc_id': po_doc_id,
        'po_number': po_number,
        'return_po': return_po,
        'return_q': return_q,
        'suggested_job_name': po_job_name,
        'suggested_qty': _format_display_qty(suggested_item.get('quantity')),
        'suggested_delivery_date': suggested_item.get('delivery_date') or '-',
        'recipe_status': (current_recipe.master_data_status if current_recipe else 'missing'),
        'missing_required_fields': _missing_required_master_fields(current_recipe, po_job_name),
        'mismatch_alerts': mismatch_alerts,
    }
    return render(request, 'planning/pending_sku_master_entry.html', context)


@login_required
@permission_required('can_edit_jobcard')
def po_inbox(request):
    """PO intake queue after upload: split-ready documents with repeat/new counts."""
    docs = PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('-created_at')[:400]
    deduped_docs = []
    seen_po_numbers = set()
    for doc in docs:
        payload = doc.extracted_payload or {}
        po_number = (payload.get('po_number') or '').strip().upper()
        if po_number:
            if po_number in seen_po_numbers:
                continue
            seen_po_numbers.add(po_number)
        deduped_docs.append(doc)

    docs = deduped_docs[:200]
    rows = []
    for doc in docs:
        payload = doc.extracted_payload or {}
        items, _ = _deduplicate_po_items_by_sku(payload.get('items', []))
        recipe_map = _build_recipe_map(items)
        _, _, _, missing_skus = _annotate_items_with_recipe(items, recipe_map)
        repeat_count, new_count = _history_repeat_new_counts(items)
        rows.append(
            {
                'doc': doc,
                'po_number': payload.get('po_number') or '-',
                'supplier': payload.get('supplier_name') or '-',
                'item_count': len(items),
                'repeat_count': repeat_count,
                'new_count': new_count,
                'missing_count': len(missing_skus),
                'repeat_jobs_created_count': payload.get('repeat_jobs_created_count') or 0,
                'repeat_jobs_updated_count': payload.get('repeat_jobs_updated_count') or 0,
                'repeat_jobs_locked_count': payload.get('repeat_jobs_locked_count') or 0,
                'repeat_jobs_missing_recipe_count': payload.get('repeat_jobs_missing_recipe_count') or 0,
                'new_skus_sent_to_planning_count': len(payload.get('new_skus_sent_to_planning') or []),
                'merged_duplicates': bool(payload.get('merged_duplicate_skus')),
                'merged_duplicate_skus': payload.get('merged_duplicate_skus') or [],
                'source_item_count': payload.get('source_item_count') or len(items),
            }
        )

    return render(request, 'planning/po_inbox.html', {'rows': rows})


@login_required
@permission_required('can_edit_jobcard')
def approval_queue(request):
    """Queue page to forward planning jobs to QC then Production Manager."""
    draft_jobs = PlanningJob.objects.filter(status__iexact='draft').order_by('-updated_at', '-id')[:300]
    reviewed_jobs = PlanningJob.objects.filter(status__iexact='reviewed').order_by('-updated_at', '-id')[:300]
    context = {
        'draft_jobs': draft_jobs,
        'reviewed_jobs': reviewed_jobs,
    }
    return render(request, 'planning/approval_queue.html', context)


@login_required
@permission_required('can_edit_jobcard')
def upload_po(request):
    """Upload a PO PDF, extract its content, store it, and redirect to review."""
    if request.method == 'POST':
        pdf_file = request.FILES.get('po_pdf')
        if not pdf_file:
            messages.error(request, 'Please select a PDF file.')
            return redirect('planning:upload_po')

        if not pdf_file.name.lower().endswith('.pdf'):
            messages.error(request, 'Only PDF files are supported.')
            return redirect('planning:upload_po')

        try:
            extracted = extract_po_from_pdf(pdf_file)
        except ValueError as exc:
            messages.error(request, f'Extraction failed: {exc}')
            return redirect('planning:upload_po')
        except Exception as exc:
            messages.error(request, f'Unexpected error reading PDF: {exc}')
            return redirect('planning:upload_po')

        # Surface partial-extraction warning before saving so user sees it on review page.
        extraction_warning = extracted.pop('extraction_warning', None)
        expected_count = extracted.pop('expected_line_count', None)

        items = extracted.get('items', [])
        source_item_count = len(items)
        deduped_items, duplicate_skus = _deduplicate_po_items_by_sku(items)
        extracted['items'] = deduped_items
        extracted['source_item_count'] = source_item_count
        extracted['merged_duplicate_skus'] = sorted(set(duplicate_skus))

        # Reset file pointer so Django can save it
        pdf_file.seek(0)

        po_number = (extracted.get('po_number') or '').strip()
        existing_doc = None
        if po_number:
            existing_doc = PoDocument.objects.filter(extracted_payload__po_number=po_number).order_by('-id').first()

        if existing_doc:
            existing_payload = existing_doc.extracted_payload or {}
            existing_items, _ = _deduplicate_po_items_by_sku(existing_payload.get('items', []))
            merged_items, added_skus, updated_skus, ignored_lines = _merge_po_items_for_existing_po(
                existing_items,
                deduped_items,
            )

            merged_payload = dict(existing_payload)
            merged_payload.update(extracted)
            merged_payload['items'] = merged_items
            merged_payload['source_item_count'] = len(merged_items)
            merged_payload['merged_duplicate_skus'] = sorted(
                set(existing_payload.get('merged_duplicate_skus') or []) | set(duplicate_skus)
            )
            configured_skus = sorted(set(existing_payload.get('new_skus_configured') or []))
            if configured_skus:
                merged_payload['new_skus_configured'] = configured_skus

            existing_doc.po_file = pdf_file
            existing_doc.extracted_payload = merged_payload
            existing_doc.extraction_status = 'processed'
            existing_doc.uploaded_by = request.user
            existing_doc.save(update_fields=['po_file', 'extracted_payload', 'extraction_status', 'uploaded_by'])
            po_doc = existing_doc
        else:
            po_doc = PoDocument.objects.create(
                po_file=pdf_file,
                extracted_payload=extracted,
                extraction_status='processed',
                uploaded_by=request.user,
            )

        sync_result = _sync_repeat_jobs_from_po(po_doc, actor=request.user)
        item_count = len(extracted.get('items', []))
        if existing_doc and ignored_lines:
            preview = ', '.join(
                f"{row['sku']} ({row['qty'] if row['qty'] is not None else '-'})"
                for row in ignored_lines[:8]
            )
            remainder = len(ignored_lines) - 8
            if remainder > 0:
                preview += f" +{remainder} more"
            messages.warning(
                request,
                f"Ignored duplicate line(s) for same PO (same SKU and Qty): {preview}.",
            )

        if extraction_warning:
            messages.warning(
                request,
                f"Partial extraction: {extraction_warning}",
            )
        else:
            if existing_doc:
                final_item_count = len((existing_doc.extracted_payload or {}).get('items', []))
                msg = (
                    f"PO {extracted.get('po_number', '?')} updated. "
                    f"Unique added SKU(s): {len(added_skus)}; corrected SKU(s): {len(updated_skus)}; "
                    f"current PO line count: {final_item_count}."
                )
                if duplicate_skus:
                    msg += f" Duplicate SKU lines merged in upload: {', '.join(sorted(set(duplicate_skus)))}."
                msg += " Same PO + same SKU + same Qty lines are ignored."
            else:
                msg = (
                    f"PO {extracted.get('po_number', '?')} extracted with "
                    f"{item_count} of {expected_count or item_count} line items. Sent to PO Intake queue."
                )
                if duplicate_skus:
                    msg += f" Duplicate SKU lines merged: {', '.join(sorted(set(duplicate_skus)))}."
            if sync_result['created'] or sync_result['updated']:
                msg += (
                    f" Repeat jobs sent to Planning: created {sync_result['created']}, "
                    f"updated {sync_result['updated']}."
                )
            if sync_result['missing_recipe']:
                msg += f" Repeat SKU(s) missing approved master data: {sync_result['missing_recipe']}."
            messages.success(request, msg)
        return redirect('planning:po_inbox')

    return render(request, 'planning/po_upload.html')


@login_required
@permission_required('can_edit_jobcard')
def po_review(request, doc_id):
    """Review extracted PO data and create PlanningJob records."""
    po_doc = get_object_or_404(PoDocument, id=doc_id)
    payload = po_doc.extracted_payload or {}
    items, duplicate_skus = _deduplicate_po_items_by_sku(payload.get('items', []))
    if duplicate_skus:
        payload['source_item_count'] = payload.get('source_item_count') or len(payload.get('items', []))
        payload['items'] = items
        payload['merged_duplicate_skus'] = sorted(set(duplicate_skus))
        po_doc.extracted_payload = payload
        po_doc.save(update_fields=['extracted_payload'])
        messages.info(
            request,
            f"Duplicate SKU lines were merged for this PO: {', '.join(sorted(set(duplicate_skus)))}.",
        )
    po_number = payload.get('po_number', '')
    configured_new_skus = {_sku_key(sku) for sku in (payload.get('new_skus_configured') or []) if sku}
    recipe_map = _build_recipe_map(items)
    annotated_items, repeat_count, new_count, missing_skus = _annotate_items_with_recipe(items, recipe_map)

    item_sku_keys = {_sku_key(item.get('sku')) for item in annotated_items if item.get('sku')}
    existing_jobs_by_sku = {}
    if po_number and item_sku_keys:
        existing_jobs = PlanningJob.objects.filter(po_number=po_number).order_by('-updated_at', '-id')
        for job in existing_jobs:
            key = _sku_key(job.sku)
            if key in item_sku_keys and key not in existing_jobs_by_sku:
                existing_jobs_by_sku[key] = job

    existing_any_jobs_skus = set()
    if item_sku_keys:
        sku_any_query = Q()
        for sku_key in item_sku_keys:
            sku_any_query |= Q(sku__iexact=sku_key)
        existing_any_jobs_skus = {
            _sku_key(sku)
            for sku in PlanningJob.objects.filter(sku_any_query).values_list('sku', flat=True)
            if sku
        }

    seen_skus_in_payload = set()
    for item in annotated_items:
        sku_key = _sku_key(item.get('sku'))
        is_first_production = bool(
            sku_key
            and sku_key not in existing_any_jobs_skus
            and sku_key not in seen_skus_in_payload
        )
        # Force first-ever job of an SKU as NEW; subsequent entries are REPEAT.
        item['forward_flag'] = 'New' if is_first_production else 'Repeat'
        item['is_first_production'] = is_first_production
        existing_job = existing_jobs_by_sku.get(sku_key)
        item['existing_job_id'] = existing_job.id if existing_job else None
        item['existing_jc_number'] = existing_job.jc_number if existing_job else ''
        recipe = recipe_map.get(sku_key)
        item['default_print_sheet_size'] = recipe.print_sheet_size if recipe else ''
        item['default_purchase_sheet_size'] = recipe.purchase_sheet_size if recipe else ''
        item['default_ups'] = recipe.ups if recipe else ''
        item['default_machine_name'] = recipe.machine_name if recipe else ''
        if sku_key:
            seen_skus_in_payload.add(sku_key)

    repeat_count = sum(1 for item in annotated_items if item.get('forward_flag') == 'Repeat')
    new_count = sum(1 for item in annotated_items if item.get('forward_flag') == 'New')

    if request.method == 'POST':
        action = request.POST.get('action', '')
        if action == 'create_jobs':
            with transaction.atomic():
                created_count = 0
                updated_count = 0
                skipped_count = 0
                locked_count = 0
                missing_recipe_count = 0
                po_date_raw = payload.get('po_date')
                po_date = _parse_iso_date(po_date_raw)
                delivery_location = payload.get('delivery_location', '')
                department = payload.get('department', '')

                for item in annotated_items:
                    sku = (item.get('sku') or '').strip()
                    job_name = (item.get('job_name') or '').strip() or sku
                    if not sku:
                        skipped_count += 1
                        continue

                    field_prefix = f"item_{item['line_no']}_"
                    skip_flag = request.POST.get(f"{field_prefix}skip") == '1'

                    if skip_flag:
                        skipped_count += 1
                        continue

                    recipe = recipe_map.get(_sku_key(sku))
                    if not recipe:
                        missing_recipe_count += 1
                        continue

                    sku_key = _sku_key(sku)
                    is_first_production = bool(
                        sku_key
                        and sku_key not in existing_any_jobs_skus
                    )
                    forward_as_new = is_first_production

                    delivery_date = _parse_iso_date(item.get('delivery_date'))
                    plan_date = delivery_date or po_date

                    # Re-use existing PO+SKU job to avoid duplicates; otherwise allocate a new serial.
                    existing_job = existing_jobs_by_sku.get(sku_key)

                    if existing_job:
                        if _normalize_status(existing_job.status) == 'approved':
                            locked_count += 1
                            continue
                        jc_number = existing_job.jc_number
                    else:
                        jc_number = allocate_next_jc_number(plan_date)

                    current_requirement = existing_job.requirement if existing_job else ''

                    qty = item.get('quantity')
                    order_qty = int(qty) if qty is not None else None

                    unit_cost_val = item.get('unit_cost')
                    unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None
                    override_print_sheet_size = (request.POST.get(f"{field_prefix}print_sheet_size") or '').strip()
                    override_purchase_sheet_size = (request.POST.get(f"{field_prefix}purchase_sheet_size") or '').strip()
                    override_machine_name = (request.POST.get(f"{field_prefix}machine_name") or '').strip()
                    override_ups = _to_optional_positive_int(request.POST.get(f"{field_prefix}ups"))

                    defaults = {
                        'po_number': po_number,
                        'sku': sku,
                        'job_name': recipe.job_name or job_name,
                        'order_qty': order_qty,
                        'department': recipe.department or department,
                        'destination': delivery_location,
                        'unit_cost': unit_cost_dec if unit_cost_dec is not None else recipe.default_unit_cost,
                        'status': 'draft',
                        'repeat_flag': 'New' if forward_as_new else 'Repeat',
                        'requirement': _sync_new_sku_requirement(current_requirement, forward_as_new),
                        'material': recipe.material,
                        'color_spec': recipe.color_spec,
                        'application': recipe.application,
                        'size_w_mm': recipe.size_w_mm,
                        'size_h_mm': recipe.size_h_mm,
                        'ups': override_ups if override_ups is not None else recipe.ups,
                        'print_sheet_size': override_print_sheet_size or recipe.print_sheet_size,
                        'purchase_sheet_size': override_purchase_sheet_size or recipe.purchase_sheet_size,
                        'purchase_material': recipe.purchase_material,
                        'machine_name': override_machine_name or recipe.machine_name,
                    }
                    if plan_date:
                        defaults['plan_date'] = plan_date

                    job_obj, created = PlanningJob.objects.update_or_create(
                        jc_number=jc_number,
                        defaults=defaults,
                    )
                    if created:
                        created_count += 1
                    else:
                        updated_count += 1
                    existing_jobs_by_sku[sku_key] = job_obj
                    if sku_key:
                        existing_any_jobs_skus.add(sku_key)

            po_doc.extraction_status = 'processed'
            po_doc.save(update_fields=['extraction_status'])

            if created_count == 0 and updated_count == 0:
                messages.warning(
                    request,
                    f'No jobs created. Skipped {skipped_count}, missing-recipe {missing_recipe_count}, locked-skip {locked_count}. Add missing SKU master data from Pending SKUs and run create again.',
                )
                return redirect('planning:pending_skus')

            messages.success(
                request,
                f'Done. Created {created_count}, updated {updated_count}, skipped {skipped_count}, missing-recipe {missing_recipe_count}, locked-skip {locked_count} planning job(s).',
            )
            if missing_recipe_count > 0:
                messages.warning(
                    request,
                    f'{missing_recipe_count} SKU(s) are still pending master data. Open Pending SKUs tab to configure them.',
                )
            return redirect('planning:approval_queue')

    context = {
        'po_doc': po_doc,
        'payload': payload,
        'items': annotated_items,
        'items_json': json.dumps(annotated_items),
        'repeat_count': repeat_count,
        'new_count': new_count,
        'configured_new_count': len(configured_new_skus),
        'missing_skus': missing_skus,
    }
    return render(request, 'planning/po_review.html', context)


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def po_new_skus(request, doc_id):
    po_doc = get_object_or_404(PoDocument, id=doc_id)
    payload = po_doc.extracted_payload or {}
    items = payload.get('items', [])

    recipe_map = _build_recipe_map(items)
    _, _, _, missing_skus = _annotate_items_with_recipe(items, recipe_map)
    missing_recipe_defaults = {}
    if missing_skus:
        recipe_query = Q()
        for sku in missing_skus:
            recipe_query |= Q(sku__iexact=sku)
        for recipe in SkuRecipe.objects.filter(recipe_query):
            missing_recipe_defaults[recipe.sku.upper()] = recipe
    missing_sku_rows = [
        {
            'sku': sku,
            'recipe': missing_recipe_defaults.get(_sku_key(sku)),
        }
        for sku in missing_skus
    ]

    if request.method == 'POST':
        created_count = 0
        saved_skus = []
        for sku in missing_skus:
            prefix = f"sku_{sku}"
            job_name = (request.POST.get(f"{prefix}_job_name") or '').strip()
            material = (request.POST.get(f"{prefix}_material") or '').strip()
            color_spec = (request.POST.get(f"{prefix}_color_spec") or '').strip()
            application = (request.POST.get(f"{prefix}_application") or '').strip()
            machine_name = (request.POST.get(f"{prefix}_machine_name") or '').strip()
            department = (request.POST.get(f"{prefix}_department") or '').strip()
            print_sheet_size = (request.POST.get(f"{prefix}_print_sheet_size") or '').strip()
            purchase_sheet_size = (request.POST.get(f"{prefix}_purchase_sheet_size") or '').strip()
            ups = _to_optional_positive_int(request.POST.get(f"{prefix}_ups"))
            purchase_material = (request.POST.get(f"{prefix}_purchase_material") or '').strip()

            unit_cost_raw = (request.POST.get(f"{prefix}_default_unit_cost") or '').strip()
            unit_cost = None
            if unit_cost_raw:
                try:
                    unit_cost = Decimal(unit_cost_raw)
                except InvalidOperation:
                    unit_cost = None

            if not job_name and not material and not machine_name:
                # Keep save requirements simple, but avoid empty recipe rows.
                continue

            SkuRecipe.objects.update_or_create(
                sku=sku,
                defaults={
                    'job_name': job_name,
                    'material': material,
                    'color_spec': color_spec,
                    'application': application,
                    'machine_name': machine_name,
                    'department': department,
                    'print_sheet_size': print_sheet_size,
                    'purchase_sheet_size': purchase_sheet_size,
                    'ups': ups,
                    'purchase_material': purchase_material,
                    'default_unit_cost': unit_cost,
                    'created_by': request.user,
                    'master_data_status': 'draft',
                    'reviewed_by': None,
                    'reviewed_at': None,
                    'approved_by': None,
                    'approved_at': None,
                },
            )
            created_count += 1
            saved_skus.append(sku)

        if saved_skus:
            configured = set(payload.get('new_skus_configured') or [])
            configured.update(saved_skus)
            payload['new_skus_configured'] = sorted(configured)
            po_doc.extracted_payload = payload
            po_doc.save(update_fields=['extracted_payload'])

        messages.success(
            request,
            f'SKU recipes saved: {created_count}. These SKU jobs will be forwarded as NEW for production shade/setup checks.',
        )
        return redirect('planning:po_review', doc_id=po_doc.id)

    return render(
        request,
        'planning/po_new_skus.html',
        {
            'po_doc': po_doc,
            'payload': payload,
            'missing_skus': missing_skus,
            'missing_sku_rows': missing_sku_rows,
            'missing_recipe_defaults': missing_recipe_defaults,
            'example_form': SkuRecipeForm(),
        },
    )
