import io
import json
import re
from datetime import datetime, date
from decimal import Decimal
from collections import defaultdict
from difflib import SequenceMatcher
from django.db import transaction
from django.db.models import Sum, Q
from django.db.models.functions import Upper
from django.utils import timezone
from core.models import Machine, Department, Material
from .models import PLANNING_STATUS_ALIASES, PlanningJob, PoDocument, SkuRecipe
from workflow.services import _append_unique_note_line, _parse_iso_date, _format_display_qty, _build_cost_mismatch_note, _normalize_status, _to_int, _to_decimal
from core.jc_numbering import allocate_next_jc_number

NEW_SKU_REQUIREMENT_NOTE = 'NEW SKU: Shade matching and setup verification required before production run.'

def _user_is_admin(user):
    profile = getattr(user, 'profile', None)
    return getattr(user, 'is_superuser', False) or (profile is not None and getattr(profile, 'role', None) == 'admin')



def _planning_status_filter_values(status):
    normalized_status = _normalize_status(status, default='')
    if not normalized_status:
        return []
    return sorted(PLANNING_STATUS_ALIASES.get(normalized_status, {normalized_status}))



def _parse_date_filter(raw_value):
    if not raw_value:
        return None
    try:
        return datetime.strptime(str(raw_value).strip(), '%Y-%m-%d').date()
    except (ValueError, TypeError):
        return None


def _normalize_purchase_material_origin(raw_value):
    value = (raw_value or '').strip().lower()
    if value in {'local'}:
        return 'local'
    if value in {'import', 'imported'}:
        return 'import'
    return ''


def _normalize_po_number(raw_value):
    value = (raw_value or '').strip().upper()
    if not value:
        return ''
    if value.startswith('PO'):
        match = re.search(r'(\d+)$', value)
        if match:
            return match.group(1)
    return re.sub(r'[^A-Z0-9]+', '', value)


def _build_job_card_pdf_bytes(job, scan_url):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError('reportlab is required to generate PDF job cards. Install reportlab and restart the server.')
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    normal = styles['Normal']
    normal.fontName = 'Helvetica'
    normal.fontSize = 9
    normal.leading = 11

    title_style = ParagraphStyle('Title', parent=normal, fontName='Helvetica-Bold', fontSize=18, leading=20)
    subtitle_style = ParagraphStyle('Subtitle', parent=normal, fontName='Helvetica-Bold', fontSize=11, leading=13)
    section_title_style = ParagraphStyle('SectionTitle', parent=normal, fontName='Helvetica-Bold', fontSize=10.5, leading=12)
    label_style = ParagraphStyle('Label', parent=normal, fontName='Helvetica-Bold', fontSize=8.5, leading=10)

    story = [Paragraph('UTOPIA PRINTING & PACKAGING', title_style), Spacer(1, 4), Paragraph('PRODUCTION JOB CARD', subtitle_style), Spacer(1, 8)]

    header_data = [
        [Paragraph('JOB CARD #', label_style), _format_job_value(job.jc_number), Paragraph('PO #', label_style), _format_job_value(job.po_number)],
        [Paragraph('DATE', label_style), _format_job_value(job.plan_date), Paragraph('STATUS', label_style), _format_job_value(job.workflow_status_label)],
        [Paragraph('SKU', label_style), _format_job_value(job.sku), Paragraph('JOB NAME', label_style), _format_job_value(job.job_name)],
        [Paragraph('REPEAT FLAG', label_style), _format_job_value(job.repeat_flag), Paragraph('DEPARTMENT', label_style), _format_job_value(job.department)],
    ]
    header_table = Table(header_data, colWidths=[32 * mm, 65 * mm, 32 * mm, 65 * mm], hAlign='LEFT')
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.35, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dedede')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    story.extend([header_table, Spacer(1, 10)])

    material_data = [
        [Paragraph('ORDER QTY', label_style), _format_job_value(job.order_qty), Paragraph('PRINT PCS', label_style), _format_job_value(job.print_pcs)],
        [Paragraph('MATERIAL TYPE', label_style), _format_job_value(job.material), Paragraph('COLOR', label_style), _format_job_value(job.color_spec)],
        [Paragraph('APPLICATION', label_style), _format_job_value(job.application), Paragraph('PRINT SHEET SIZE', label_style), _format_job_value(job.print_sheet_size)],
        [Paragraph('UPS', label_style), _format_job_value(job.ups), Paragraph('PRINT SHEETS', label_style), _format_job_value(job.print_sheets)],
        [Paragraph('ACTUAL SHEETS', label_style), _format_job_value(job.calculated_sheets_required), Paragraph('WASTAGE', label_style), _format_job_value(job.wastage_sheets)],
        [Paragraph('PURCHASE ORIGIN', label_style), _format_job_value(job.purchase_material_origin), Paragraph('PURCHASE SHEET SIZE', label_style), _format_job_value(job.purchase_sheet_size)],
        [Paragraph('PURCHASE SHEET UPS', label_style), _format_job_value(job.purchase_sheet_ups), Paragraph('PURCHASE REQ', label_style), _format_job_value(job.purchase_sheet_required)],
        [Paragraph('MACHINE', label_style), _format_job_value(job.machine_name), Paragraph('TOTAL COLORS', label_style), _format_job_value(job.number_of_colors)],
        [Paragraph('PLATE SET NO.', label_style), _format_job_value(job.plate_set_no), Paragraph('AWC NO.', label_style), _format_job_value(job.awc_no_display)],
        [Paragraph('AGING DAYS', label_style), _format_job_value(job.aging_days), Paragraph('DIE CUTTING', label_style), _format_job_value(job.die_cutting_display)],
    ]
    material_table = Table(material_data, colWidths=[32 * mm, 65 * mm, 32 * mm, 65 * mm], hAlign='LEFT')
    material_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#eeeeee')),
    ]))
    story.extend([Paragraph('MATERIAL AND WORK PROCESS', section_title_style), Spacer(1, 4), material_table, Spacer(1, 10)])

    application_data = [
        [Paragraph('LAMINATION', label_style), _format_job_value(job.application), Paragraph('DIE CUTTING', label_style), _format_job_value(job.die_cutting_display)],
        [Paragraph('ART WORK NO.', label_style), '-', Paragraph('P SET NO.', label_style), _format_job_value(job.plate_set_no)],
        [Paragraph('SPECIAL INSTRUCTIONS', label_style), _paragraph_text(job.remarks or job.requirement or '-'), '', ''],
    ]
    application_table = Table(application_data, colWidths=[30 * mm, 67 * mm, 30 * mm, 65 * mm], hAlign='LEFT')
    application_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('INNERGRID', (0, 0), (-1, 1), 0.25, colors.grey),
        ('BOX', (0, 0), (-1, 1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f3f3f3')),
    ]))
    story.extend([application_table, Spacer(1, 12)])

    signature_data = [
        [Paragraph('Prepared by', label_style), '', Paragraph('Checked By', label_style), '', Paragraph('Plate Check By', label_style), '', Paragraph('Approved By', label_style), ''],
    ]
    signature_table = Table(signature_data, colWidths=[28 * mm, 34 * mm, 28 * mm, 34 * mm, 28 * mm, 34 * mm, 28 * mm, 34 * mm], hAlign='LEFT')
    signature_table.setStyle(TableStyle([
        ('LINEABOVE', (1, 0), (1, 0), 0.25, colors.black),
        ('LINEABOVE', (3, 0), (3, 0), 0.25, colors.black),
        ('LINEABOVE', (5, 0), (5, 0), 0.25, colors.black),
        ('LINEABOVE', (7, 0), (7, 0), 0.25, colors.black),
    ]))
    story.extend([signature_table, Spacer(1, 10)])

    material_issue_data = [[Paragraph('MATERIAL ISSUANCE', section_title_style), '', '', '', '', '']]
    material_issue_data.append([Paragraph('Date', label_style), Paragraph('Machine', label_style), Paragraph('Operator', label_style), Paragraph('Shift A/B', label_style), Paragraph('Sheet Size', label_style), Paragraph('Full Sheet Qty', label_style)])
    for _ in range(3):
        material_issue_data.append(['-', '-', '-', '-', '-', '-'])
    material_issue_table = Table(material_issue_data, colWidths=[24 * mm, 30 * mm, 35 * mm, 28 * mm, 35 * mm, 30 * mm], hAlign='LEFT')
    material_issue_table.setStyle(TableStyle([
        ('GRID', (0, 1), (-1, -1), 0.25, colors.grey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#d9d9d9')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    printing_data = [[Paragraph('PRINTING', section_title_style), '', '', '', '', '', '']]
    printing_data.append([Paragraph('Date', label_style), Paragraph('Machine', label_style), Paragraph('Operator', label_style), Paragraph('Shift A/B', label_style), Paragraph('Print Sheet Qty', label_style), Paragraph('Wastage Sheet', label_style), Paragraph('Half Good', label_style)])
    for _ in range(4):
        printing_data.append(['-', '-', '-', '-', '-', '-', '-'])
    printing_table = Table(printing_data, colWidths=[24 * mm, 30 * mm, 30 * mm, 28 * mm, 34 * mm, 34 * mm, 26 * mm], hAlign='LEFT')
    printing_table.setStyle(TableStyle([
        ('GRID', (0, 1), (-1, -1), 0.25, colors.grey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#d9d9d9')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    story.extend([material_issue_table, Spacer(1, 10), printing_table, Spacer(1, 12)])

    dispatch_data = [[Paragraph('DISPATCH', section_title_style), '', '', '', '', '']]
    dispatch_data.append([Paragraph('Delivery Date', label_style), Paragraph('DC #', label_style), Paragraph('Qty', label_style), Paragraph('Packing', label_style), Paragraph('Delivered To', label_style), ''])
    for _ in range(6):
        dispatch_data.append(['-', '-', '-', '-', '-', '-'])
    dispatch_table = Table(dispatch_data, colWidths=[30 * mm, 24 * mm, 24 * mm, 30 * mm, 35 * mm, 40 * mm], hAlign='LEFT')
    dispatch_table.setStyle(TableStyle([
        ('GRID', (0, 1), (-1, -1), 0.25, colors.grey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#d9d9d9')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    story.extend([dispatch_table, Spacer(1, 10)])

    cutting_data = [
        [Paragraph('CUTTING SLIP', section_title_style), '', '', '', '', ''],
        [Paragraph('Job Card #', label_style), _format_job_value(job.jc_number), Paragraph('Job Name', label_style), _format_job_value(job.job_name), Paragraph('Purch sheet size', label_style), _format_job_value(job.purchase_sheet_size)],
        [Paragraph('Purch sheet Ups', label_style), _format_job_value(job.purchase_sheet_ups), Paragraph('Print sheet size', label_style), _format_job_value(job.print_sheet_size), Paragraph('Type', label_style), _format_job_value(job.material)],
        [Paragraph('Purch sheet Qty', label_style), _format_job_value(job.purchase_sheet_required), Paragraph('Remarks', label_style), _paragraph_text(job.remarks or job.requirement or '-'), '', ''],
    ]
    cutting_table = Table(cutting_data, colWidths=[30 * mm, 35 * mm, 30 * mm, 35 * mm, 30 * mm, 35 * mm], hAlign='LEFT')
    cutting_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (-1, 0)),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 1), (-1, -1), 0.25, colors.grey),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f2f2f2')),
    ]))
    story.extend([cutting_table])

    doc.build(story)
    return buffer.getvalue()



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

    recipes = (
        SkuRecipe.objects
        .annotate(sku_upper=Upper('sku'))
        .filter(sku_upper__in=sku_values, master_data_status='approved')
    )
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



def _po_payload_items(payload, exclude_ignored=True):
    items = _sanitize_po_payload_items(payload)
    if not exclude_ignored:
        return items

    ignored_skus = {
        _sku_key(s)
        for s in (payload.get('new_skus_ignored') or [])
        if s
    }
    if not ignored_skus:
        return items

    return [
        item
        for item in items
        if _sku_key(item.get('sku')) not in ignored_skus
    ]



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

        existing_qty = _to_int(existing.get('quantity'))
        current_qty = _to_int(item_copy.get('quantity'))
        if existing_qty is None:
            existing['quantity'] = current_qty
        elif current_qty is not None:
            existing['quantity'] = existing_qty + current_qty

        existing_net = _to_decimal(existing.get('net_total'))
        current_net = _to_decimal(item_copy.get('net_total'))
        if existing_net is None:
            existing['net_total'] = _format_decimal_string(current_net)
        elif current_net is not None:
            existing['net_total'] = _format_decimal_string(existing_net + current_net)

        existing_subtotal = _to_decimal(existing.get('subtotal'))
        current_subtotal = _to_decimal(item_copy.get('subtotal'))
        if existing_subtotal is None:
            existing['subtotal'] = _format_decimal_string(current_subtotal)
        elif current_subtotal is not None:
            existing['subtotal'] = _format_decimal_string(existing_subtotal + current_subtotal)

        for field in ['job_name', 'delivery_date', 'unit', 'unit_cost']:
            if not existing.get(field) and item_copy.get(field):
                existing[field] = item_copy.get(field)

    deduped = [merged[key] for key in order]
    for idx, item in enumerate(deduped, start=1):
        item['line_no'] = idx
    return deduped, sorted(duplicate_skus)



def _history_repeat_new_counts(items, recipe_map=None):
    """Classify Repeat/New from approved SKU recipes, not historical PlanningJobs."""
    if recipe_map is None:
        recipe_map = _build_recipe_map(items)

    repeat_count = 0
    new_count = 0
    for item in items:
        sku = item.get('sku')
        sku_key = _sku_key(sku)
        if not sku_key:
            continue
        if sku_key in recipe_map:
            repeat_count += 1
        else:
            new_count += 1

    return repeat_count, new_count



def _sync_repeat_jobs_from_po(po_doc, actor=None):
    """Create or update draft planning jobs for all PO lines from one PO document."""
    payload = po_doc.extracted_payload or {}
    items, _ = _deduplicate_po_items_by_sku(payload.get('items', []))
    po_number = (payload.get('po_number') or '').strip()
    pr_number = (payload.get('pr_number') or '').strip()
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
    seen_skus_in_payload = set()

    for item in items:
        sku = (item.get('sku') or '').strip()
        sku_key = _sku_key(sku)
        if not sku_key:
            continue

        recipe = recipe_map.get(sku_key)
        if not recipe:
            missing_recipe_count += 1

        existing_job = existing_jobs_by_sku.get(sku_key)
        if existing_job and _normalize_status(existing_job.status) != 'draft':
            locked_count += 1
            continue

        delivery_date = _parse_iso_date(item.get('delivery_date'))
        plan_date = delivery_date or po_date
        qty = item.get('quantity')
        order_qty = int(qty) if qty is not None else None
        unit_cost_val = item.get('unit_cost')
        unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None
        jc_number = (
            (item.get('jc_number') or item.get('jc') or item.get('job_card_no') or item.get('jobcardno'))
            or (existing_job.jc_number if existing_job else None)
            or allocate_next_jc_number(plan_date)
        )
        is_first_production = bool(
            sku_key
            and sku_key not in existing_any_jobs_skus
            and sku_key not in seen_skus_in_payload
        )

        explicit_repeat_flag = (item.get('repeat_flag') or item.get('repeat') or '').strip()
        if explicit_repeat_flag.lower() in {'new', 'repeat'}:
            forward_as_new = explicit_repeat_flag.lower() == 'new'
        elif existing_job:
            existing_repeat_flag = (existing_job.repeat_flag or '').strip().lower()
            if existing_repeat_flag in {'new', 'repeat'}:
                forward_as_new = existing_repeat_flag == 'new'
            else:
                prior_jobs_exist = PlanningJob.objects.filter(sku__iexact=sku).exclude(id=existing_job.id).exists()
                forward_as_new = not prior_jobs_exist
        else:
            forward_as_new = is_first_production

        current_requirement = existing_job.requirement if existing_job else ''

        fallback_job_name = (item.get('job_name') or '').strip() or sku
        if item.get('job_name') and (item.get('job_name') or '').strip():
            job_name_value = item.get('job_name').strip()
        elif recipe and (recipe.job_name or '').strip():
            job_name_value = recipe.job_name
        elif existing_job and (existing_job.job_name or '').strip():
            job_name_value = existing_job.job_name
        else:
            job_name_value = fallback_job_name

        material_value = (item.get('material') or '').strip() or (recipe.material if recipe else (existing_job.material if existing_job else ''))
        color_spec_value = (item.get('color_spec') or item.get('color') or '').strip() or (recipe.color_spec if recipe else (existing_job.color_spec if existing_job else ''))
        application_value = (item.get('application') or '').strip() or (recipe.application if recipe else (existing_job.application if existing_job else ''))
        size_w_mm_value = _to_decimal(item.get('size_w_mm') or '') or (recipe.size_w_mm if recipe else (existing_job.size_w_mm if existing_job else None))
        size_h_mm_value = _to_decimal(item.get('size_h_mm') or '') or (recipe.size_h_mm if recipe else (existing_job.size_h_mm if existing_job else None))
        ups_value = _to_int(item.get('ups') or item.get('no_of_ups')) or (recipe.ups if recipe else (existing_job.ups if existing_job else None))
        print_sheet_size_value = (item.get('print_sheet_size') or '').strip() or (recipe.print_sheet_size if recipe else (existing_job.print_sheet_size if existing_job else ''))
        purchase_sheet_size_value = (item.get('purchase_sheet_size') or '').strip() or (recipe.purchase_sheet_size if recipe else (existing_job.purchase_sheet_size if existing_job else ''))
        purchase_sheet_ups_value = _to_int(item.get('purchase_sheet_ups') or '') or (recipe.purchase_sheet_ups if recipe else (existing_job.purchase_sheet_ups if existing_job else None))
        daily_demand_value = _to_decimal(item.get('daily_demand') or '') or (recipe.daily_demand if recipe else (existing_job.daily_demand if existing_job else None))
        unit_cost_value = unit_cost_dec if unit_cost_dec is not None else (recipe.default_unit_cost if recipe else (existing_job.unit_cost if existing_job else None))
        actual_sheet_required_value = _to_int(item.get('actual_sheet_required') or item.get('actual_sheet_require') or item.get('sheet')) or (existing_job.actual_sheet_required if existing_job else None)
        wastage_sheets_value = _to_int(item.get('wastage') or item.get('wastage_sheets')) or (existing_job.wastage_sheets if existing_job else None)
        purchase_sheet_required_value = _to_int(item.get('purchase_sheet_required') or item.get('purchase_sheet_require')) or (existing_job.purchase_sheet_required if existing_job else None)
        pkt_value = _to_decimal(item.get('pkt') or item.get('pkt_value') or '') or (existing_job.pkt_value if existing_job else None)
        stock_qty_value = _to_decimal(item.get('stock_qty') or item.get('stock') or '') or (existing_job.stock_qty if existing_job else None)
        balance_qty_value = _to_int(item.get('balance_qty') or item.get('balance') or '') or (existing_job.balance_qty if existing_job else None)
        plate_set_no_value = (item.get('plate_set_no') or item.get('p_set_no') or '').strip() or (existing_job.plate_set_no if existing_job else '')
        die_cutting_value = (item.get('die_cutting') or '').strip() or (recipe.die_cutting if recipe else (existing_job.die_cutting if hasattr(existing_job, 'die_cutting') else ''))
        purchase_material_origin_value = _normalize_purchase_material_origin(item.get('purchase_material_origin') or item.get('purchase_material') or '') or (existing_job.purchase_material_origin if existing_job else '')
        machine_name_value = (item.get('machine_name') or item.get('machine') or '').strip() or (existing_job.machine_name if existing_job else '')
        status_value = _normalize_status(item.get('status') or '') or 'draft'
        requirement_value = (item.get('requirement') or '').strip() or current_requirement

        if existing_job and not requirement_value:
            requirement_value = existing_job.requirement or ''

        requirement_value = _sync_new_sku_requirement(requirement_value, forward_as_new)
        if recipe and not forward_as_new:
            requirement_value = _append_unique_note_line(
                requirement_value,
                _build_cost_mismatch_note(recipe.default_unit_cost, unit_cost_dec),
            )

        defaults = {
            'po_number': po_number,
            'pr_reference': pr_number,
            'sku': sku,
            'job_name': job_name_value,
            'order_qty': order_qty,
            'department': department,
            'destination': delivery_location,
            'delivery_date': delivery_date,
            'unit_cost': unit_cost_value,
            'status': status_value,
            'repeat_flag': 'New' if forward_as_new else 'Repeat',
            'requirement': requirement_value,
            'material': material_value,
            'color_spec': color_spec_value,
            'application': application_value,
            'size_w_mm': size_w_mm_value,
            'size_h_mm': size_h_mm_value,
            'ups': ups_value,
            'print_sheet_size': print_sheet_size_value,
            'purchase_sheet_size': purchase_sheet_size_value,
            'purchase_sheet_ups': purchase_sheet_ups_value,
            'daily_demand': daily_demand_value,
            'plate_set_no': plate_set_no_value,
            'machine_name': machine_name_value,
            'actual_sheet_required': actual_sheet_required_value,
            'wastage_sheets': wastage_sheets_value,
            'purchase_sheet_required': purchase_sheet_required_value,
            'pkt_value': pkt_value,
            'remarks': (item.get('remarks') or '').strip() or (existing_job.remarks if existing_job else ''),
            'purchase_material_origin': purchase_material_origin_value,
            'stock_qty': stock_qty_value,
            'balance_qty': balance_qty_value,
        }
        if plan_date:
            defaults['plan_date'] = plan_date
        if payload.get('plan_month'):
            defaults['plan_month'] = payload.get('plan_month')
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
        existing_jobs_by_sku[sku_key] = job_obj
        existing_any_jobs_skus.add(sku_key)
        seen_skus_in_payload.add(sku_key)

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
    """After SKU master approval, refresh matching existing Planning Jobs only."""
    sku_key = _sku_key(sku)
    if not sku_key:
        return {'created': 0, 'updated': 0, 'locked': 0, 'sent': 0, 'missing_jobs': 0}

    recipe = SkuRecipe.objects.filter(sku__iexact=sku, master_data_status='approved').first()
    if not recipe:
        return {'created': 0, 'updated': 0, 'locked': 0, 'sent': 0, 'missing_jobs': 0}

    existing_any_jobs_skus = {
        _sku_key(value)
        for value in PlanningJob.objects.values_list('sku', flat=True)
        if value
    }

    created_count = 0
    updated_count = 0
    locked_count = 0
    sent_count = 0
    missing_job_count = 0

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
        if not existing_job:
            missing_job_count += 1
            continue

        if existing_job and _normalize_status(existing_job.status) != 'draft':
            locked_count += 1
            continue

        delivery_date = _parse_iso_date(target_item.get('delivery_date'))
        po_date = _parse_iso_date(payload.get('po_date'))
        plan_date = delivery_date or po_date
        qty = target_item.get('quantity')
        order_qty = int(qty) if qty is not None else None
        unit_cost_val = target_item.get('unit_cost')
        unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None
        jc_number = existing_job.jc_number
        current_requirement = existing_job.requirement

        existing_repeat_flag = (existing_job.repeat_flag or '').strip().lower()
        if existing_repeat_flag in {'new', 'repeat'}:
            forward_as_new = existing_repeat_flag == 'new'
        else:
            prior_jobs_exist = PlanningJob.objects.filter(sku__iexact=sku).exclude(id=existing_job.id).exists()
            forward_as_new = not prior_jobs_exist

        defaults = {
            'po_number': po_number,
            'sku': sku,
            'job_name': recipe.job_name or (target_item.get('job_name') or '').strip() or sku,
            'order_qty': order_qty,
            'department': payload.get('department') or '',
            'destination': payload.get('delivery_location') or '',
            'delivery_date': delivery_date,
            'unit_cost': unit_cost_dec if unit_cost_dec is not None else recipe.default_unit_cost,
            'status': 'draft',
            'repeat_flag': 'New' if forward_as_new else 'Repeat',
            'requirement': _sync_new_sku_requirement(current_requirement, forward_as_new),
            'material': recipe.material,
            'color_spec': recipe.color_spec,
            'application': recipe.application,
            'size_w_mm': recipe.size_w_mm,
            'size_h_mm': recipe.size_h_mm,
            'ups': recipe.ups,
            'print_sheet_size': recipe.print_sheet_size,
            'purchase_sheet_size': recipe.purchase_sheet_size,
            'purchase_sheet_ups': recipe.purchase_sheet_ups,
            'daily_demand': recipe.daily_demand,
            'plate_set_no': existing_job.plate_set_no,
        }

        if not forward_as_new:
            defaults['requirement'] = _append_unique_note_line(
                defaults['requirement'],
                _build_cost_mismatch_note(recipe.default_unit_cost, unit_cost_dec),
            )
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
        'missing_jobs': missing_job_count,
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



def _collect_pending_sku_rows(po_docs):
    """Build pending SKU rows from PO documents where SKU recipe is missing."""
    rows = []
    for po_doc in po_docs:
        payload = po_doc.extracted_payload or {}
        items = _po_payload_items(payload)
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
        ignored_skus = {
            _sku_key(s)
            for s in (payload.get('new_skus_ignored') or [])
            if s
        }
        for sku in missing_skus:
            if _sku_key(sku) in ignored_skus:
                continue
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



