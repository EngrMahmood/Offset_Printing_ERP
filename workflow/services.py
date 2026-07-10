import csv
import io
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher

from django.core.exceptions import ValidationError
from django.db import transaction
from django.db.models import Q

from core.jc_numbering import allocate_next_jc_number
from core.jobcard_service import (
    execute_job_card_action,
    release_to_production,
    reject_planning,
    reject_qc,
    approve_pm,
    approve_qc,
    submit_to_qc,
    ensure_job_card_from_planning_job,
    start_production as _start_production,
    complete_production as _complete_production,
    close_job_card as _close_job_card,
)
from core.models import JOB_CARD_PLANNING_EDITABLE_STATUSES
from planning.models import (
    PLANNING_STATUS_ALIASES,
    PLANNING_STATUS_CHOICES,
    PlanningJob,
    PoDocument,
    SkuRecipe,
)

PLANNING_STATUS_SET = {value for value, _ in PLANNING_STATUS_CHOICES}


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
    value = _to_decimal(raw_value)
    if value is None:
        return raw_value if raw_value not in (None, '') else '-'

    if value == value.to_integral_value():
        return str(int(value))

    normalized = value.normalize()
    text = format(normalized, 'f').rstrip('0').rstrip('.')
    return text or '0'


def _format_decimal_string(raw_value):
    if raw_value is None:
        return None
    value = _to_decimal(raw_value)
    if value is None:
        return None
    if value == value.to_integral_value():
        return str(int(value))
    normalized = value.normalize()
    text = format(normalized, 'f').rstrip('0').rstrip('.')
    return text or '0'


def _normalize_color_spec_input(raw_value):
    from core.print_colors import normalize_color_spec_value
    return normalize_color_spec_value(raw_value)


def _normalize_application_input(raw_value):
    value = str(raw_value or '').strip()
    if not value:
        return ''
    lowered = value.lower()
    if lowered in {'no', 'none', 'n/a', 'na', 'nil', 'not applicable'}:
        return 'NO'
    if 'uv' in lowered or 'u.v' in lowered:
        return 'UV'
    if 'matt' in lowered or 'matte' in lowered:
        return 'Lamination Matt'
    if 'lamination' in lowered or 'lam' in lowered or 'lamin' in lowered:
        return 'Lamination Gloss'
    if 'gloss' in lowered or 'shine' in lowered:
        return 'Lamination Gloss'
    if 'varnish' in lowered or 'op' in lowered:
        return 'NO'
    return 'NO'


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
    return f"COST ALERT: PO unit cost {po} differs from master default {master}. PO cost is applied to this job."


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
    value = PLANNING_STATUS_ALIASES.get(value, value)
    if value in PLANNING_STATUS_SET:
        return value
    return default


def _sku_key(sku):
    return (sku or '').strip().upper()


def _has_letters_and_digits(value):
    text = str(value or '')
    return any(ch.isalpha() for ch in text) and any(ch.isdigit() for ch in text)

SKU_MASTER_APPROVAL_REQUIRED_FIELDS = [
    ('job_name', 'Job Name'),
    ('material', 'Material'),
    ('color_spec', 'Print Color'),
    ('application', 'Application'),
    ('product_type', 'Product Type'),
    ('size_w_mm', 'Size Width (mm)'),
    ('size_h_mm', 'Size Height (mm)'),
    ('print_sheet_size', 'Print Sheet'),
    ('purchase_sheet_size', 'Purchase Sheet'),
    ('ups', 'UPS'),
    ('purchase_sheet_ups', 'Purchase Sheet Ups'),
    ('awc_no', 'AWC #'),
    ('die_cutting', 'Die Cutting'),
    ('plate_set_no', 'Plate Set No.'),
    ('print_passes', 'No. of Passes'),
]


def _missing_required_master_fields(recipe, fallback_job_name='', *, allow_missing_plate_set_no=False):
    missing = []
    cut_and_pack = bool(
        recipe and (getattr(recipe, 'job_process_type', '') or 'print_and_pack') == 'cut_and_pack'
    )
    skip_for_cut_and_pack = {'color_spec', 'print_passes'}

    if not recipe:
        fallback = (fallback_job_name or '').strip()
        return [
            label
            for field, label in SKU_MASTER_APPROVAL_REQUIRED_FIELDS
            if not (field == 'job_name' and fallback)
            and not (allow_missing_plate_set_no and field == 'plate_set_no')
        ]

    for field, label in SKU_MASTER_APPROVAL_REQUIRED_FIELDS:
        if cut_and_pack and field in skip_for_cut_and_pack:
            continue
        if allow_missing_plate_set_no and field == 'plate_set_no':
            continue
        value = getattr(recipe, field, None)
        if isinstance(value, str):
            if not value.strip():
                missing.append(label)
        elif value is None:
            missing.append(label)
    return missing


PLATE_SET_NO_WARNING_LABEL = 'Plate Set No.'


def _warning_master_fields(recipe, fallback_job_name=''):
    del fallback_job_name
    if not recipe:
        return [PLATE_SET_NO_WARNING_LABEL]
    value = getattr(recipe, 'plate_set_no', None)
    if isinstance(value, str):
        if value.strip():
            return []
    elif value is not None:
        return []
    return [PLATE_SET_NO_WARNING_LABEL]


def _build_recipe_map(items):
    """Return a map of SKU-upper -> SkuRecipe for any existing recipe (any status).
    Priority: approved > reviewed > pending_review > draft.
    """
    sku_values = sorted({_sku_key(item.get('sku')) for item in items if item.get('sku')})
    if not sku_values:
        return {}

    from django.db.models.functions import Upper
    STATUS_PRIORITY = {'approved': 0, 'reviewed': 1, 'pending_review': 2, 'draft': 3}

    recipes = (
        SkuRecipe.objects
        .annotate(sku_upper=Upper('sku'))
        .filter(sku_upper__in=sku_values)
        .order_by('sku_upper')
    )

    result = {}
    for recipe in recipes:
        key = recipe.sku.upper()
        if key not in result:
            result[key] = recipe
        else:
            existing_priority = STATUS_PRIORITY.get(result[key].master_data_status, 99)
            incoming_priority = STATUS_PRIORITY.get(recipe.master_data_status, 99)
            if incoming_priority < existing_priority:
                result[key] = recipe
    return result


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


def _deduplicate_po_items_by_sku(items):
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


def _sanitize_po_payload_items(payload):
    items, _ = _deduplicate_po_items_by_sku((payload or {}).get('items', []))

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
            if similar and sku_norm != ex_norm and _has_letters_and_digits(sku_norm) and _has_letters_and_digits(ex_norm):
                continue
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


def _annotate_items_with_recipe(items, recipe_map, current_po_number=None, po_doc_created_at=None, po_doc_id=None):
    from planning.sku_classification import annotate_items_repeat_new

    return annotate_items_repeat_new(
        items,
        recipe_map,
        po_number=current_po_number,
        po_doc_created_at=po_doc_created_at,
        po_doc_id=po_doc_id,
    )


def _collect_pending_sku_rows(po_docs):
    """Build pending SKU rows from PO documents where SKU master is not yet approved."""
    rows = []
    for po_doc in po_docs:
        payload = po_doc.extracted_payload or {}
        items = _po_payload_items(payload)
        if not items:
            continue

        recipe_map = _build_recipe_map(items)
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
        seen_skus = set()
        for item in items:
            sku = (item.get('sku') or '').strip()
            key = _sku_key(sku)
            if not key or key in ignored_skus or key in seen_skus:
                continue
            seen_skus.add(key)

            recipe = recipe_map.get(key)
            is_approved = bool(recipe and recipe.master_data_status == 'approved')
            if is_approved:
                continue

            item = item_map.get(key, {})
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


def _sync_new_sku_requirement(existing_requirement, is_new):
    lines = [line.strip() for line in str(existing_requirement or '').splitlines() if line.strip()]
    filtered_lines = [line for line in lines if line != 'NEW SKU: Shade matching and setup verification required before production run.']

    if is_new:
        return '\n'.join(['NEW SKU: Shade matching and setup verification required before production run.'] + filtered_lines)
    return '\n'.join(filtered_lines)


def sync_job_card_for_planning_status(job, target_status, actor):
    if target_status != 'pending_qc':
        return None

    job_card, _created = ensure_job_card_from_planning_job(job, actor=actor)
    if job_card.workflow_status == 'pm_rejected':
        execute_job_card_action(job_card, 'reopen', actor=actor, reason='Reopening rejected job card for QC resubmit')
        submit_to_qc(job_card, actor=actor, reason='Planning job sent to QC after reopen')
        return job_card

    if job_card.workflow_status in JOB_CARD_PLANNING_EDITABLE_STATUSES:
        submit_to_qc(job_card, actor=actor, reason='Planning job sent to QC')
    elif job_card.workflow_status == 'planning_approved':
        submit_to_qc(job_card, actor=actor, reason='Planning job sent to QC')
    elif job_card.workflow_status == 'qc_rejected':
        submit_to_qc(job_card, actor=actor, reason='Reopened QC rejected job card for QC resubmit')
    return job_card


def _sync_new_jobs_for_approved_sku(sku, actor=None):
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
        if existing_job and _normalize_status(existing_job.status) != 'draft':
            locked_count += 1
            continue

        delivery_date = _parse_iso_date(target_item.get('delivery_date'))
        po_date = _parse_iso_date(payload.get('po_date'))
        plan_date = delivery_date or po_date
        qty = target_item.get('quantity')
        order_qty = _to_int(qty)
        unit_cost_val = target_item.get('unit_cost')
        unit_cost_dec = _to_decimal(unit_cost_val)

        if existing_job:
            jc_number = existing_job.jc_number
            current_requirement = existing_job.requirement
        else:
            jc_number = allocate_next_jc_number(plan_date)
            current_requirement = ''

        from planning.sku_classification import repeat_flag_value_for_po_line

        repeat_flag_value = repeat_flag_value_for_po_line(
            target_item,
            po_number=po_number,
            po_doc_created_at=getattr(po_doc, 'created_at', None),
            po_doc_id=getattr(po_doc, 'id', None),
            recipe=recipe,
            existing_job=existing_job,
        )
        forward_as_new = repeat_flag_value == 'New'

        defaults = {
            'po_number': po_number,
            'sku': sku,
            'job_name': recipe.job_name or (target_item.get('job_name') or '').strip() or sku,
            'order_qty': order_qty,
            'department': payload.get('department') or '',
            'destination': payload.get('delivery_location') or '',
            'unit_cost': unit_cost_dec if unit_cost_dec is not None else recipe.default_unit_cost,
            'status': 'draft',
            'repeat_flag': repeat_flag_value,
            'requirement': _sync_new_sku_requirement(current_requirement, forward_as_new),
            'material': recipe.material,
            'job_process_type': recipe.job_process_type or 'print_and_pack',
            'color_spec': recipe.color_spec,
            'application': recipe.application,
            'size_w_mm': recipe.size_w_mm,
            'size_h_mm': recipe.size_h_mm,
            'ups': recipe.ups,
            'print_sheet_size': recipe.print_sheet_size,
            'purchase_sheet_size': recipe.purchase_sheet_size,
            'purchase_sheet_ups': recipe.purchase_sheet_ups,
            'daily_demand': recipe.daily_demand,
            'plate_set_no': existing_job.plate_set_no if existing_job else (recipe.plate_set_no or ''),
            'machine_name': recipe.machine_name,
            'remarks': (target_item.get('remarks') or '').strip() or (existing_job.remarks if existing_job else '') or recipe.notes or '',
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


def start_production(job_card, actor=None, reason=''):
    return _start_production(job_card, actor=actor, reason=reason)


def complete_production(job_card, actor=None, reason=''):
    return _complete_production(job_card, actor=actor, reason=reason)


def close_job_card(job_card, actor=None, reason=''):
    return _close_job_card(job_card, actor=actor, reason=reason)


def _user_is_admin(user):
    profile = getattr(user, 'profile', None)
    return getattr(user, 'is_superuser', False) or (profile is not None and getattr(profile, 'role', None) == 'admin')


