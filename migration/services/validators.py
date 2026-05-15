import datetime
import itertools
from collections import Counter
from decimal import Decimal, InvalidOperation

from django.db.models.functions import Lower

from migration.models import PlanningImportStaging, RowImportStatus
from planning.models import PlanningJob


def _parse_int(value):
    if value is None or value == '':
        return None
    try:
        text = str(value).replace(',', '').strip()
        return int(text)
    except (TypeError, ValueError):
        return None


def _pick_raw_value(raw_data, keys):
    if not raw_data:
        return None
    for key in keys:
        if key in raw_data and raw_data[key] not in (None, ''):
            return raw_data[key]
    return None


def _normalize_jc_value(value):
    if value is None or value == '':
        return None
    text = str(value).strip()
    return text if text else None


def _chunked_iterable(iterable, size=200):
    iterator = iter(iterable)
    while True:
        chunk = list(itertools.islice(iterator, size))
        if not chunk:
            break
        yield chunk


def _parse_date(value):
    if value is None or value == '':
        return None
    if isinstance(value, datetime.date):
        return value
    text = str(value).strip()
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y'):
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def validate_planning_rows(import_job):
    """Validate staging rows individually without stopping the full batch."""
    rows = list(PlanningImportStaging.objects.filter(import_job=import_job).order_by('row_number', 'id'))
    valid_count = 0
    error_count = 0

    po_values = [row.po_number.strip().lower() for row in rows if row.po_number and row.po_number.strip()]
    sku_values = [row.sku.strip().lower() for row in rows if row.sku and row.sku.strip()]
    lower_po_set = set(po_values)
    lower_sku_set = set(sku_values)

    existing_job_keys = set()
    lookup_by_po = len(lower_po_set) <= len(lower_sku_set)

    if lookup_by_po:
        other_skus = lower_sku_set
        for po_chunk in _chunked_iterable(lower_po_set, size=200):
            candidates = PlanningJob.objects.annotate(
                lower_po=Lower('po_number'),
                lower_sku=Lower('sku'),
            ).filter(lower_po__in=list(po_chunk)).values_list('po_number', 'sku', 'jc_number', 'order_qty', 'print_sheet_size')
            for po, sku, jc, order_qty, print_sheet_size in candidates:
                sku_key = sku.strip().lower() if sku else None
                if sku_key not in other_skus:
                    continue
                po_key = po.strip().lower() if po else None
                jc_key = _normalize_jc_value(jc)
                print_size = print_sheet_size.strip().lower() if print_sheet_size else None
                if po_key and sku_key and jc_key is not None and order_qty is not None and print_size:
                    existing_job_keys.add((po_key, sku_key, jc_key, order_qty, print_size))
    else:
        other_pos = lower_po_set
        for sku_chunk in _chunked_iterable(lower_sku_set, size=200):
            candidates = PlanningJob.objects.annotate(
                lower_po=Lower('po_number'),
                lower_sku=Lower('sku'),
            ).filter(lower_sku__in=list(sku_chunk)).values_list('po_number', 'sku', 'jc_number', 'order_qty', 'print_sheet_size')
            for po, sku, jc, order_qty, print_sheet_size in candidates:
                po_key = po.strip().lower() if po else None
                if po_key not in other_pos:
                    continue
                sku_key = sku.strip().lower() if sku else None
                jc_key = _normalize_jc_value(jc)
                print_size = print_sheet_size.strip().lower() if print_sheet_size else None
                if po_key and sku_key and jc_key is not None and order_qty is not None and print_size:
                    existing_job_keys.add((po_key, sku_key, jc_key, order_qty, print_size))

    staged_keys = Counter()
    for row in rows:
        po_key = row.po_number.strip().lower() if row.po_number and row.po_number.strip() else None
        sku_key = row.sku.strip().lower() if row.sku and row.sku.strip() else None
        jc_key = _normalize_jc_value(_pick_raw_value(row.raw_data or {}, ['jc', 'jc_number', 'jc_no', 'job_card_no', 'jobcard_no', 'job_card', 'jobcard']))
        qty = _parse_int(row.quantity)
        print_size = _pick_raw_value(row.raw_data or {}, ['print_sheet_size'])
        print_size_key = print_size.strip().lower() if print_size else None
        if po_key and sku_key and jc_key is not None and qty is not None and print_size_key:
            staged_keys[(po_key, sku_key, jc_key, qty, print_size_key)] += 1

    for row in rows:
        errors = []
        po_number = (row.po_number or '').strip()
        customer = (row.customer or '').strip()
        sku = (row.sku or '').strip()

        if not po_number:
            errors.append('PO number is required.')
        if not customer:
            errors.append('Customer is required.')
        if not sku:
            errors.append('SKU is required.')

        qty = _parse_int(row.quantity)
        if qty is None:
            errors.append('Quantity must be a whole number.')
        elif qty <= 0:
            errors.append('Quantity must be greater than zero.')

        if row.delivery_date and _parse_date(row.delivery_date) is None:
            errors.append('Delivery date format is invalid.')

        raw = row.raw_data or {}
        if not _pick_raw_value(raw, ['jc', 'jc_number', 'jc_no', 'job_card_no', 'jobcard_no', 'job_card', 'jobcard']):
            errors.append('JC / Job Card No is required.')
        if not _pick_raw_value(raw, ['month', 'plan_month']):
            errors.append('Month is required.')
        if not _pick_raw_value(raw, ['date', 'plan_date', 'po_received_date', 'po_approval_date']):
            errors.append('Date is required.')
        if not _pick_raw_value(raw, ['job_name']):
            errors.append('Job Name is required.')
        if not _pick_raw_value(raw, ['repeat', 'repeat_flag']):
            errors.append('Repeat/New is required.')
        if not _pick_raw_value(raw, ['material']):
            errors.append('Material is required.')
        if not _pick_raw_value(raw, ['application']):
            errors.append('Application is required.')
        if not _pick_raw_value(raw, ['size_w_mm']):
            errors.append('Size W mm is required.')
        if not _pick_raw_value(raw, ['size_h_mm']):
            errors.append('Size H mm is required.')
        if not _pick_raw_value(raw, ['ups', 'no_of_ups']):
            errors.append('UPS is required.')
        if not _pick_raw_value(raw, ['print_sheet_size']):
            errors.append('Print Sheet Size is required.')
        if not _pick_raw_value(raw, ['actual_sheet_required', 'actual_sheet_require', 'sheet']):
            errors.append('Actual Sheet require is required.')
        if not _pick_raw_value(raw, ['purchase_sheet_size']):
            errors.append('Purchase Sheet Size is required.')
        if not _pick_raw_value(raw, ['purchase_sheet_ups']):
            errors.append('Purchase Sheet ups is required.')
        if not _pick_raw_value(raw, ['purchase_sheet_required', 'purchase_sheet_require']):
            errors.append('Purchase Sheet require is required.')
        if not _pick_raw_value(raw, ['destination', 'delivery_location']):
            errors.append('Destination is required.')
        if not _pick_raw_value(raw, ['department']):
            errors.append('Department is required.')
        if not _pick_raw_value(raw, ['wastage', 'wastage_sheets']):
            errors.append('Wastage is required.')
        # Die cutting is optional for this import flow.
        # Optional fields: no error if blank
        # AWC No., Plate Set No., Machine Name, Purchase Material, Stock, PKT, Remarks, Requirement, Die cutting

        row_jc = _normalize_jc_value(_pick_raw_value(raw, ['jc', 'jc_number', 'jc_no', 'job_card_no', 'jobcard_no', 'job_card', 'jobcard']))
        print_size = _pick_raw_value(raw, ['print_sheet_size'])
        print_size_key = print_size.strip().lower() if print_size else None
        if po_number and sku and row_jc is not None and qty is not None and print_size_key:
            duplicate_key = (po_number.lower(), sku.lower(), row_jc, qty, print_size_key)
            if duplicate_key in existing_job_keys:
                errors.append('Duplicate PO + SKU + JC + Quantity + Print Sheet Size found in planning jobs.')
            if staged_keys[duplicate_key] > 1:
                errors.append('Duplicate PO + SKU + JC + Quantity + Print Sheet Size found in staging rows.')

        if errors:
            row.import_status = RowImportStatus.ERROR
            row.error_message = ' | '.join(errors)
            error_count += 1
        else:
            row.import_status = RowImportStatus.VALID
            row.error_message = ''
            valid_count += 1

    for rows_chunk in _chunked_iterable(rows, size=150):
        PlanningImportStaging.objects.bulk_update(rows_chunk, ['import_status', 'error_message'])

    return {
        'total': len(rows),
        'valid': valid_count,
        'errors': error_count,
    }
