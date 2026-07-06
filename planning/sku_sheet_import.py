"""Google Sheet / CSV restore for SkuRecipe master + planning job blanks (never wipe)."""

from __future__ import annotations

import csv
import io
from decimal import Decimal, InvalidOperation

from planning.forms import _normalize_application_value, _normalize_color_spec_value
from planning.models import PlanningJob, SkuRecipe
from planning.services import (
    _field_is_blank,
    _normalize_purchase_material_origin,
    normalize_awc_no,
)

# Google Sheet headers (and aliases) → SkuRecipe master fields.
SHEET_HEADER_TO_FIELD = {
    'SKU': 'sku',
    'JOB NAME': 'job_name',
    'JOBNAME': 'job_name',
    'Material': 'material',
    'MATERIAL': 'material',
    'Color': 'color_spec',
    'Colour': 'color_spec',
    'COLOUR': 'color_spec',
    'Application': 'application',
    'Product Type': 'product_type',
    'Size W mm': 'size_w_mm',
    'Size H mm': 'size_h_mm',
    'Ups': 'ups',
    'UPS': 'ups',
    'Print Sheet Size': 'print_sheet_size',
    'Purchase Sheet Size': 'purchase_sheet_size',
    'Purchase Sheet ups': 'purchase_sheet_ups',
    'Purchase Sheet Ups': 'purchase_sheet_ups',
    'Cost': 'default_unit_cost',
    'Default Unit Cost': 'default_unit_cost',
    'Daily Demand': 'daily_demand',
    'AWC No.': 'awc_no',
    'AWC No': 'awc_no',
    'AWC #': 'awc_no',
    'Die': 'die_cutting',
    'Die Cutting': 'die_cutting',
    # Google Sheet "Remarks" was intentionally stored on master as notes.
    'Notes': 'notes',
    'Remarks': 'notes',
    # Master now owns these (were planning-only); fill blanks from sheet when jobs lack them.
    'Machine': 'machine_name',
    'Machine Name': 'machine_name',
    'Plate Set No': 'plate_set_no',
    'Plate Set No.': 'plate_set_no',
    'Plate Set': 'plate_set_no',
}

# Sheet-only / display columns (not stored as-is on master).
SHEET_IGNORED_HEADERS = {
    'Sno.',
    'Sno',
    'Order Status',
    'Size W Inch',
    'Size H Inch',
}

# Sheet "Purchase Material" → planning job purchase_material_origin (not on SkuRecipe).
JOB_ORIGIN_HEADERS = ('Purchase Material', 'Purchase Material Origin', 'Purchase Origin')

INT_FIELDS = {'size_w_mm', 'size_h_mm', 'ups', 'purchase_sheet_ups'}
DECIMAL_FIELDS = {'default_unit_cost', 'daily_demand'}
# Master fields the Google Sheet may fill (blanks only by default).
RESTORE_FIELDS = [
    'job_name',
    'material',
    'color_spec',
    'application',
    'product_type',
    'machine_name',
    'plate_set_no',
    'size_w_mm',
    'size_h_mm',
    'ups',
    'print_sheet_size',
    'purchase_sheet_size',
    'purchase_sheet_ups',
    'default_unit_cost',
    'daily_demand',
    'awc_no',
    'die_cutting',
    'notes',
]


def _norm_header(value):
    return str(value or '').strip()


def _sheet_row_get(source, header):
    """Case-insensitive header lookup for DictReader / Excel rows."""
    if header in source:
        return source.get(header)
    target = header.casefold()
    for key, value in source.items():
        if _norm_header(key).casefold() == target:
            return value
    return None


def parse_sheet_rows(upload_file):
    """Return list of row dicts from CSV or XLSX upload file / path-like."""
    name = (getattr(upload_file, 'name', '') or '').lower()
    if hasattr(upload_file, 'read'):
        raw = upload_file.read()
        if isinstance(raw, str):
            raw = raw.encode('utf-8')
    else:
        with open(upload_file, 'rb') as handle:
            raw = handle.read()
            name = str(upload_file).lower()

    rows = []
    if name.endswith('.csv') or (not name.endswith('.xlsx') and b',' in raw[:200]):
        decoded = raw.decode('utf-8-sig')
        reader = csv.DictReader(io.StringIO(decoded))
        for row in reader:
            rows.append({_norm_header(k): v for k, v in row.items() if _norm_header(k)})
        return rows

    if name.endswith('.xlsx'):
        import openpyxl

        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        ws = wb.active
        header_row_idx = None
        header = []
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
            values = [_norm_header(c) for c in row]
            if 'SKU' in values and any(v.upper() in {'JOB NAME', 'JOBNAME'} for v in values):
                header_row_idx = i
                header = values
                break
        if not header_row_idx:
            raise ValueError('Could not find header row (need SKU and JOB NAME).')
        for values in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            row = {}
            for idx, key in enumerate(header):
                if key:
                    row[key] = values[idx] if idx < len(values) else None
            rows.append(row)
        return rows

    raise ValueError('Unsupported file type. Use CSV or XLSX.')


def _clean_int(value):
    if value is None or str(value).strip() == '':
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def _clean_decimal(value):
    if value is None or str(value).strip() == '':
        return None
    try:
        return Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        return None


def _is_none_color(value):
    text = str(value or '').strip().lower()
    return text in {'', 'no', 'none', 'n/a', 'na', 'nil', 'not applicable', '-'}


def row_to_field_values(source):
    """Map one sheet row to model field values (only non-blank sheet cells)."""
    payload = {}
    for header, field in SHEET_HEADER_TO_FIELD.items():
        if field == 'sku':
            continue
        raw = _sheet_row_get(source, header)
        if raw is None or str(raw).strip() == '':
            continue

        if field == 'application':
            value = _normalize_application_value(raw)
            if value:
                payload[field] = value
            continue

        if field == 'color_spec':
            if _is_none_color(raw):
                # Sheet "No" means no printing — do not force invalid print-color master values.
                continue
            value = _normalize_color_spec_value(raw)
            if value and not _is_none_color(value):
                payload[field] = value
            continue

        if field in INT_FIELDS:
            value = _clean_int(raw)
            if value is not None:
                payload[field] = value
            continue

        if field in DECIMAL_FIELDS:
            value = _clean_decimal(raw)
            if value is not None:
                payload[field] = value
            continue

        if field == 'awc_no':
            value = normalize_awc_no(raw)
            if value:
                payload[field] = value
            continue

        text = str(raw).strip()
        if field == 'plate_set_no':
            # Excel often exports 1499.0 — plate set stays numeric-looking string.
            try:
                text = str(int(float(text)))
            except (TypeError, ValueError):
                pass
        payload[field] = text

    # Cut & Pack heuristic from sheet Color/Application = No
    color_raw = _sheet_row_get(source, 'Color')
    if color_raw is None:
        color_raw = _sheet_row_get(source, 'Colour')
    app_raw = _sheet_row_get(source, 'Application')
    if _is_none_color(color_raw) and _normalize_application_value(app_raw) == 'NO':
        payload.setdefault('job_process_type', 'cut_and_pack')

    return payload


def row_to_job_origin(source):
    """Sheet Purchase Material → ERP purchase_material_origin (local|import)."""
    for header in JOB_ORIGIN_HEADERS:
        raw = _sheet_row_get(source, header)
        if raw is None or str(raw).strip() == '':
            continue
        origin = _normalize_purchase_material_origin(raw)
        if origin:
            return origin
    return ''


def apply_sheet_values_to_recipe(recipe, values, *, fill_blanks_only=True):
    """
    Apply sheet values onto a recipe.

    fill_blanks_only=True (default): only write when recipe field is blank.
    Returns list of field names updated.
    """
    updated = []
    for field_name, value in values.items():
        if field_name not in RESTORE_FIELDS and field_name != 'job_process_type':
            continue
        if _field_is_blank(value):
            continue
        current = getattr(recipe, field_name, None)
        if fill_blanks_only and not _field_is_blank(current):
            continue
        if current == value:
            continue
        setattr(recipe, field_name, value)
        updated.append(field_name)
    return updated


# Jobs past release / production must never be mutated by sheet restore.
EXECUTED_JOB_STATUSES = frozenset({
    'released',
    'in_production',
    'completed',
})


def planning_job_is_executed(job):
    """True when the job has already left editable planning (do not change)."""
    if not job:
        return False
    status = (getattr(job, 'workflow_status', None) or getattr(job, 'status', None) or '').strip()
    if status in EXECUTED_JOB_STATUSES:
        return True
    stage = (getattr(job, 'planning_stage', None) or '').strip()
    if stage == 'planning_done':
        return True
    # Job card exists — treat as executed even if status is lagging.
    try:
        if job.job_card is not None:
            return True
    except PlanningJob.job_card.RelatedObjectDoesNotExist:
        pass
    except Exception:
        pass
    return False


def apply_purchase_origin_to_jobs(sku, origin, *, fill_blanks_only=True):
    """Fill blank purchase_material_origin on non-executed planning jobs only."""
    if not sku or not origin:
        return 0
    jobs = PlanningJob.objects.filter(sku__iexact=sku).select_related('job_card')
    updated_jobs = 0
    for job in jobs.iterator():
        if planning_job_is_executed(job):
            continue
        current = (job.purchase_material_origin or '').strip()
        if fill_blanks_only and current:
            continue
        if current == origin:
            continue
        job.purchase_material_origin = origin
        job.save(update_fields=['purchase_material_origin', 'updated_at'])
        updated_jobs += 1
    return updated_jobs


def restore_sku_recipes_from_rows(rows, *, fill_blanks_only=True, create_missing=False, user=None):
    """
    Restore blank SkuRecipe fields and planning-job purchase origin from sheet rows.

    Does not demote approved recipes. Does not overwrite non-blank fields when
    fill_blanks_only is True.

    Mapping notes:
    - Remarks → notes (master)
    - AWC No → awc_no (master)
    - Machine / Plate Set No → master blanks (also hydrate from jobs separately)
    - Purchase Material → purchase_material_origin on planning jobs (local|import)
    """
    created = 0
    updated = 0
    skipped = 0
    missing_sku = 0
    jobs_origin_updated = 0
    field_hits = {}
    samples = []

    for source in rows:
        sku = str(_sheet_row_get(source, 'SKU') or '').strip()
        if not sku:
            continue
        values = row_to_field_values(source)
        job_name = values.get('job_name') or str(_sheet_row_get(source, 'JOB NAME') or '').strip()
        if job_name:
            values['job_name'] = job_name
        origin = row_to_job_origin(source)

        recipe = SkuRecipe.objects.filter(sku__iexact=sku).first()
        recipe_changed = []
        if not recipe:
            if not create_missing:
                missing_sku += 1
                if origin:
                    jobs_origin_updated += apply_purchase_origin_to_jobs(
                        sku, origin, fill_blanks_only=fill_blanks_only,
                    )
                continue
            recipe = SkuRecipe(sku=sku, created_by=user, master_data_status='draft', legacy_produced=True)
            if job_name:
                recipe.job_name = job_name
            recipe_changed = apply_sheet_values_to_recipe(recipe, values, fill_blanks_only=False)
            recipe.save()
            created += 1
            for name in recipe_changed:
                field_hits[name] = field_hits.get(name, 0) + 1
            if len(samples) < 20:
                samples.append((recipe.pk, sku, 'created', recipe_changed))
        else:
            recipe_changed = apply_sheet_values_to_recipe(
                recipe, values, fill_blanks_only=fill_blanks_only,
            )
            if not recipe.legacy_produced:
                recipe.legacy_produced = True
                recipe_changed = list(recipe_changed) + ['legacy_produced']
            if recipe_changed:
                recipe.save(update_fields=list(dict.fromkeys(recipe_changed + ['updated_at'])))
                updated += 1
                for name in recipe_changed:
                    field_hits[name] = field_hits.get(name, 0) + 1
                if len(samples) < 20:
                    samples.append((recipe.pk, sku, recipe.master_data_status, recipe_changed))
            else:
                skipped += 1

        if origin:
            n = apply_purchase_origin_to_jobs(sku, origin, fill_blanks_only=fill_blanks_only)
            jobs_origin_updated += n
            if n:
                field_hits['purchase_material_origin'] = (
                    field_hits.get('purchase_material_origin', 0) + n
                )

    return {
        'created': created,
        'updated': updated,
        'skipped': skipped,
        'missing_sku': missing_sku,
        'jobs_origin_updated': jobs_origin_updated,
        'field_hits': field_hits,
        'samples': samples,
    }
