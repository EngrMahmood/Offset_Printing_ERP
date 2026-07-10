"""Google Sheet / CSV restore for SkuRecipe master + planning job blanks (never wipe)."""

from __future__ import annotations

import csv
import io
from decimal import Decimal, InvalidOperation

from planning.forms import _normalize_application_value, _normalize_color_spec_value
from planning.sku_migration_phases import (
    build_sku_phase_map,
    get_sku_migration_phase,
    infer_product_type_from_sku,
    sku_eligible_for_migration_phase,
)
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
    'Front Pass',
    'Back Pass',
    'No. Of Clrs Back',
    'No. of Clrs Front',
    'Total Crls',
    'Total M/R Time (15m/clr)',
}

# Sheet "Purchase Material" → planning job purchase_material_origin (not on SkuRecipe).
JOB_ORIGIN_HEADERS = ('Purchase Material', 'Purchase Material Origin', 'Purchase Origin')

INT_FIELDS = {'size_w_mm', 'size_h_mm', 'ups', 'purchase_sheet_ups', 'print_passes'}
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
    'print_passes',
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


def _find_sheet_header(all_rows):
    """Locate header row containing SKU and JOB NAME."""
    for i, row in enumerate(all_rows[:15]):
        values = [_norm_header(c) for c in row]
        if 'SKU' in values and any(v.upper() in {'JOB NAME', 'JOBNAME'} for v in values):
            return i, values
    raise ValueError('Could not find header row (need SKU and JOB NAME).')


def _tabular_rows(header, data_rows):
    rows = []
    for values in data_rows:
        row = {}
        for idx, key in enumerate(header):
            if key:
                row[key] = values[idx] if idx < len(values) else None
        rows.append(row)
    return rows


def _normalize_job_process_type(raw):
    text = str(raw or '').strip().lower()
    if not text:
        return ''
    if 'cut' in text and 'pack' in text:
        return 'cut_and_pack'
    if 'print' in text and 'pack' in text:
        return 'print_and_pack'
    return ''


def _clean_print_passes(value):
    passes = _clean_int(value)
    if passes in {1, 2, 3, 4}:
        return passes
    return None


def dedupe_sheet_rows(rows):
    """Keep one row per SKU (case-insensitive); later rows win."""
    by_sku = {}
    order = []
    for row in rows:
        sku = str(_sheet_row_get(row, 'SKU') or '').strip()
        if not sku:
            continue
        key = sku.casefold()
        if key not in by_sku:
            order.append(key)
        by_sku[key] = row
    return [by_sku[key] for key in order]


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
    if name.endswith('.csv') or (
        not name.endswith('.xlsx') and not name.endswith('.xlsb') and b',' in raw[:200]
    ):
        decoded = raw.decode('utf-8-sig')
        reader = csv.DictReader(io.StringIO(decoded))
        for row in reader:
            rows.append({_norm_header(k): v for k, v in row.items() if _norm_header(k)})
        return rows

    if name.endswith('.xlsx'):
        import openpyxl

        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        ws = wb.active
        all_rows = [list(values) for values in ws.iter_rows(values_only=True)]
        header_idx, header = _find_sheet_header(all_rows)
        return _tabular_rows(header, all_rows[header_idx + 1:])

    if name.endswith('.xlsb'):
        from pyxlsb import open_workbook

        with open_workbook(io.BytesIO(raw)) as wb:
            with wb.get_sheet(wb.sheets[0]) as ws:
                all_rows = [
                    [cell.v if hasattr(cell, 'v') else cell for cell in row]
                    for row in ws.rows()
                ]
        header_idx, header = _find_sheet_header(all_rows)
        return _tabular_rows(header, all_rows[header_idx + 1:])

    raise ValueError('Unsupported file type. Use CSV, XLSX, or XLSB.')


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

    job_process = _normalize_job_process_type(_sheet_row_get(source, 'Job Process'))
    if job_process:
        payload['job_process_type'] = job_process

    if payload.get('job_process_type') == 'cut_and_pack':
        payload['print_passes'] = None
    else:
        passes = _clean_print_passes(_sheet_row_get(source, 'No. of Passes'))
        if passes is not None:
            payload['print_passes'] = passes

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


def _job_process_fill_allowed(current, new_value, *, fill_blanks_only):
    """Allow correcting default print_and_pack → cut_and_pack when filling blanks."""
    current = (current or 'print_and_pack').strip()
    if current == new_value:
        return False
    if fill_blanks_only:
        if current == 'print_and_pack' and new_value == 'cut_and_pack':
            return True
        return _field_is_blank(current)
    return True


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
        if field_name == 'job_process_type':
            if not value:
                continue
            current = getattr(recipe, field_name, None)
            if not _job_process_fill_allowed(current, value, fill_blanks_only=fill_blanks_only):
                continue
            setattr(recipe, field_name, value)
            updated.append(field_name)
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

    if (getattr(recipe, 'job_process_type', None) or 'print_and_pack') == 'cut_and_pack':
        if recipe.print_passes is not None:
            recipe.print_passes = None
            if 'print_passes' not in updated:
                updated.append('print_passes')

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


def restore_sku_recipes_from_rows(
    rows,
    *,
    fill_blanks_only=True,
    create_missing=False,
    user=None,
    migration_phase=None,
    infer_product_type=False,
):
    """
    Restore blank SkuRecipe master fields from sheet rows.

    Does not demote approved recipes. Does not overwrite non-blank fields when
    fill_blanks_only is True.

    migration_phase: 1, 2, or 3 — only touch SKUs in that safety tier (see
    planning.sku_migration_phases). None = all phases (legacy behaviour).

    infer_product_type: when True with migration_phase 1 or 2, fill blank product_type
    from SKU prefix rules for ERP-only SKUs not covered by the sheet.

    Mapping notes:
    - Remarks → notes (master)
    - AWC No → awc_no (master)
    - Machine / Plate Set No → master blanks (also hydrate from jobs separately)
    - Purchase Material on sheet is NOT written here (job finalize field).
    """
    rows = dedupe_sheet_rows(rows)
    phase_map = build_sku_phase_map() if migration_phase is not None else {}
    created = 0
    updated = 0
    skipped = 0
    phase_skipped = 0
    missing_sku = 0
    inferred_product_type = 0
    field_hits = {}
    samples = []

    for source in rows:
        sku = str(_sheet_row_get(source, 'SKU') or '').strip()
        if not sku:
            continue
        if not sku_eligible_for_migration_phase(sku, migration_phase, phase_map=phase_map):
            phase_skipped += 1
            continue
        values = row_to_field_values(source)
        job_name = values.get('job_name') or str(_sheet_row_get(source, 'JOB NAME') or '').strip()
        if job_name:
            values['job_name'] = job_name

        recipe = SkuRecipe.objects.filter(sku__iexact=sku).first()
        recipe_changed = []
        if not recipe:
            if not create_missing:
                missing_sku += 1
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
            if (
                infer_product_type
                and migration_phase in (1, 2)
                and _field_is_blank(recipe.product_type)
            ):
                inferred = infer_product_type_from_sku(sku)
                if inferred:
                    recipe.product_type = inferred
                    recipe_changed = list(recipe_changed) + ['product_type']
                    inferred_product_type += 1
            if recipe_changed:
                recipe.save(update_fields=list(dict.fromkeys(recipe_changed + ['updated_at'])))
                updated += 1
                for name in recipe_changed:
                    field_hits[name] = field_hits.get(name, 0) + 1
                if len(samples) < 20:
                    samples.append((recipe.pk, sku, recipe.master_data_status, recipe_changed))
            else:
                skipped += 1

    if infer_product_type and migration_phase in (1, 2):
        from planning.models import SkuRecipe as SkuRecipeModel

        for recipe in SkuRecipeModel.objects.filter(product_type='').iterator(chunk_size=500):
            if get_sku_migration_phase(recipe.sku, phase_map=phase_map) != migration_phase:
                continue
            inferred = infer_product_type_from_sku(recipe.sku)
            if not inferred:
                continue
            recipe.product_type = inferred
            recipe.save(update_fields=['product_type', 'updated_at'])
            inferred_product_type += 1
            updated += 1
            field_hits['product_type'] = field_hits.get('product_type', 0) + 1
            if len(samples) < 20:
                samples.append((recipe.pk, recipe.sku, recipe.master_data_status, ['product_type (inferred)']))

    return {
        'created': created,
        'updated': updated,
        'skipped': skipped,
        'phase_skipped': phase_skipped,
        'missing_sku': missing_sku,
        'inferred_product_type': inferred_product_type,
        'field_hits': field_hits,
        'samples': samples,
        'migration_phase': migration_phase,
    }
