"""Pre-migration audit: legacy Google Sheet (.xlsb) vs SkuRecipe database."""
from __future__ import annotations

import os
import sys
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

import django

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from planning.models import PlanningJob, SkuRecipe
from planning.sku_sheet_import import (
    RESTORE_FIELDS,
    SHEET_HEADER_TO_FIELD,
    row_to_field_values,
    row_to_job_origin,
)

SHEET_PATH = ROOT / 'Offset_Master Data_Planning_ERP 1.xlsb'

COMPARE_FIELDS = list(RESTORE_FIELDS) + ['job_process_type']


def _norm_header(value):
    return str(value or '').strip()


def _sheet_row_get(source, header):
    if header in source:
        return source.get(header)
    target = header.casefold()
    for key, value in source.items():
        if _norm_header(key).casefold() == target:
            return value
    return None


def parse_xlsb(path):
    from pyxlsb import open_workbook

    rows = []
    with open_workbook(path) as wb:
        sheet_name = wb.sheets[0]
        with wb.get_sheet(sheet_name) as ws:
            all_rows = list(ws.rows())
            header_row_idx = None
            header = []
            for i, row in enumerate(all_rows[:15]):
                values = [_norm_header(c.v if hasattr(c, 'v') else c) for c in row]
                if 'SKU' in values and any(v.upper() in {'JOB NAME', 'JOBNAME'} for v in values):
                    header_row_idx = i
                    header = values
                    break
            if header_row_idx is None:
                raise ValueError('Could not find header row (need SKU and JOB NAME).')
            for row in all_rows[header_row_idx + 1:]:
                values = [c.v if hasattr(c, 'v') else c for c in row]
                record = {}
                for idx, key in enumerate(header):
                    if key:
                        record[key] = values[idx] if idx < len(values) else None
                rows.append(record)
    return rows


def _ser(value):
    if value is None:
        return None
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, bool):
        return value
    text = str(value).strip()
    return text if text else None


def _blank(value):
    return value in (None, '')


def main():
    if not SHEET_PATH.exists():
        print(f'MISSING FILE: {SHEET_PATH}')
        return 1

    rows = parse_xlsb(SHEET_PATH)
    print(f'=== Pre-migration audit ===')
    print(f'Sheet: {SHEET_PATH.name}')
    print(f'Sheet rows (with data): {len(rows)}')

    # Headers in sheet
    headers = set()
    for row in rows:
        headers.update(row.keys())
    print(f'Sheet columns ({len(headers)}): {sorted(headers)}')

    mapped_headers = set(SHEET_HEADER_TO_FIELD.keys())
    unmapped = sorted(h for h in headers if h and h not in mapped_headers and h not in {
        'Sno.', 'Sno', 'Order Status', 'Size W Inch', 'Size H Inch', 'Purchase Material',
        'Purchase Material Origin', 'Purchase Origin',
    })
    if unmapped:
        print(f'WARNING unmapped sheet columns: {unmapped}')

    # SKU analysis
    sheet_skus = []
    duplicate_skus = Counter()
    empty_sku_rows = 0
    for row in rows:
        sku = str(_sheet_row_get(row, 'SKU') or '').strip()
        if not sku:
            empty_sku_rows += 1
            continue
        sheet_skus.append(sku)
        duplicate_skus[sku.casefold()] += 1

    dup_list = [sku for sku, count in duplicate_skus.items() if count > 1]
    print(f'\n--- SKU counts ---')
    print(f'Rows with SKU: {len(sheet_skus)}')
    print(f'Rows without SKU: {empty_sku_rows}')
    print(f'Unique SKUs (case-insensitive): {len(duplicate_skus)}')
    print(f'Duplicate SKUs in sheet: {len(dup_list)}')
    if dup_list[:10]:
        print('  Examples:', dup_list[:10])

    db_recipes = {r.sku.casefold(): r for r in SkuRecipe.objects.all()}
    db_skus = set(db_recipes.keys())
    sheet_sku_set = set(duplicate_skus.keys())

    in_sheet_not_db = sorted(sheet_sku_set - db_skus)
    in_db_not_sheet = sorted(db_skus - sheet_sku_set)
    in_both = sheet_sku_set & db_skus

    print(f'\n--- Sheet vs DB SKU overlap ---')
    print(f'In DB (SkuRecipe): {len(db_skus)}')
    print(f'In sheet only (would need CREATE): {len(in_sheet_not_db)}')
    print(f'In DB only (not in legacy sheet): {len(in_db_not_sheet)}')
    print(f'In both: {len(in_both)}')

    # Status breakdown for DB-only and both
    status_counts = Counter(r.master_data_status for r in db_recipes.values())
    print(f'\nDB master_data_status: {dict(status_counts)}')
    legacy_count = SkuRecipe.objects.filter(legacy_produced=True).count()
    print(f'DB legacy_produced=True: {legacy_count}')

    # Field-level comparison for overlapping SKUs
    conflicts = []
    would_fill = Counter()
    sheet_only_data = Counter()
    db_blank_sheet_has = []
    cancel_rows = []
    cut_and_pack_rows = 0

    for row in rows:
        sku_raw = str(_sheet_row_get(row, 'SKU') or '').strip()
        if not sku_raw:
            continue
        sku_key = sku_raw.casefold()
        order_status = str(_sheet_row_get(row, 'Order Status') or '').strip().lower()
        remarks = str(_sheet_row_get(row, 'Remarks') or _sheet_row_get(row, 'Notes') or '').strip().lower()
        if 'cancel' in order_status or 'cancel' in remarks or 'po cancel' in remarks:
            cancel_rows.append(sku_raw)

        values = row_to_field_values(row)
        if values.get('job_process_type') == 'cut_and_pack':
            cut_and_pack_rows += 1

        recipe = db_recipes.get(sku_key)
        if not recipe:
            if values:
                sheet_only_data['new_sku_with_data'] += 1
            continue

        for field in COMPARE_FIELDS:
            sheet_val = _ser(values.get(field))
            db_val = _ser(getattr(recipe, field, None))
            if _blank(db_val) and not _blank(sheet_val):
                would_fill[field] += 1
                db_blank_sheet_has.append((sku_raw, field, sheet_val))
            elif not _blank(db_val) and not _blank(sheet_val) and db_val != sheet_val:
                conflicts.append({
                    'sku': sku_raw,
                    'field': field,
                    'db': db_val,
                    'sheet': sheet_val,
                    'status': recipe.master_data_status,
                })

    print(f'\n--- Sheet content signals ---')
    print(f'Rows flagged cancel/cancelled: {len(cancel_rows)}')
    print(f'Cut & Pack rows (Color=No + Application=No): {cut_and_pack_rows}')

    print(f'\n--- Fill-blanks-only migration preview ---')
    print(f'Fields that WOULD be filled (DB blank, sheet has value):')
    for field, count in sorted(would_fill.items(), key=lambda x: (-x[1], x[0])):
        print(f'  {field}: {count}')

    print(f'\n--- Conflicts (DB has value, sheet differs) ---')
    print(f'Total conflicts: {len(conflicts)}')
    by_field = Counter(c['field'] for c in conflicts)
    for field, count in sorted(by_field.items(), key=lambda x: (-x[1], x[0])):
        print(f'  {field}: {count}')
    approved_conflicts = [c for c in conflicts if c['status'] == 'approved']
    print(f'Conflicts on APPROVED recipes: {len(approved_conflicts)}')

    if conflicts[:15]:
        print('\nSample conflicts (first 15):')
        for c in conflicts[:15]:
            print(f"  {c['sku']} | {c['field']} | DB={c['db']!r} | Sheet={c['sheet']!r} | status={c['status']}")

    # Planning jobs for sheet SKUs not in DB
    jobs_missing_recipe = PlanningJob.objects.filter(
        sku__isnull=False,
    ).exclude(sku='').values_list('sku', flat=True).distinct()
    job_skus_no_recipe = []
    for js in jobs_missing_recipe:
        if js and js.casefold() not in db_skus:
            job_skus_no_recipe.append(js)
    print(f'\n--- Planning jobs without SkuRecipe ---')
    print(f'Distinct job SKUs with no master recipe: {len(job_skus_no_recipe)}')
    if job_skus_no_recipe[:10]:
        print('  Examples:', job_skus_no_recipe[:10])

    # Purchase origin preview
    origin_would_fill = 0
    for row in rows:
        sku = str(_sheet_row_get(row, 'SKU') or '').strip()
        if not sku:
            continue
        origin = row_to_job_origin(row)
        if not origin:
            continue
        jobs = PlanningJob.objects.filter(sku__iexact=sku).exclude(
            purchase_material_origin__isnull=False,
        ).exclude(purchase_material_origin='')
        origin_would_fill += jobs.count()
    print(f'\nPlanning jobs that would get purchase_material_origin: {origin_would_fill}')

    # New SKUs sample
    print(f'\n--- New SKUs in sheet (not in DB) — first 20 ---')
    for sku_cf in in_sheet_not_db[:20]:
        # find original casing from sheet
        for row in rows:
            s = str(_sheet_row_get(row, 'SKU') or '').strip()
            if s.casefold() == sku_cf:
                job_name = str(_sheet_row_get(row, 'JOB NAME') or '').strip()
                mat = str(_sheet_row_get(row, 'Material') or '').strip()
                print(f'  {s} | {job_name[:50]} | material={mat}')
                break

    # DB-only sample
    print(f'\n--- DB SKUs not in sheet — first 15 ---')
    for sku_cf in in_db_not_sheet[:15]:
        r = db_recipes[sku_cf]
        print(f'  {r.sku} | status={r.master_data_status} | legacy={r.legacy_produced}')

    # Data quality issues in sheet
    print(f'\n--- Sheet data quality issues ---')
    missing_material = 0
    missing_job_name = 0
    zero_size = 0
    for row in rows:
        sku = str(_sheet_row_get(row, 'SKU') or '').strip()
        if not sku:
            continue
        if not str(_sheet_row_get(row, 'Material') or '').strip():
            missing_material += 1
        if not str(_sheet_row_get(row, 'JOB NAME') or '').strip():
            missing_job_name += 1
        w = _sheet_row_get(row, 'Size W mm')
        h = _sheet_row_get(row, 'Size H mm')
        try:
            if w is not None and h is not None and float(w) == 0 and float(h) == 0:
                zero_size += 1
        except (TypeError, ValueError):
            pass
    print(f'  Missing material: {missing_material}')
    print(f'  Missing job name: {missing_job_name}')
    print(f'  Zero size (0x0 mm): {zero_size}')

    # File format note
    print(f'\n--- Technical notes ---')
    print(f'  File format: .xlsb (binary) — current import command supports CSV/XLSX only.')
    print(f'  Must export to XLSX/CSV or extend parse_sheet_rows() for .xlsb.')
    print(f'  Default migration mode: fill_blanks_only (won\'t overwrite conflicts above).')
    print(f'  Use --create-missing to add {len(in_sheet_not_db)} new SKUs as draft recipes.')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
