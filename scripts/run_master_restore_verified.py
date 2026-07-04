"""
Snapshot → hydrate master from jobs → restore blanks from Master_Data sheet → verify.

Rules:
- SkuRecipe: fill blanks only (never overwrite non-blank, never demote approved).
- PlanningJob: only blank purchase_material_origin on NON-executed jobs.
- Executed jobs (released / in_production / completed / planning_done / has job card)
  must be byte-identical after the run.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import django

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from planning.models import PlanningJob, SkuRecipe
from planning.services import hydrate_sku_recipe_from_planning_jobs
from planning.sku_sheet_import import (
    parse_sheet_rows,
    planning_job_is_executed,
    restore_sku_recipes_from_rows,
)

SHEET_CANDIDATES = [
    ROOT / 'planning' / 'docs' / 'Master_Data_extracted.csv',
    ROOT / 'Planning 2026.xlsx',
]

RECIPE_FIELDS = [
    'sku', 'job_name', 'material', 'color_spec', 'application', 'product_type',
    'machine_name', 'plate_set_no', 'size_w_mm', 'size_h_mm', 'ups',
    'print_sheet_size', 'purchase_sheet_size', 'purchase_sheet_ups',
    'default_unit_cost', 'daily_demand', 'awc_no', 'die_cutting', 'notes',
    'remarks', 'job_process_type', 'master_data_status', 'is_active',
]

JOB_FIELDS = [
    'id', 'jc_number', 'sku', 'status', 'planning_stage', 'is_active',
    'job_name', 'material', 'color_spec', 'application', 'machine_name',
    'plate_set_no', 'size_w_mm', 'size_h_mm', 'ups', 'print_sheet_size',
    'purchase_sheet_size', 'purchase_sheet_ups', 'purchase_material_origin',
    'remarks', 'wastage_sheets', 'print_passes', 'destination',
]


def _ser(value):
    if value is None:
        return None
    if hasattr(value, 'isoformat'):
        return value.isoformat()
    return str(value) if not isinstance(value, (int, float, bool, str)) else value


def snapshot_recipes():
    data = {}
    for recipe in SkuRecipe.objects.all().iterator(chunk_size=500):
        data[recipe.pk] = {f: _ser(getattr(recipe, f, None)) for f in RECIPE_FIELDS}
    return data


def snapshot_jobs():
    data = {}
    for job in PlanningJob.objects.all().iterator(chunk_size=500):
        row = {f: _ser(getattr(job, f, None)) for f in JOB_FIELDS}
        row['executed'] = planning_job_is_executed(job)
        data[job.pk] = row
    return data


def compare_recipes(before, after):
    errors = []
    filled = {}
    for pk, old in before.items():
        new = after.get(pk)
        if not new:
            errors.append(f'recipe #{pk} missing after restore')
            continue
        for field, old_val in old.items():
            new_val = new.get(field)
            old_blank = old_val in (None, '')
            if not old_blank and old_val != new_val:
                errors.append(
                    f'recipe #{pk} {old.get("sku")}: non-blank {field} changed '
                    f'{old_val!r} -> {new_val!r}'
                )
            elif old_blank and new_val not in (None, '') and old_val != new_val:
                filled[field] = filled.get(field, 0) + 1
    for pk in after:
        if pk not in before:
            errors.append(f'recipe #{pk} created unexpectedly')
    return errors, filled


def compare_jobs(before, after):
    errors = []
    origin_filled = 0
    executed_count = 0
    for pk, old in before.items():
        new = after.get(pk)
        if not new:
            errors.append(f'job #{pk} missing after restore')
            continue
        executed = old.get('executed')
        if executed:
            executed_count += 1
            for field in JOB_FIELDS:
                if old.get(field) != new.get(field):
                    errors.append(
                        f'EXECUTED job #{pk} {old.get("jc_number")}: {field} changed '
                        f'{old.get(field)!r} -> {new.get(field)!r}'
                    )
            continue
        for field in JOB_FIELDS:
            if field == 'purchase_material_origin':
                old_val = old.get(field)
                new_val = new.get(field)
                if old_val not in (None, '') and old_val != new_val:
                    errors.append(
                        f'job #{pk} {old.get("jc_number")}: non-blank origin changed '
                        f'{old_val!r} -> {new_val!r}'
                    )
                elif old_val in (None, '') and new_val not in (None, ''):
                    origin_filled += 1
                continue
            if old.get(field) != new.get(field):
                errors.append(
                    f'job #{pk} {old.get("jc_number")}: field {field} changed '
                    f'{old.get(field)!r} -> {new.get(field)!r}'
                )
    return errors, origin_filled, executed_count


def main():
    sheet = next((p for p in SHEET_CANDIDATES if p.exists()), None)
    if not sheet:
        print('ERROR: Master_Data sheet/csv not found')
        return 1

    print(f'Using sheet: {sheet}')
    print('Snapshotting recipes and jobs...')
    before_recipes = snapshot_recipes()
    before_jobs = snapshot_jobs()
    print(f'  recipes={len(before_recipes)} jobs={len(before_jobs)} '
          f'executed_jobs={sum(1 for r in before_jobs.values() if r["executed"])}')

    print('Hydrating blank master fields from planning jobs...')
    hydrated = 0
    for recipe in SkuRecipe.objects.filter(is_active=True).iterator(chunk_size=200):
        if hydrate_sku_recipe_from_planning_jobs(recipe):
            hydrated += 1
    print(f'  hydrated recipes: {hydrated}')

    print('Restoring blanks from Google Sheet Master_Data...')
    rows = parse_sheet_rows(str(sheet))
    result = restore_sku_recipes_from_rows(rows, fill_blanks_only=True, create_missing=False)
    print(
        f"  recipes_updated={result['updated']} created={result['created']} "
        f"skipped={result['skipped']} missing_sku={result['missing_sku']} "
        f"jobs_origin_updated={result.get('jobs_origin_updated', 0)}"
    )
    if result['field_hits']:
        print('  field hits:')
        for name, count in sorted(result['field_hits'].items(), key=lambda i: (-i[1], i[0])):
            print(f'    {name}: {count}')

    print('Re-snapshot and verify...')
    after_recipes = snapshot_recipes()
    after_jobs = snapshot_jobs()

    recipe_errors, recipe_filled = compare_recipes(before_recipes, after_recipes)
    job_errors, origin_filled, executed_count = compare_jobs(before_jobs, after_jobs)

    print()
    print('=== VERIFICATION ===')
    print(f'Executed jobs protected: {executed_count}')
    print(f'Master blank fields filled (by field): {recipe_filled}')
    print(f'Non-executed jobs purchase origin filled: {origin_filled}')
    print(f'Recipe integrity errors: {len(recipe_errors)}')
    print(f'Job integrity errors: {len(job_errors)}')

    report_path = ROOT / 'planning' / 'docs' / 'master_restore_verification.json'
    report = {
        'sheet': str(sheet),
        'hydrated_recipes': hydrated,
        'restore_result': {
            k: v for k, v in result.items() if k != 'samples'
        },
        'restore_samples': result.get('samples', []),
        'recipe_filled': recipe_filled,
        'jobs_origin_filled': origin_filled,
        'executed_jobs': executed_count,
        'recipe_errors': recipe_errors[:50],
        'job_errors': job_errors[:50],
        'ok': not recipe_errors and not job_errors,
    }
    report_path.write_text(json.dumps(report, indent=2), encoding='utf-8')
    print(f'Report: {report_path}')

    if recipe_errors:
        print('RECIPE ERRORS (sample):')
        for line in recipe_errors[:20]:
            print(' ', line)
    if job_errors:
        print('JOB ERRORS (sample):')
        for line in job_errors[:20]:
            print(' ', line)

    if recipe_errors or job_errors:
        print('FAILED verification')
        return 2

    # Extra: sheet coverage for AWC blanks remaining
    from planning.sku_sheet_import import _field_is_blank, _sheet_row_get, row_to_field_values

    sheet_awc_by_sku = {}
    for source in rows:
        sku = str(_sheet_row_get(source, 'SKU') or '').strip()
        if not sku:
            continue
        vals = row_to_field_values(source)
        if vals.get('awc_no'):
            sheet_awc_by_sku[sku.lower()] = vals['awc_no']

    still_blank_awc = 0
    blank_but_sheet_has = 0
    for recipe in SkuRecipe.objects.filter(is_active=True).iterator():
        if not _field_is_blank(recipe.awc_no):
            continue
        still_blank_awc += 1
        if recipe.sku.strip().lower() in sheet_awc_by_sku:
            blank_but_sheet_has += 1

    print()
    print(f'Active recipes still blank AWC: {still_blank_awc}')
    print(f'  of which sheet still has AWC (should be 0): {blank_but_sheet_has}')
    if blank_but_sheet_has:
        print('WARNING: some AWC blanks remain despite sheet values')
        return 3

    print('PASSED: executed jobs unchanged; master/job blanks filled safely')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
