"""Audit blank fields on SKU Recipe Master list and restore sources."""
import os
import sys
from collections import Counter

import django

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from django.db.models import Q

from core.models import JobCard
from planning.models import PlanningJob, SkuRecipe
from planning.services import _field_is_blank, build_sku_recipe_initial_from_recipe

LIST_FIELDS = [
    'job_name',
    'material',
    'color_spec',
    'application',
    'size_w_mm',
    'size_h_mm',
    'ups',
    'print_sheet_size',
    'purchase_sheet_size',
    'default_unit_cost',
]

EXTRA_FIELDS = [
    'awc_no',
    'die_cutting',
    'machine_name',
    'plate_set_no',
    'daily_demand',
    'purchase_sheet_ups',
    'product_type',
    'remarks',
    'notes',
]

def is_blank(val):
    return _field_is_blank(val)


def blank_count(qs, field):
    if field in {'size_w_mm', 'size_h_mm', 'ups', 'default_unit_cost', 'daily_demand', 'purchase_sheet_ups'}:
        return qs.filter(**{f'{field}__isnull': True}).count()
    return qs.filter(Q(**{field: ''}) | Q(**{f'{field}__isnull': True})).count()


def jobcard_value(card, field):
    if not card:
        return None
    if field == 'material':
        mat = getattr(card, 'material', None)
        if mat is None:
            return None
        return getattr(mat, 'name', None) or str(mat)
    if field == 'color_spec':
        return getattr(card, 'colour', None)
    if field == 'application':
        return getattr(card, 'application', None)
    if field == 'ups':
        return getattr(card, 'ups', None)
    if field == 'print_sheet_size':
        return getattr(card, 'print_sheet_size', None)
    if field == 'purchase_sheet_size':
        return getattr(card, 'purchase_sheet_size', None)
    return None


def main():
    qs = SkuRecipe.objects.filter(is_active=True)
    total = qs.count()
    print(f'Active recipes: {total}')
    print()
    print('=== Blank rates (list columns) ===')
    for field in LIST_FIELDS:
        n = blank_count(qs, field)
        print(f'  {field}: {n} blank ({100 * n / total:.1f}%)')

    print()
    print('=== Blank rates (extra master fields) ===')
    for field in EXTRA_FIELDS:
        n = blank_count(qs, field)
        print(f'  {field}: {n} blank ({100 * n / total:.1f}%)')

    for status in ['approved', 'draft', 'pending_review', 'reviewed']:
        status_qs = qs.filter(master_data_status=status)
        count = status_qs.count()
        if not count:
            continue
        print()
        print(f'=== Status {status} ({count}) list blanks ===')
        for field in LIST_FIELDS:
            n = blank_count(status_qs, field)
            if n:
                print(f'  {field}: {n} ({100 * n / count:.1f}%)')

    # Preload latest job card per SKU for fields not on planning job initial
    print()
    print('Scanning recipes for restore sources...')
    restorable_job = Counter()
    restorable_card = Counter()
    neither = Counter()
    missing_any = 0
    fully_ok = 0
    samples_job = []
    samples_card_only = []
    samples_neither = []

    # Build jobcard map: sku lower -> best card values
    card_by_sku = {}
    for card in JobCard.objects.select_related('material').only(
        'id', 'SKU', 'material', 'colour', 'application',
        'ups', 'print_sheet_size', 'purchase_sheet_size',
    ).iterator(chunk_size=500):
        sku = (card.SKU or '').strip().lower()
        if not sku:
            continue
        existing = card_by_sku.get(sku)
        if not existing or (card.id or 0) > (existing.id or 0):
            card_by_sku[sku] = card

    for recipe in qs.iterator(chunk_size=200):
        blanks = [f for f in LIST_FIELDS if is_blank(getattr(recipe, f))]
        if not blanks:
            fully_ok += 1
            continue
        missing_any += 1
        initial = build_sku_recipe_initial_from_recipe(recipe)
        card = card_by_sku.get((recipe.sku or '').strip().lower())
        filled_job = []
        filled_card = []
        unfilled = []
        for field in blanks:
            job_val = initial.get(field)
            if not is_blank(job_val):
                restorable_job[field] += 1
                filled_job.append(field)
                continue
            card_val = jobcard_value(card, field) if card else None
            if not is_blank(card_val):
                restorable_card[field] += 1
                filled_card.append(field)
                continue
            neither[field] += 1
            unfilled.append(field)

        if filled_job and len(samples_job) < 12:
            samples_job.append((recipe.pk, recipe.sku, recipe.master_data_status, blanks, filled_job))
        if filled_card and not filled_job and len(samples_card_only) < 12:
            samples_card_only.append((recipe.pk, recipe.sku, recipe.master_data_status, blanks, filled_card))
        if unfilled and not filled_job and not filled_card and len(samples_neither) < 12:
            samples_neither.append((recipe.pk, recipe.sku, recipe.master_data_status, unfilled))

    print()
    print(f'Recipes with all list fields present: {fully_ok}')
    print(f'Recipes missing at least one list field: {missing_any}')
    print()
    print('Still blank after hydrate (need job card or sheet):')
    # Recompute: blanks not fillable from jobs
    print('Restorable from planning jobs (still blank on recipe):')
    for field, count in restorable_job.most_common():
        print(f'  {field}: {count}')
    print('Not on jobs, but on job cards:')
    for field, count in restorable_card.most_common():
        print(f'  {field}: {count}')
    print('Not on jobs or job cards (need Google Sheet / manual):')
    for field, count in neither.most_common():
        print(f'  {field}: {count}')

    print()
    print('Samples restorable from jobs:')
    for row in samples_job:
        print(' ', row)
    print('Samples restorable only from job cards:')
    for row in samples_card_only:
        print(' ', row)
    print('Samples with no internal source:')
    for row in samples_neither:
        print(' ', row)

    # Color values that look like plate inks (pollution)
    ink_like = 0
    ink_samples = []
    from printing_plates.constants import is_plate_ink_spec
    for recipe in qs.exclude(color_spec='').only('sku', 'color_spec').iterator():
        if is_plate_ink_spec(recipe.color_spec):
            ink_like += 1
            if len(ink_samples) < 10:
                ink_samples.append((recipe.sku, recipe.color_spec))
    print()
    print(f'color_spec looks like plate inks (not print color): {ink_like}')
    for row in ink_samples:
        print(' ', row)

    # Application values not in dropdown choices
    from planning.forms import APPLICATION_CHOICES
    allowed = {c[0] for c in APPLICATION_CHOICES if c[0]}
    odd_app = Counter()
    for recipe in qs.exclude(application='').only('application').iterator():
        if recipe.application not in allowed:
            odd_app[recipe.application] += 1
    print()
    print(f'application values not in dropdown ({sum(odd_app.values())} recipes):')
    for val, count in odd_app.most_common(15):
        print(f'  {val!r}: {count}')


if __name__ == '__main__':
    main()
