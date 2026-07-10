"""Audit SKU master migration phases against the current database."""

from __future__ import annotations

import os
import sys
from collections import Counter
from pathlib import Path

import django

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from planning.models import PlanningJob, SkuRecipe
from planning.sku_migration_phases import build_sku_phase_map, get_sku_migration_phase


def main():
    phase_map = build_sku_phase_map()
    phase_counts = Counter()
    missing_product_type = Counter()
    missing_print_passes = Counter()

    for recipe in SkuRecipe.objects.all().iterator(chunk_size=500):
        phase = get_sku_migration_phase(recipe.sku, phase_map=phase_map)
        phase_counts[phase] += 1
        if not (recipe.product_type or '').strip():
            missing_product_type[phase] += 1
        if recipe.job_process_type == 'print_and_pack' and recipe.print_passes is None:
            missing_print_passes[phase] += 1

    job_count = PlanningJob.objects.count()
    print('=== SKU migration phase audit ===')
    print(f'PlanningJob records: {job_count}')
    print(f'SkuRecipe records: {SkuRecipe.objects.count()}')
    print()
    for phase in (1, 2, 3):
        label = {1: 'Never in planning/production', 2: 'Planning only (not released)', 3: 'Released to production'}[phase]
        print(f'Phase {phase} — {label}')
        print(f'  SKUs: {phase_counts.get(phase, 0)}')
        print(f'  Missing product_type: {missing_product_type.get(phase, 0)}')
        print(f'  Missing print_passes (print_and_pack): {missing_print_passes.get(phase, 0)}')
        print()

    print('Phase 1 is safe for full sheet fill-blanks + prefix product_type inference.')
    print('Phase 2/3 should be handled in later migration steps with stricter guards.')


if __name__ == '__main__':
    main()
