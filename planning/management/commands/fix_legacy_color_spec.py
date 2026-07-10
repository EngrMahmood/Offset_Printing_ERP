"""Repair mangled legacy print colors on SKU masters and planning jobs (phases 1–3)."""

from __future__ import annotations

from django.core.management.base import BaseCommand
from django.db import transaction
from django.utils import timezone

from core.print_colors import (
    apply_print_color_to_planning_job,
    normalize_color_spec_value,
    repair_mangled_decimal_color_spec,
)
from planning.models import PlanningJob, SkuRecipe
from planning.sku_migration_phases import build_sku_phase_map, get_sku_migration_phase


def _sku_in_scope(sku, *, phase, phase_map):
    if phase in (None, 'all'):
        return True
    try:
        phase_int = int(phase)
    except (TypeError, ValueError):
        return False
    from planning.sku_migration_phases import get_sku_migration_phase
    return get_sku_migration_phase(sku, phase_map=phase_map) == phase_int

def _target_color_spec(current):
    current = (current or '').strip()
    if not current:
        return current, False
    repaired = repair_mangled_decimal_color_spec(current)
    normalized = normalize_color_spec_value(repaired if repaired != current else current)
    if normalized == current:
        return current, False
    return normalized, True


class Command(BaseCommand):
    help = (
        'Fix legacy sheet color_spec mangling (40→4+0, 4.0→4+0, etc.) on approved legacy SKU masters '
        'and their planning jobs for migration phases 1, 2, and/or 3.'
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--phase',
            choices=('1', '2', '3', 'all'),
            default='all',
            help='Migration phase to repair (default: all phases).',
        )
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Report changes without saving.',
        )
        parser.add_argument(
            '--include-non-legacy',
            action='store_true',
            help='Also repair non-legacy_produced SKU masters in scope.',
        )

    def handle(self, *args, **options):
        phase = options['phase']
        dry_run = options['dry_run']
        phase_map = build_sku_phase_map()

        recipe_qs = SkuRecipe.objects.filter(is_active=True).exclude(color_spec='')
        if not options['include_non_legacy']:
            recipe_qs = recipe_qs.filter(legacy_produced=True)

        recipe_updates = []
        for recipe in recipe_qs.iterator(chunk_size=500):
            if not _sku_in_scope(recipe.sku, phase=phase, phase_map=phase_map):
                continue
            new_color, changed = _target_color_spec(recipe.color_spec)
            if changed:
                recipe_updates.append((recipe, new_color))

        job_updates = []
        job_qs = PlanningJob.objects.filter(is_active=True).exclude(color_spec='')
        for job in job_qs.iterator(chunk_size=500):
            if not _sku_in_scope(job.sku, phase=phase, phase_map=phase_map):
                continue
            new_color, changed = _target_color_spec(job.color_spec)
            if changed:
                job_updates.append((job, new_color))

        mode = 'DRY-RUN' if dry_run else 'APPLIED'
        self.stdout.write(
            f'[{mode}] phase={phase} recipes_to_fix={len(recipe_updates)} jobs_to_fix={len(job_updates)}'
        )

        for recipe, new_color in recipe_updates[:30]:
            self.stdout.write(f'  recipe {recipe.pk} {recipe.sku}: {recipe.color_spec!r} -> {new_color!r}')
        if len(recipe_updates) > 30:
            self.stdout.write(f'  ... and {len(recipe_updates) - 30} more recipes')

        for job, new_color in job_updates[:30]:
            self.stdout.write(f'  job {job.pk} {job.sku}: {job.color_spec!r} -> {new_color!r}')
        if len(job_updates) > 30:
            self.stdout.write(f'  ... and {len(job_updates) - 30} more jobs')

        if dry_run:
            return

        now = timezone.now()
        with transaction.atomic():
            for recipe, new_color in recipe_updates:
                recipe.color_spec = new_color
                recipe.save(update_fields=['color_spec', 'updated_at'])

            for job, new_color in job_updates:
                job.color_spec = new_color
                apply_print_color_to_planning_job(job)
                job.updated_at = now
                job.save(update_fields=['color_spec', 'total_colors', 'updated_at'])

        self.stdout.write(self.style.SUCCESS(f'[{mode}] Done.'))
