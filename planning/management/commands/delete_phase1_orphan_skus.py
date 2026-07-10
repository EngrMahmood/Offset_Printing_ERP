"""Delete Phase 1 orphan SKU master records that were never approved."""

from __future__ import annotations

from django.core.management.base import BaseCommand
from django.db import transaction

from core.models import JobCard
from planning.models import PlanningJob, SkuRecipe
from planning.sku_migration_phases import build_sku_phase_map, get_sku_migration_phase


class Command(BaseCommand):
    help = 'Delete unapproved Phase 1 SKU recipes with no planning jobs or job cards.'

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Report SKUs that would be deleted without saving.',
        )

    def handle(self, *args, **options):
        phase_map = build_sku_phase_map()
        orphan_ids = []
        samples = []

        for recipe in SkuRecipe.objects.filter(is_active=True).exclude(master_data_status='approved'):
            if get_sku_migration_phase(recipe.sku, phase_map) != 1:
                continue
            if PlanningJob.objects.filter(sku__iexact=recipe.sku).exists():
                continue
            if JobCard.objects.filter(SKU__iexact=recipe.sku).exists():
                continue
            orphan_ids.append(recipe.id)
            if len(samples) < 10:
                samples.append(recipe.sku)

        def _run():
            if not orphan_ids:
                return 0
            result = SkuRecipe.objects.filter(id__in=orphan_ids).delete()
            return result[0] if isinstance(result, tuple) else result

        if options['dry_run']:
            with transaction.atomic():
                deleted = _run()
                transaction.set_rollback(True)
            mode = 'DRY-RUN'
        else:
            deleted = _run()
            mode = 'APPLIED'

        self.stdout.write(self.style.SUCCESS(
            f'[{mode}] phase1_orphan_candidates={len(orphan_ids)} deleted={deleted if not options["dry_run"] else len(orphan_ids)}'
        ))
        if samples:
            self.stdout.write('Samples: ' + ', '.join(samples))
