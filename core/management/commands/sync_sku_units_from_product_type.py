from django.core.management.base import BaseCommand
from django.db import transaction

from core.models import ProductType
from planning.models import PlanningJob, SkuRecipe


class Command(BaseCommand):
    help = (
        "Backfill unit_type/pcs_per_unit on existing SkuRecipe rows from their "
        "Product Type's category defaults (e.g. 'A4 Rim' -> rim/500), for SKUs "
        "created before that Product Type had its defaults configured — the "
        "auto-fill on SkuRecipe.save() only applies going forward, so a SKU "
        "master that predates a category's default_unit_type stays stale until "
        "backfilled here. Then re-saves any still-editable (draft/pending_qc/"
        "qc_rejected) PlanningJob using an updated SKU, so its order_qty_pcs/ "
        "required-sheets math and packing/dispatch validation pick up the "
        "correction immediately. Jobs already past that stage are reported, "
        "not touched — their units are frozen once in-flight, by design; "
        "review those by hand. Safe to re-run any time a Product Type's "
        "defaults are added or changed."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run', action='store_true',
            help='Report what would change without saving anything.',
        )

    def handle(self, *args, **options):
        dry_run = options['dry_run']
        categorized_types = ProductType.objects.exclude(default_unit_type='pcs')
        if not categorized_types.exists():
            self.stdout.write('No Product Types have a non-pcs default_unit_type configured. Nothing to do.')
            return

        recipes_updated = 0
        jobs_resynced = 0
        frozen_jobs_flagged = []

        for product_type in categorized_types:
            stale_recipes = SkuRecipe.objects.filter(
                product_type__iexact=product_type.name, unit_type='pcs', pcs_per_unit=1,
            )
            for recipe in stale_recipes:
                self.stdout.write(
                    f'SkuRecipe {recipe.sku!r} (product_type={product_type.name!r}): '
                    f'pcs/1 -> {product_type.default_unit_type}/{product_type.default_pcs_per_unit or "unset"}'
                )
                if not dry_run:
                    with transaction.atomic():
                        recipe.unit_type = product_type.default_unit_type
                        if product_type.default_pcs_per_unit:
                            recipe.pcs_per_unit = product_type.default_pcs_per_unit
                        recipe.save(update_fields=['unit_type', 'pcs_per_unit'])
                recipes_updated += 1

                jobs = PlanningJob.objects.filter(sku__iexact=recipe.sku)
                for job in jobs:
                    if job._print_passes_frozen():
                        frozen_jobs_flagged.append(job)
                        continue
                    if dry_run:
                        self.stdout.write(f'  would re-sync PlanningJob {job.jc_number} (status={job.status})')
                        continue
                    if job.sync_unit_type_from_sku_master():
                        job.save(update_fields=['unit_type', 'pcs_per_unit'])
                        self.stdout.write(self.style.SUCCESS(
                            f'  re-synced PlanningJob {job.jc_number}: unit_type={job.unit_type}, '
                            f'pcs_per_unit={job.pcs_per_unit}'
                        ))
                        jobs_resynced += 1

        self.stdout.write('')
        if dry_run:
            self.stdout.write(self.style.WARNING(f'DRY RUN: {recipes_updated} SkuRecipe row(s) would be updated.'))
        else:
            self.stdout.write(self.style.SUCCESS(
                f'Updated {recipes_updated} SkuRecipe row(s); re-synced {jobs_resynced} still-editable PlanningJob(s).'
            ))
        if frozen_jobs_flagged:
            self.stdout.write(self.style.WARNING(
                f'\n{len(frozen_jobs_flagged)} job(s) are past QC approval and were NOT touched '
                f'(units are frozen once in-flight) — review by hand if these need correcting:'
            ))
            for job in frozen_jobs_flagged:
                self.stdout.write(f'  {job.jc_number}  sku={job.sku!r}  status={job.status}')
