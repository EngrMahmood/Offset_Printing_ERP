"""One-time / on-demand backfill: fill blank SkuRecipe fields from planning jobs."""

from django.core.management.base import BaseCommand

from planning.models import SkuRecipe
from planning.services import (
    SKU_RECIPE_DESIGNER_FIELDS,
    SKU_RECIPE_PLANNER_FIELDS,
    _field_is_blank,
    build_sku_recipe_initial_from_recipe,
    hydrate_sku_recipe_from_planning_jobs,
)


class Command(BaseCommand):
    help = (
        'Fill blank SKU master fields from planning jobs for the same SKU. '
        'Does not overwrite non-blank recipe values.'
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Report what would change without saving.',
        )
        parser.add_argument(
            '--status',
            action='append',
            dest='statuses',
            choices=['draft', 'pending_review', 'reviewed', 'approved'],
            help='Limit to recipe status (repeatable). Default: all active recipes.',
        )
        parser.add_argument(
            '--sku',
            action='append',
            dest='skus',
            help='Limit to specific SKU(s) (repeatable).',
        )

    def handle(self, *args, **options):
        dry_run = options['dry_run']
        statuses = options.get('statuses') or None
        skus = options.get('skus') or None

        qs = SkuRecipe.objects.filter(is_active=True).order_by('sku', 'id')
        if statuses:
            qs = qs.filter(master_data_status__in=statuses)
        if skus:
            from django.db.models import Q
            q = Q()
            for sku in skus:
                q |= Q(sku__iexact=sku.strip())
            qs = qs.filter(q)

        updated = 0
        skipped = 0
        field_hits = {}
        samples = []

        for recipe in qs.iterator():
            initial = build_sku_recipe_initial_from_recipe(recipe)
            would_fill = []
            for field_name in SKU_RECIPE_DESIGNER_FIELDS + SKU_RECIPE_PLANNER_FIELDS:
                if field_name in {'notes', 'remarks'}:
                    continue
                current = getattr(recipe, field_name, None)
                if not _field_is_blank(current):
                    continue
                value = initial.get(field_name)
                if _field_is_blank(value):
                    continue
                would_fill.append(field_name)
                field_hits[field_name] = field_hits.get(field_name, 0) + 1

            if not would_fill:
                skipped += 1
                continue

            if dry_run:
                updated += 1
                if len(samples) < 25:
                    samples.append((recipe.pk, recipe.sku, recipe.master_data_status, would_fill))
                continue

            if hydrate_sku_recipe_from_planning_jobs(recipe):
                updated += 1
                if len(samples) < 25:
                    samples.append((recipe.pk, recipe.sku, recipe.master_data_status, would_fill))
            else:
                skipped += 1

        mode = 'DRY-RUN' if dry_run else 'APPLIED'
        self.stdout.write(self.style.SUCCESS(
            f'[{mode}] recipes updated: {updated}, unchanged: {skipped}, scanned: {qs.count()}'
        ))
        if field_hits:
            self.stdout.write('Fields filled (counts):')
            for name, count in sorted(field_hits.items(), key=lambda item: (-item[1], item[0])):
                self.stdout.write(f'  {name}: {count}')
        if samples:
            self.stdout.write('Samples:')
            for pk, sku, status, fields in samples:
                self.stdout.write(f'  #{pk} [{status}] {sku}: {", ".join(fields)}')
