"""Apply user update workbook and optionally remove Phase 1 orphan SKUs."""

from __future__ import annotations

import importlib.util
from pathlib import Path

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from core.models import JobCard, ProductType
from planning.models import PlanningJob, SkuRecipe
from planning.sku_migration_phases import build_sku_phase_map, get_sku_migration_phase

ROOT = Path(__file__).resolve().parents[3]


def _load_review_module():
    spec = importlib.util.spec_from_file_location(
        'review_master_data_update',
        ROOT / 'scripts' / 'review_master_data_update.py',
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _resolve_update_path(path: str | None) -> Path:
    if path:
        candidate = Path(path)
        if not candidate.is_absolute():
            candidate = ROOT / candidate
        if not candidate.exists():
            raise CommandError(f'Update file not found: {candidate}')
        return candidate

    for name in (
        'all_phases_missing_master_data_update_clean.xlsx',
        'all_phases_missing_master_data_update.xlsx',
    ):
        candidate = ROOT / name
        if candidate.exists():
            return candidate
    raise CommandError('No update workbook found in project root.')


def _collect_phase1_orphan_ids(phase_map, keep_skus: set[str]):
    orphan_ids = []
    for recipe in SkuRecipe.objects.all().only('id', 'sku', 'job_process_type', 'print_passes'):
        if get_sku_migration_phase(recipe.sku, phase_map) != 1:
            continue
        job_process = recipe.job_process_type or 'print_and_pack'
        if job_process != 'print_and_pack' or recipe.print_passes is not None:
            continue
        if recipe.sku.casefold() in keep_skus:
            continue
        if PlanningJob.objects.filter(sku__iexact=recipe.sku).exists():
            continue
        if JobCard.objects.filter(SKU__iexact=recipe.sku).exists():
            continue
        orphan_ids.append(recipe.id)
    return orphan_ids


class Command(BaseCommand):
    help = (
        'Apply all_phases_missing_master_data_update.xlsx to SkuRecipe master fields '
        'and optionally delete Phase 1 orphan SKUs with missing print_passes.'
    )

    def add_arguments(self, parser):
        parser.add_argument(
            'path',
            nargs='?',
            help='Path to update workbook (defaults to clean/original file in project root).',
        )
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Report actions without saving.',
        )
        parser.add_argument(
            '--skip-delete-orphans',
            action='store_true',
            help='Do not delete Phase 1 orphan SKUs missing print_passes.',
        )
        parser.add_argument(
            '--skip-apply',
            action='store_true',
            help='Only run orphan deletion.',
        )

    def handle(self, *args, **options):
        update_path = _resolve_update_path(options.get('path'))
        dry_run = options['dry_run']
        review = _load_review_module()
        review.PATH = update_path

        valid_product_types = set(ProductType.objects.values_list('name', flat=True))
        _, rows = review.load_rows()
        if not rows:
            raise CommandError('No SKU rows found in update workbook.')

        errors = []
        for row in rows:
            sku = row['sku']
            product_type = row['resolved_product_type']
            job_process = row['resolved_job_process']
            print_passes = review.parse_passes(row['resolved_print_passes'])

            if not SkuRecipe.objects.filter(sku__iexact=sku).exists():
                errors.append(f'{sku}: not in database')
                continue
            if not product_type or product_type not in valid_product_types:
                errors.append(f'{sku}: invalid product_type {product_type!r}')
                continue
            if job_process not in {'print_and_pack', 'cut_and_pack'}:
                errors.append(f'{sku}: invalid job_process {job_process!r}')
                continue
            if job_process == 'print_and_pack':
                if print_passes not in review.VALID_PRINT_PASSES:
                    errors.append(f'{sku}: invalid print_passes {row.get("resolved_print_passes")!r}')
            elif print_passes not in (None, ''):
                errors.append(f'{sku}: cut_and_pack cannot have print_passes')

        if errors:
            raise CommandError('Validation failed:\n' + '\n'.join(errors[:40]))

        keep_skus = {row['sku'].casefold() for row in rows}
        phase_map = build_sku_phase_map()
        orphan_ids = [] if options['skip_delete_orphans'] else _collect_phase1_orphan_ids(phase_map, keep_skus)

        def _run():
            deleted = 0
            updated = 0
            unchanged = 0
            field_hits = {'product_type': 0, 'job_process_type': 0, 'print_passes': 0}

            if orphan_ids and not options['skip_delete_orphans']:
                delete_result = SkuRecipe.objects.filter(id__in=orphan_ids).delete()
                deleted = delete_result[0] if isinstance(delete_result, tuple) else delete_result

            if not options['skip_apply']:
                for row in rows:
                    recipe = SkuRecipe.objects.get(sku__iexact=row['sku'])
                    product_type = row['resolved_product_type']
                    job_process = row['resolved_job_process']
                    print_passes = review.parse_passes(row['resolved_print_passes'])
                    new_passes = print_passes if job_process == 'print_and_pack' else None

                    changed_fields = []
                    if (recipe.product_type or '').strip() != product_type:
                        recipe.product_type = product_type
                        changed_fields.append('product_type')
                        field_hits['product_type'] += 1
                    if (recipe.job_process_type or 'print_and_pack') != job_process:
                        recipe.job_process_type = job_process
                        changed_fields.append('job_process_type')
                        field_hits['job_process_type'] += 1
                    if recipe.print_passes != new_passes:
                        recipe.print_passes = new_passes
                        changed_fields.append('print_passes')
                        field_hits['print_passes'] += 1

                    if changed_fields:
                        recipe.save(update_fields=changed_fields + ['updated_at'])
                        updated += 1
                    else:
                        unchanged += 1

            return {
                'deleted': deleted,
                'updated': updated,
                'unchanged': unchanged,
                'field_hits': field_hits,
                'rows': len(rows),
            }

        if dry_run:
            with transaction.atomic():
                result = _run()
                transaction.set_rollback(True)
        else:
            result = _run()

        mode = 'DRY-RUN' if dry_run else 'APPLIED'
        self.stdout.write(self.style.SUCCESS(
            f'[{mode}] source={update_path.name} rows={result["rows"]} '
            f'deleted_orphans={result["deleted"]} updated={result["updated"]} '
            f'unchanged={result["unchanged"]}'
        ))
        if result['field_hits']:
            self.stdout.write('Fields changed:')
            for name, count in sorted(result['field_hits'].items(), key=lambda item: (-item[1], item[0])):
                self.stdout.write(f'  {name}: {count}')
