"""Restore blank SKU master + job purchase origin from a Google Sheet export."""

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from planning.sku_sheet_import import parse_sheet_rows, restore_sku_recipes_from_rows


class Command(BaseCommand):
    help = (
        'Fill blank SkuRecipe master fields from a Google Sheet CSV/XLSX/XLSB export. '
        'Does not overwrite non-blank values, does not touch planning-job finalize fields, '
        'and does not demote approved recipes.'
    )

    def add_arguments(self, parser):
        parser.add_argument('path', help='Path to CSV, XLSX, or XLSB export from Google Sheet')
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Report what would change without saving.',
        )
        parser.add_argument(
            '--create-missing',
            action='store_true',
            help='Create draft recipes for SKUs in the sheet that are not in the DB.',
        )
        parser.add_argument(
            '--overwrite',
            action='store_true',
            help='Overwrite non-blank recipe/job fields with sheet values (dangerous).',
        )
        parser.add_argument(
            '--phase',
            type=int,
            choices=[1, 2, 3],
            help='Only update SKUs in this migration phase (1=never planned, 2=planning only, 3=production).',
        )
        parser.add_argument(
            '--infer-product-type',
            action='store_true',
            help='Infer blank product_type from SKU prefix for ERP-only SKUs (phase 1 or 2).',
        )

    def handle(self, *args, **options):
        path = options['path']
        dry_run = options['dry_run']
        phase = options.get('phase')
        infer_product_type = options.get('infer_product_type')
        if infer_product_type and phase not in (None, 1, 2):
            raise CommandError('--infer-product-type is only valid with --phase 1 or --phase 2.')
        try:
            rows = parse_sheet_rows(path)
        except Exception as exc:
            raise CommandError(str(exc)) from exc

        if not rows:
            raise CommandError('No rows found in file.')

        def _run():
            return restore_sku_recipes_from_rows(
                rows,
                fill_blanks_only=not options['overwrite'],
                create_missing=options['create_missing'],
                migration_phase=phase,
                infer_product_type=infer_product_type,
            )

        if dry_run:
            with transaction.atomic():
                result = _run()
                transaction.set_rollback(True)
        else:
            result = _run()

        mode = 'DRY-RUN' if dry_run else 'APPLIED'
        phase_label = f' phase={result.get("migration_phase")}' if result.get('migration_phase') else ''
        self.stdout.write(self.style.SUCCESS(
            f"[{mode}{phase_label}] recipes_updated={result['updated']} created={result['created']} "
            f"skipped={result['skipped']} phase_skipped={result.get('phase_skipped', 0)} "
            f"missing_sku={result['missing_sku']} "
            f"inferred_product_type={result.get('inferred_product_type', 0)} rows={len(rows)}"
        ))
        if result['field_hits']:
            self.stdout.write('Fields filled (counts):')
            for name, count in sorted(result['field_hits'].items(), key=lambda item: (-item[1], item[0])):
                self.stdout.write(f'  {name}: {count}')
        if result['samples']:
            self.stdout.write('Samples:')
            for pk, sku, status, fields in result['samples']:
                self.stdout.write(f'  #{pk} [{status}] {sku}: {", ".join(fields)}')
