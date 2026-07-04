"""Restore blank SKU master + job purchase origin from a Google Sheet export."""

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from planning.sku_sheet_import import parse_sheet_rows, restore_sku_recipes_from_rows


class Command(BaseCommand):
    help = (
        'Fill blank SkuRecipe fields and planning-job purchase_material_origin '
        'from a Google Sheet CSV/XLSX export. Does not overwrite non-blank values '
        'and does not demote approved recipes.'
    )

    def add_arguments(self, parser):
        parser.add_argument('path', help='Path to CSV or XLSX export from Google Sheet')
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

    def handle(self, *args, **options):
        path = options['path']
        dry_run = options['dry_run']
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
            )

        if dry_run:
            with transaction.atomic():
                result = _run()
                transaction.set_rollback(True)
        else:
            result = _run()

        mode = 'DRY-RUN' if dry_run else 'APPLIED'
        self.stdout.write(self.style.SUCCESS(
            f"[{mode}] recipes_updated={result['updated']} created={result['created']} "
            f"skipped={result['skipped']} missing_sku={result['missing_sku']} "
            f"jobs_origin_updated={result.get('jobs_origin_updated', 0)} rows={len(rows)}"
        ))
        if result['field_hits']:
            self.stdout.write('Fields filled (counts):')
            for name, count in sorted(result['field_hits'].items(), key=lambda item: (-item[1], item[0])):
                label = name
                if name == 'purchase_material_origin':
                    label = 'purchase_material_origin (planning jobs)'
                self.stdout.write(f'  {label}: {count}')
        if result['samples']:
            self.stdout.write('Samples:')
            for pk, sku, status, fields in result['samples']:
                self.stdout.write(f'  #{pk} [{status}] {sku}: {", ".join(fields)}')
