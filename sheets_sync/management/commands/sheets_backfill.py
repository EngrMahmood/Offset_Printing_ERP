from django.apps import apps
from django.core.management.base import BaseCommand, CommandError

from sheets_sync import client as sheets_client
from sheets_sync.registry import SYNCED_MODELS
from sheets_sync.serializers import SERIALIZERS
from sheets_sync.models import SheetsRowIndex

BATCH_SIZE = 500


class Command(BaseCommand):
    help = (
        "One-time push of existing database rows into the Google Sheets DR mirror. "
        "Run this once (with sync disabled) before enabling live incremental sync, "
        "so there is no gap between historical and incremental data."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--tabs', type=str, default='',
            help="Comma-separated dotted model paths to scope the backfill (e.g. core.JobCard). "
                 "Defaults to every model in the sync registry.",
        )
        parser.add_argument(
            '--force', action='store_true',
            help="Re-push tabs even if they already have a full SheetsRowIndex.",
        )

    def handle(self, *args, **options):
        requested = [p.strip() for p in options['tabs'].split(',') if p.strip()]
        entries = SYNCED_MODELS
        if requested:
            entries = [e for e in entries if e['dotted_path'] in requested]
            missing = set(requested) - {e['dotted_path'] for e in entries}
            if missing:
                raise CommandError(f"Unknown model(s) in --tabs: {', '.join(sorted(missing))}")

        try:
            spreadsheet = sheets_client.open_spreadsheet()
        except Exception as exc:
            raise CommandError(f"Could not open the configured spreadsheet: {exc}") from exc

        for entry in entries:
            self._backfill_entry(spreadsheet, entry, force=options['force'])

        self.stdout.write(self.style.SUCCESS("Backfill complete."))

    def _backfill_entry(self, spreadsheet, entry, force):
        app_label, model_name = entry['dotted_path'].split('.')
        model = apps.get_model(app_label, model_name)
        tab_name = entry['tab_name']
        headers = entry['headers']
        serializer = SERIALIZERS[entry['serializer']]

        total = model.objects.count()
        already_indexed = SheetsRowIndex.objects.filter(tab_name=tab_name).count()
        if total and already_indexed >= total and not force:
            self.stdout.write(f"Skipping {tab_name} ({already_indexed}/{total} already indexed). Use --force to redo.")
            return

        self.stdout.write(f"Backfilling {tab_name} ({total} records)...")
        worksheet = sheets_client.get_or_create_worksheet(spreadsheet, tab_name, headers)

        next_row = 2
        written = 0
        for chunk_start in range(0, total, BATCH_SIZE):
            queryset = model.objects.all().order_by('pk')[chunk_start:chunk_start + BATCH_SIZE]
            rows = []
            index_rows = []
            for instance in queryset.iterator(chunk_size=BATCH_SIZE):
                row_dict = serializer(instance, deleted=False)
                rows.append([row_dict.get(h, '') for h in headers])
                index_rows.append(SheetsRowIndex(
                    tab_name=tab_name, object_pk=str(instance.pk), row_number=next_row,
                ))
                next_row += 1

            if not rows:
                continue

            last_col = _col_letter(len(headers))
            start_row = next_row - len(rows)
            if worksheet.row_count < next_row:
                worksheet.add_rows(next_row - worksheet.row_count + 50)

            worksheet.update(
                f"A{start_row}:{last_col}{next_row - 1}", rows, value_input_option='RAW',
            )
            SheetsRowIndex.objects.bulk_create(index_rows, ignore_conflicts=True)
            written += len(rows)
            self.stdout.write(f"  ...{written}/{total} rows written")

        self.stdout.write(self.style.SUCCESS(f"{tab_name}: {written} rows backfilled."))


def _col_letter(n):
    letters = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters
