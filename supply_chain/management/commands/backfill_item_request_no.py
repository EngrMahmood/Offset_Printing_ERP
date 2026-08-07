from django.core.management.base import BaseCommand

from supply_chain.item_request_service import generate_request_no
from supply_chain.models import ItemRequest


class Command(BaseCommand):
    help = (
        "One-time fix for item requests submitted before IR-IDs were assigned "
        "at submission time (they only got one once supply-chain approved them, "
        "so anything still mid-workflow shows a blank ID). Assigns each request "
        "missing a request_no its per-type-per-year IR-ID, oldest first so "
        "numbering still reflects submission order. Defaults to a dry run; pass "
        "--apply to write changes."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--apply',
            action='store_true',
            help='Actually write the changes. Without this flag, only reports what would change.',
        )

    def handle(self, *args, **options):
        apply_changes = options['apply']

        qs = (
            ItemRequest.objects
            .filter(request_no__isnull=True)
            .exclude(status='SUBMITTED')
            .select_related('request_type')
            .order_by('request_date', 'pk')
        )
        # SQLite can store '' instead of NULL for a blank CharField.
        qs = qs | ItemRequest.objects.filter(request_no='').exclude(status='SUBMITTED').select_related('request_type').order_by('request_date', 'pk')
        pending = list({ir.pk: ir for ir in qs}.values())
        pending.sort(key=lambda ir: (ir.request_date, ir.pk))

        if not pending:
            self.stdout.write('No item requests need an IR-ID.')
            return

        for item_request in pending:
            if not item_request.request_type_id:
                self.stdout.write(self.style.WARNING(
                    f'Skipping #{item_request.pk} "{item_request.item_title}" — no request_type set.'
                ))
                continue
            if apply_changes:
                generate_request_no(item_request)
                self.stdout.write(self.style.SUCCESS(
                    f'#{item_request.pk} "{item_request.item_title}" -> {item_request.request_no}'
                ))
            else:
                self.stdout.write(
                    f'Would assign an IR-ID to #{item_request.pk} "{item_request.item_title}" '
                    f'(status={item_request.status}, type={item_request.request_type.code})'
                )

        if not apply_changes:
            self.stdout.write(self.style.WARNING('Dry run — re-run with --apply to write changes.'))
