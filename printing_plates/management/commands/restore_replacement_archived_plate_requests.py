"""One-off data fix: request_plate_remake() used to auto-archive the prior
plate request whenever a replacement was created, even though the prior
request was a completed, successful transaction (e.g. already Released to
Production). That auto-archive side effect has been removed from the code
(see request_plate_remake in printing_plates/services.py), but requests
archived by the old behavior before this fix stay archived until corrected
here.

Scope: PlateRequest rows with status=archived that are referenced by a
newer request's replaces_request FK (i.e. they were superseded by a
replacement, not genuinely cancelled). Genuine cancels (cancel_plate_request,
bulk stale-cleanup) have no replacement_requests and are left untouched.
"""

from django.core.management.base import BaseCommand
from django.utils import timezone

from printing_plates.models import PlateRequest


class Command(BaseCommand):
    help = (
        'Restore plate requests that were wrongly archived when a replacement was '
        'created against them, back to Available for Production.'
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Report which requests would be restored without changing data.',
        )

    def handle(self, *args, **options):
        dry_run = options['dry_run']

        qs = PlateRequest.objects.filter(
            status=PlateRequest.STATUS_ARCHIVED,
            replacement_requests__isnull=False,
        ).distinct().order_by('id')

        rows = list(qs)
        mode = 'DRY-RUN' if dry_run else 'APPLIED'
        self.stdout.write(f'[{mode}] {len(rows)} archived plate request(s) to restore.')
        for row in rows:
            replaced_by = list(row.replacement_requests.values_list('id', flat=True))
            self.stdout.write(
                f'  PR#{row.id} (JC {row.job_card_id}) -> restore to '
                f'"{PlateRequest.STATUS_AVAILABLE}"; replaced by PR#{replaced_by}'
            )

        if dry_run or not rows:
            return

        now = timezone.now()
        for row in rows:
            row.status = PlateRequest.STATUS_AVAILABLE
            row.updated_at = now
            row.save(update_fields=['status', 'updated_at'])

        self.stdout.write(self.style.SUCCESS(f'[{mode}] Restored {len(rows)} plate request(s).'))
