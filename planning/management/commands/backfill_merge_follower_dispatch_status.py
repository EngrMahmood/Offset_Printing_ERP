"""Start production on merge-group follower cards stuck at 'released'.

Before the merge-split auto-start fix, a lead job's printing entry was mirrored
onto every follower card's Production table (merge_parent set), but the
follower card itself was never moved out of 'released'. The current split
logic (core/signals.py: split_merge_group_printing_entry) now calls
start_production() on a follower before creating its derived entry, so this
only affects splits that ran before that fix shipped.

A card stuck this way has a real, already-recorded derived printing entry but
sits below JOB_CARD_DISPATCHABLE_STATUSES, so it silently never appears in
Dispatch Entry even though production is done.

Idempotent: only touches cards currently at 'released' that already have a
derived (merge_parent_id set) printing entry.
"""

from django.core.management.base import BaseCommand

from core.jobcard_service import start_production
from core.models import JobCard, Production


class Command(BaseCommand):
    help = "Backfill: start production on merge-follower cards stuck at 'released' so they show in dispatch."

    def add_arguments(self, parser):
        parser.add_argument('--apply', action='store_true', help='Persist changes (otherwise dry-run).')

    def handle(self, *args, **options):
        apply = options['apply']

        stuck_card_ids = (
            Production.objects.filter(
                entry_type='printing', is_active=True, merge_parent_id__isnull=False,
            )
            .values_list('job_card_id', flat=True)
            .distinct()
        )
        candidates = JobCard.objects.filter(id__in=stuck_card_ids, status='released', is_active=True)

        touched = 0
        for card in candidates:
            touched += 1
            if apply:
                start_production(
                    card,
                    reason='Backfill: merge-group split ran before the follower auto-start fix shipped',
                )
            self.stdout.write(f'{"started" if apply else "would start"} production on {card.job_card_no}')

        self.stdout.write(self.style.SUCCESS(
            f'{"Started" if apply else "Would start"} production on {touched} card(s).'
            + ('' if apply else ' Re-run with --apply to persist.')
        ))
