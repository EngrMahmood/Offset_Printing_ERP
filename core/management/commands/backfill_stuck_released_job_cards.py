"""Resume production on job cards stuck at 'released' despite already
having real printing/packing entries recorded against them.

Root cause: a job card can be reopened for a data correction (e.g. an
approved "Edit Reopen Request" for a machine/pass-count change) after
production had genuinely already started. Reopening resets the job card
all the way to 'draft' and walks it back through the full planning
approval pipeline (submit -> QC -> PM -> release) — but that pipeline
only ever lands on 'released', the same status a genuinely fresh job
sits at before any production has happened. Nothing along that path
re-checks whether production had already started before the reopen, so
the job silently stays at 'released' even though it has real production
history — and since JOB_CARD_DISPATCHABLE_STATUSES excludes 'released',
it disappears from Dispatch Entry despite having a real remaining
balance. core.jobcard_service.transition_job_card_status() now cascades
'released' -> 'in_production' automatically going forward whenever this
exact situation recurs; this command is a one-off backfill for job
cards already stuck this way today.

This is the general form of the fix already applied for merge-group
follower cards specifically (see backfill_merge_follower_dispatch_status
in the planning app) — same failure shape, not limited to merges.

Idempotent: only touches cards currently at 'released' that already have
an active printing or packing entry.
"""

from django.core.management.base import BaseCommand

from core.jobcard_service import transition_job_card_status
from core.models import JobCard


class Command(BaseCommand):
    help = "Backfill: resume production on job cards stuck at 'released' despite having recorded production entries."

    def add_arguments(self, parser):
        parser.add_argument('--apply', action='store_true', help='Persist changes (otherwise dry-run).')

    def handle(self, *args, **options):
        apply = options['apply']

        touched = 0
        for card in JobCard.objects.filter(is_active=True, status='released').select_related('planning_job'):
            if card.total_printed_pcs <= 0 and card.total_packed_pcs <= 0:
                continue
            touched += 1
            self.stdout.write(
                f'{"resuming" if apply else "would resume"} {card.job_card_no} '
                f'(printed={card.total_printed_pcs}, packed={card.total_packed_pcs}, '
                f'dispatched={card.total_dispatch}, balance={card.balance_qty})'
            )
            if apply:
                transition_job_card_status(
                    card,
                    'in_production',
                    reason='Backfill: production was already recorded before this job card was reopened and re-released',
                )

        self.stdout.write(self.style.SUCCESS(
            f'{"Resumed" if apply else "Would resume"} {touched} card(s).'
            + ('' if apply else ' Re-run with --apply to persist.')
        ))
