"""One-off data fix for the PR-then-PO duplicate job card bug.

Manual PO/PR entry used to leave PlanningJob.po_number set to the literal
placeholder "-" (or similar: "N/A", "NA", em-dash) instead of blank whenever
only a PR number was entered. The PR->PO reconciliation in
_sync_repeat_jobs_from_po only ever matched jobs with po_number == '', so
those placeholder rows were invisible to it: when the real PO later arrived
for the same SKU/qty, reconciliation found no candidate and created a brand
new duplicate draft job card instead of linking the PO onto the original.

That input has been normalized at the source (manual PO entry, PDF
extraction, manual PO-number update) so this can no longer happen going
forward. This command is the one-off cleanup for rows already affected:

1. Normalizes any existing po_number placeholder values to ''.
2. Finds duplicate pairs: an "original" PR-only job (po_number now blank)
   and a later "duplicate" draft job with the same SKU + order_qty that
   already carries a real po_number, where the duplicate is still an empty
   draft (no job card, no plate requests) — safe to fold without losing any
   work. Only auto-merges when exactly one such duplicate matches a given
   original; ambiguous matches are left alone for manual review.

Folding: the original job keeps its jc_number/pr_reference and gets the
duplicate's po_number (and po_approval_date if it didn't have one). The
duplicate is deactivated (is_active=False) with an audit note, never
deleted.
"""

from django.core.management.base import BaseCommand
from django.db import transaction
from django.utils import timezone

from planning.models import PlanningJob
from planning.services import _sku_key

PLACEHOLDER_PO_NUMBERS = ('-', '—', 'N/A', 'NA')


class Command(BaseCommand):
    help = 'Fix PlanningJob po_number placeholders and fold PR/PO duplicate job cards created by that bug.'

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Report what would change without saving anything.',
        )

    def handle(self, *args, **options):
        dry_run = options['dry_run']
        mode = 'DRY-RUN' if dry_run else 'APPLIED'

        placeholder_jobs = list(PlanningJob.objects.filter(po_number__in=PLACEHOLDER_PO_NUMBERS))
        self.stdout.write(f'[{mode}] {len(placeholder_jobs)} job(s) with a placeholder po_number to normalize.')
        for job in placeholder_jobs[:30]:
            self.stdout.write(f'  {job.jc_number}: po_number {job.po_number!r} -> \'\'')
        if len(placeholder_jobs) > 30:
            self.stdout.write(f'  ... and {len(placeholder_jobs) - 30} more')

        if not dry_run and placeholder_jobs:
            with transaction.atomic():
                PlanningJob.objects.filter(id__in=[j.id for j in placeholder_jobs]).update(
                    po_number='', updated_at=timezone.now(),
                )

        # Determine who qualifies as a "PR-only" original: either already
        # blank, or blank as of the normalization pass above (checked against
        # placeholder_ids rather than re-querying, so this works identically
        # under --dry-run before any save has happened).
        placeholder_ids = {j.id for j in placeholder_jobs}
        originals = []
        for job in PlanningJob.objects.filter(is_active=True):
            po_is_blank = (job.id in placeholder_ids) or not (job.po_number or '').strip()
            if po_is_blank:
                originals.append(job)

        merges = []
        ambiguous = []
        for original in originals:
            sku_key = _sku_key(original.sku)
            if not sku_key or original.order_qty is None:
                continue
            candidates = [
                dup for dup in PlanningJob.objects.filter(
                    is_active=True,
                    status='draft',
                    sku__iexact=sku_key,
                    order_qty=original.order_qty,
                ).exclude(id=original.id)
                if (dup.po_number or '').strip() and dup.po_number not in PLACEHOLDER_PO_NUMBERS
                and not hasattr(dup, 'job_card')
                and not dup.plate_requests.exists()
            ]
            if len(candidates) == 1:
                merges.append((original, candidates[0]))
            elif len(candidates) > 1:
                ambiguous.append((original, candidates))

        self.stdout.write(f'[{mode}] {len(merges)} duplicate pair(s) to fold.')
        for original, dup in merges:
            self.stdout.write(
                f'  fold {dup.jc_number} (PO {dup.po_number}) -> {original.jc_number} '
                f'(SKU {original.sku}, qty {original.order_qty})'
            )
        if ambiguous:
            self.stdout.write(
                self.style.WARNING(
                    f'{len(ambiguous)} original job(s) have more than one matching duplicate — '
                    'left untouched, review manually:'
                )
            )
            for original, candidates in ambiguous:
                dup_list = ', '.join(d.jc_number for d in candidates)
                self.stdout.write(f'  {original.jc_number} (SKU {original.sku}): {dup_list}')

        if dry_run or not merges:
            return

        now = timezone.now()
        with transaction.atomic():
            for original, dup in merges:
                original.po_number = dup.po_number
                if not original.po_approval_date and dup.po_approval_date:
                    original.po_approval_date = dup.po_approval_date
                note = f'[Auto-merge] Duplicate JC {dup.jc_number} folded into this job on PO link; duplicate deactivated.'
                original.requirement = (original.requirement + '\n' + note).strip() if original.requirement else note
                original.updated_at = now
                original.save(update_fields=['po_number', 'po_approval_date', 'requirement', 'updated_at'])

                dup_note = f'[Auto-merge] Superseded by {original.jc_number} — same SKU/qty/PO, folded to avoid duplicate job card.'
                dup.is_active = False
                dup.requirement = (dup.requirement + '\n' if dup.requirement else '') + dup_note
                dup.updated_at = now
                dup.save(update_fields=['is_active', 'requirement', 'updated_at'])

        self.stdout.write(self.style.SUCCESS(f'[{mode}] Normalized {len(placeholder_jobs)} job(s), folded {len(merges)} duplicate pair(s).'))
