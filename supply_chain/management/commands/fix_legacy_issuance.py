from django.core.management.base import BaseCommand
from django.db import transaction

from core.models import JOB_CARD_PRODUCTION_START_STATUSES, JobCard
from supply_chain.jc_sync import _planned_purchase_sheet_qty
from supply_chain.models import StockTransaction


class Command(BaseCommand):
    help = (
        "One-time fix for job cards that left the released-to-production set "
        "(completed/closed/etc.) before the one-row-per-JC and purchase-sheet "
        "unit fixes shipped. Their approved ISSUANCE rows were never revisited "
        "by the normal sync (which only acts on released/in_production job "
        "cards), so some still hold duplicate per-production rows and/or "
        "quantities in press sheets instead of purchase sheets. This command "
        "collapses each affected job card to a single approved row in purchase "
        "sheets, preserving approval. Defaults to a dry run; pass --apply to "
        "write changes."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--apply',
            action='store_true',
            help='Actually write the changes. Without this flag, only reports what would change.',
        )

    def handle(self, *args, **options):
        apply_changes = options['apply']

        txns = (
            StockTransaction.objects
            .filter(
                source='JOB_CARD',
                transaction_type='ISSUANCE',
                is_active=True,
                is_approved=True,
            )
            .select_related('job_card')
            .order_by('job_card_id', 'id')
        )

        by_jc = {}
        for txn in txns:
            if not txn.job_card:
                continue
            by_jc.setdefault(txn.job_card_id, []).append(txn)

        affected = 0
        for job_card_id, rows in by_jc.items():
            job_card = rows[0].job_card
            if job_card.workflow_status in JOB_CARD_PRODUCTION_START_STATUSES:
                # Still released/in_production: the normal sync already keeps
                # this one correct on every save/approval.
                continue

            correct_qty = _planned_purchase_sheet_qty(job_card)
            needs_collapse = len(rows) > 1
            needs_conversion = (
                not needs_collapse
                and correct_qty > 0
                and rows[0].sheet_qty_pcs != correct_qty
            )
            if not needs_collapse and not needs_conversion:
                continue

            affected += 1
            current_qtys = [r.sheet_qty_pcs for r in rows]
            self.stdout.write(
                f"{job_card.job_card_no} ({job_card.workflow_status}): "
                f"{len(rows)} row(s) {current_qtys} -> 1 row [{correct_qty}] purchase sheets"
            )

            if not apply_changes:
                continue

            with transaction.atomic():
                keep = rows[0]
                for extra in rows[1:]:
                    extra.delete()
                if correct_qty > 0:
                    keep.sheet_qty_pcs = correct_qty
                    keep.save(update_fields=['sheet_qty_pcs'])
                else:
                    # No planned quantity / UPS available to convert — leave the
                    # quantity as-is but still collapse duplicates.
                    self.stdout.write(self.style.WARNING(
                        f"  could not determine purchase-sheet qty for "
                        f"{job_card.job_card_no}; kept existing quantity {keep.sheet_qty_pcs}"
                    ))

        if affected == 0:
            self.stdout.write(self.style.SUCCESS("No legacy issuance rows need fixing."))
        elif apply_changes:
            self.stdout.write(self.style.SUCCESS(f"Fixed {affected} job card(s)."))
        else:
            self.stdout.write(self.style.WARNING(
                f"{affected} job card(s) would be changed. Re-run with --apply to write changes."
            ))
