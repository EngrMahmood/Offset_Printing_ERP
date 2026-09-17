from datetime import date

from django.contrib.auth import get_user_model
from django.core.exceptions import ValidationError
from django.core.management.base import BaseCommand

from core.jobcard_service import ensure_job_card_from_planning_job
from core.models import Dispatch, Machine, Production, Sorter
from planning.models import PlanningJob, SkuRecipe
from planning.services import create_sibling_forms_from_sku_master

DEMO_MACHINE = 'Demo GTO 1'
DEMO_SORTER = 'Demo Sorter'
DEMO_USER = 'demo_seed_user'


class Command(BaseCommand):
    help = (
        "Seeds dummy SKU masters + jobs for the order-quantity shapes the ERP must "
        "tell apart from a plain WO/PO quantity — a normal pcs book, an A4 rim cut & "
        "pack job, and a multi-ply NCR triplicate book — then exercises "
        "printing/packing/dispatch on each (including validations that should reject "
        "bad entries) and prints a report. order_qty and dispatch_qty are tracked in "
        "the SAME unit as the customer's WO/PO for every case (pcs, rims, or books); "
        "pcs_per_unit is used only internally for press-sheet/ups planning and the "
        "packing-granularity check. Safe to re-run."
    )

    def handle(self, *args, **options):
        self.user, _ = get_user_model().objects.get_or_create(
            username=DEMO_USER, defaults={'is_staff': True}
        )
        self.machine, _ = Machine.objects.get_or_create(name=DEMO_MACHINE)
        self.sorter, _ = Sorter.objects.get_or_create(name=DEMO_SORTER)

        self.stdout.write(self.style.MIGRATE_HEADING('\n=== Case A: plain pcs book (single form) ==='))
        self.run_plain_pcs_case()

        self.stdout.write(self.style.MIGRATE_HEADING('\n=== Case B: A4 rim cut & pack job ==='))
        self.run_rim_case()

        self.stdout.write(self.style.MIGRATE_HEADING('\n=== Case C: NCR triplicate multi-form book ==='))
        self.run_multi_form_case()

        self.stdout.write(self.style.SUCCESS('\nAll demo cases completed.'))

    # ------------------------------------------------------------------
    def _release_job_card(self, job_card, **extra_fields):
        """Push a freshly-synced (pending_data) job card straight to
        in_production for the demo, filling in the fields planning approval
        would normally have required along the way."""
        job_card.po_date = date(2026, 1, 1)
        job_card.total_colors = job_card.total_colors or 1
        job_card.total_sheet_quantity = job_card.total_sheet_quantity or 0
        job_card.machine_name = self.machine
        for field, value in extra_fields.items():
            setattr(job_card, field, value)
        job_card.status = 'in_production'
        job_card.save()
        return job_card

    # ------------------------------------------------------------------
    def run_plain_pcs_case(self):
        """PO-179570-style SKU: 'STATIONERYKNITSDAILYCHECKINGREPORTDUPLICATE...'
        line said '100 PIECE' — for THIS SKU that plainly means 100 pcs, no
        SKU-master rim/book flag needed."""
        SkuRecipe.objects.get_or_create(
            sku='DEMO PLAIN BOOK',
            defaults={'job_process_type': 'print_and_pack', 'ups': 10},
        )
        planning_job, _ = PlanningJob.objects.update_or_create(
            jc_number='JC-DEMO-PLAIN-001',
            defaults={
                'sku': 'DEMO PLAIN BOOK', 'order_qty': 1000, 'ups': 10,
                'wastage_sheets': 0, 'status': 'draft', 'created_by': self.user,
                'plan_date': date(2026, 1, 1),
            },
        )
        self.stdout.write(
            f'PlanningJob {planning_job.jc_number}: unit_type={planning_job.unit_type!r}, '
            f'pcs_per_unit={planning_job.pcs_per_unit} (plain SKU — no master override, so pcs/1)'
        )

        job_card, _ = ensure_job_card_from_planning_job(planning_job, actor=self.user)
        self._release_job_card(job_card, total_sheet_quantity=100)

        if job_card.dispatch_set.filter(is_active=True).exists():
            self.stdout.write(f'  {job_card.job_card_no}: already seeded, skipping printing/packing/dispatch.')
        else:
            Production.objects.create(
                job_card=job_card, entry_type='printing', date='2026-01-02', shift='A',
                machine=self.machine, output_sheets=100, waste_sheets=0, impressions=100,
                planned_time=60, run_time=60,
            )
            Production.objects.create(
                job_card=job_card, entry_type='packing', date='2026-01-02', shift='A',
                packing_qty=1000, sorting_waste_qty=0, sorter=self.sorter,
            )
            Dispatch.objects.create(
                job_card=job_card, dc_no='DC-DEMO-PLAIN-1', dispatch_date='2026-01-03',
                dispatch_qty=1000, created_by=self.user,
            )
        job_card.refresh_from_db()
        self.stdout.write(self.style.SUCCESS(
            f'  {job_card.job_card_no}: order_qty={job_card.order_qty} pcs, order_qty_pcs={job_card.order_qty_pcs}, '
            f'dispatched={job_card.total_dispatch} pcs, balance={job_card.balance_qty} pcs'
        ))

    # ------------------------------------------------------------------
    def run_rim_case(self):
        """PO-179887-style SKU: 'A4 RIM PAPER' line said '20.0 BOX' (1 box =
        5 rims = 2,500 pcs per your CSV remark). Here we model the simpler
        'A3 PAPER RIM' line, which said '5.0 PIECE' meaning 5 rims (1 rim =
        500 pcs) — the PO's own unit word ('PIECE') is misleading; only the
        SKU master's unit_type='rim'/pcs_per_unit=500 tells the system the
        truth, and the WO/PO number (5) is what gets typed into order_qty
        and, later, dispatch_qty — never a pcs-converted number."""
        # unit_type/pcs_per_unit are NOT set here — picking product_type='A4
        # Rim' (the real master-data category from your screenshot) is what
        # auto-fills them, via SkuRecipe.sync_unit_defaults_from_product_type.
        recipe, created_recipe = SkuRecipe.objects.get_or_create(
            sku='DEMO A4 PAPER RIM',
            defaults={'job_process_type': 'cut_and_pack', 'product_type': 'A4 Rim'},
        )
        if not created_recipe and recipe.product_type != 'A4 Rim':
            recipe.product_type = 'A4 Rim'
            recipe.job_process_type = 'cut_and_pack'
            recipe.save()
        planning_job, _ = PlanningJob.objects.update_or_create(
            jc_number='JC-DEMO-RIM-001',
            defaults={
                # The WO/PO says "50" — 50 rims. That is exactly what goes
                # into order_qty; the system converts to pcs only internally.
                'sku': 'DEMO A4 PAPER RIM', 'order_qty': 50,
                'status': 'draft', 'created_by': self.user, 'plan_date': date(2026, 1, 1),
            },
        )
        self.stdout.write(
            f'PlanningJob {planning_job.jc_number}: order_qty={planning_job.order_qty} '
            f'{planning_job.unit_type}s (matches the WO/PO exactly), pcs_per_unit={planning_job.pcs_per_unit} '
            f'(inherited from the SKU master, not guessed from the PO text)'
        )

        job_card, _ = ensure_job_card_from_planning_job(planning_job, actor=self.user)
        self._release_job_card(job_card)
        self.stdout.write(
            f'  {job_card.job_card_no}: order_qty={job_card.order_qty} rims, '
            f'order_qty_pcs={job_card.order_qty_pcs} (internal press/packing planning only)'
        )

        if job_card.dispatch_set.filter(is_active=True).exists():
            self.stdout.write(f'  {job_card.job_card_no}: already seeded, skipping packing/dispatch checks.')
            job_card.refresh_from_db()
            self.stdout.write(self.style.SUCCESS(f'  balance={job_card.balance_qty} rims'))
            return

        # Packing still happens in pcs (bundling loose sheets) and must round
        # to a whole rim (500 pcs) — this check is unaffected by the pivot.
        bad_packing = Production(
            job_card=job_card, entry_type='packing', date='2026-01-02', shift='A',
            packing_qty=750, sorting_waste_qty=0, sorter=self.sorter,
        )
        try:
            bad_packing.save()
            self.stdout.write(self.style.ERROR('  UNEXPECTED: 750-pc packing (not a whole rim) was accepted!'))
        except ValidationError as exc:
            self.stdout.write(self.style.SUCCESS(f'  Correctly rejected 750-pc packing entry: {exc.messages}'))

        Production.objects.create(
            job_card=job_card, entry_type='packing', date='2026-01-02', shift='A',
            packing_qty=1000, sorting_waste_qty=0, sorter=self.sorter,
        )
        self.stdout.write(self.style.SUCCESS('  Accepted 1000-pc (2 rim) packing entry.'))

        # Dispatch is entered directly in rims (2), matching how the WO/PO
        # and the printed DC both state the quantity — not 1000 pcs.
        Dispatch.objects.create(
            job_card=job_card, dc_no='DC-DEMO-RIM-1', dispatch_date='2026-01-03',
            dispatch_qty=2, created_by=self.user,
        )
        job_card.refresh_from_db()
        self.stdout.write(self.style.SUCCESS(
            f'  Dispatched 2 rims (DC shows "2 Rims", matching the PO). balance={job_card.balance_qty} rims'
        ))

        # Over-dispatch (in rims, against the 50-rim order) is still blocked.
        over_dispatch = Dispatch(
            job_card=job_card, dc_no='DC-DEMO-RIM-BAD', dispatch_date='2026-01-03',
            dispatch_qty=49, created_by=self.user,
        )
        try:
            over_dispatch.save()
            self.stdout.write(self.style.ERROR('  UNEXPECTED: dispatch past order_qty was accepted!'))
        except ValidationError as exc:
            self.stdout.write(self.style.SUCCESS(f'  Correctly rejected over-dispatch: {exc.messages}'))

    # ------------------------------------------------------------------
    def run_multi_form_case(self):
        """PO-177042-style SKU: 'STATIONERYDAILYTPULAMINATEDFABRICPRODUCTION
        CHALLAN...-NCR' line said '30.0 BOOK'. Product Type 'Report Books'
        (your real master-data category) auto-fills unit_type='book', but its
        page count has no fixed default (it varies per document) — so this
        SKU explicitly sets pcs_per_unit=50, since it's a triplicate (3-ply)
        NCR book where each of White/Pink/Yellow only carries 50 of the
        book's leaves. That number is on the SHARED SKU master, so it applies
        identically to all three forms because they're all the same SKU."""
        recipe, _ = SkuRecipe.objects.update_or_create(
            sku='DEMO NCR TRIPLICATE BOOK',
            defaults={
                'job_process_type': 'print_and_pack', 'ups': 2,
                'product_type': 'Report Books', 'unit_type': 'book', 'pcs_per_unit': 50,
                'default_form_labels': 'White,Pink,Yellow',
            },
        )
        base_planning_job, _ = PlanningJob.objects.update_or_create(
            jc_number='JC-DEMO-NCR-001',
            defaults={
                # The WO/PO says "500" — 500 books. Each ply independently
                # needs one White/Pink/Yellow leaf per book, so all three
                # sibling forms carry the SAME order_qty (books), matching
                # your CSV's JC-02-26-1585 / .1 / .2 pattern exactly.
                'sku': 'DEMO NCR TRIPLICATE BOOK', 'order_qty': 500, 'ups': 2,
                'wastage_sheets': 0, 'status': 'draft', 'created_by': self.user,
                'plan_date': date(2026, 1, 1),
            },
        )
        siblings = create_sibling_forms_from_sku_master(base_planning_job, recipe, actor=self.user)
        base_planning_job.refresh_from_db()
        self.stdout.write(
            f'SKU master default_form_labels={recipe.default_form_labels!r}, pcs_per_unit='
            f'{recipe.pcs_per_unit} (pages/book, per ply) auto-created {len(siblings)} sibling '
            f'planning job(s) alongside the base job.'
        )

        all_planning_jobs = [base_planning_job] + list(
            base_planning_job.child_forms.order_by('jc_number')
        )
        job_cards = []
        for planning_job in all_planning_jobs:
            job_card, _ = ensure_job_card_from_planning_job(planning_job, actor=self.user)
            # required_sheets = order_qty_pcs / ups = 25,000 / 2 = 12,500 press sheets.
            self._release_job_card(job_card, total_sheet_quantity=12500)
            job_cards.append(job_card)

        base_card = job_cards[0]
        self.stdout.write(f'  Group under {base_card.job_card_no} (order_qty={base_card.order_qty} books each):')
        for jc in base_card.sibling_forms:
            self.stdout.write(
                f'    {jc.job_card_no}  form={jc.form_label!r}  order_qty={jc.order_qty} books  '
                f'order_qty_pcs={jc.order_qty_pcs} (press-planning only)'
            )

        # Printing/packing still happen per form (each ply is a separate
        # press run/material) — this part is unchanged by the unit pivot.
        for jc in job_cards:
            if jc.productions.filter(is_active=True).exists():
                self.stdout.write(f'    {jc.job_card_no}: printing/packing already seeded, skipping.')
                continue
            Production.objects.create(
                job_card=jc, entry_type='printing', date='2026-01-02', shift='A',
                machine=self.machine, output_sheets=12500, waste_sheets=0, impressions=12500,
                planned_time=120, run_time=120,
            )
            Production.objects.create(
                job_card=jc, entry_type='packing', date='2026-01-02', shift='A',
                packing_qty=25000, sorting_waste_qty=0, sorter=self.sorter,
            )
            jc.refresh_from_db()
            self.stdout.write(self.style.SUCCESS(
                f'    {jc.job_card_no} ({jc.form_label}): printed 12,500 sheets, packed 25,000 pcs.'
            ))

        # A sibling ply is NOT itself dispatched — only the collated book
        # (base job card) is, in books, matching the WO/PO's own unit.
        pink_form = next(jc for jc in job_cards if jc.form_label == 'Pink')
        if not pink_form.dispatch_set.filter(is_active=True).exists():
            bad_dispatch = Dispatch(
                job_card=pink_form, dc_no='DC-DEMO-NCR-BAD', dispatch_date='2026-01-03',
                dispatch_qty=500, created_by=self.user,
            )
            try:
                bad_dispatch.save()
                self.stdout.write(self.style.ERROR('  UNEXPECTED: dispatch on a sibling ply was accepted!'))
            except ValidationError as exc:
                self.stdout.write(self.style.SUCCESS(f'  Correctly rejected dispatch on the Pink ply directly: {exc.messages}'))

        if not base_card.dispatch_set.filter(is_active=True).exists():
            Dispatch.objects.create(
                job_card=base_card, dc_no='DC-DEMO-NCR-1', dispatch_date='2026-01-03',
                dispatch_qty=500, created_by=self.user,
            )
        base_card.refresh_from_db()
        self.stdout.write(self.style.SUCCESS(
            f'  Dispatched 500 books against {base_card.job_card_no} (DC shows "500 Books", matching the PO). '
            f'balance={base_card.balance_qty} books'
        ))
