from decimal import Decimal

from django.contrib.auth import get_user_model
from django.core.exceptions import ValidationError
from django.test import TestCase
from django.urls import reverse

from core.models import Machine

from . import services
from .calculations import compute_print_plan, pcs_per_print_meter
from .models import FlexoJobCard, FlexoPlanningJob, FlexoSkuRecipe
from .numbering import allocate_next_flexo_jc_number

User = get_user_model()


class PrintPlanCalculationTests(TestCase):
    def test_fl_sample_from_jc_template(self):
        # 18,360 pcs at 5,000 pcs/roll, 142 m/roll, 11 ups, 5% wastage
        plan = compute_print_plan(
            order_qty=18360, ups=11, meter_per_roll=142, pcs_per_roll=5000, wastage_pct=5)
        self.assertAlmostEqual(float(plan.order_rolls), 3.672, places=3)
        self.assertAlmostEqual(float(plan.base_meters), 47.402, places=3)
        self.assertAlmostEqual(float(plan.wastage_meters), 2.3701, places=4)
        self.assertAlmostEqual(float(plan.print_meters), 49.7723, places=4)

    def test_sheet_formula_rolls_times_meter_per_roll_over_ups(self):
        # 'Merge Sticker Satin Planning' row: 297 rolls x 106 m / 3 ups = 10,494 m
        plan = compute_print_plan(order_qty=297 * 1000, pcs_per_roll=1000, meter_per_roll=106, ups=3, wastage_pct=Decimal('4.9'))
        self.assertAlmostEqual(float(plan.base_meters), 10494, places=2)
        self.assertAlmostEqual(float(plan.wastage_meters), 514.206, places=3)
        self.assertAlmostEqual(float(plan.print_meters), 11008.206, places=3)

    def test_stock_matches_sheet_row_with_stock(self):
        # sheet row: (20 rolls - 14 stock) x 137 m / 4 ups = 205.5 m
        plan = compute_print_plan(order_qty=100000, stock_qty=70000, pcs_per_roll=5000,
                                  meter_per_roll=137, ups=4, wastage_pct=5)
        self.assertAlmostEqual(float(plan.base_meters), 205.5, places=3)
        self.assertAlmostEqual(float(plan.wastage_meters), 10.275, places=3)

    def test_pcs_per_roll_is_derived_from_piece_length_when_blank(self):
        # 50,000 pcs x 304.8 mm, 200 m/roll, 1 up -> 15,240 m (+5% = 16,002 m)
        plan = compute_print_plan(order_qty=50000, ups=1, size_h_mm=Decimal('304.8'),
                                  meter_per_roll=200, wastage_pct=5)
        self.assertAlmostEqual(float(plan.base_meters), 15240.0, places=2)
        self.assertAlmostEqual(float(plan.print_meters), 16002.0, places=2)

    def test_roll_orders_use_the_sheet_formula_directly(self):
        # label PO line: 138 ROLL, 106 m/roll, 3 ups -> 138 x 106 / 3 = 4,876 m (no pcs-per-roll needed)
        plan = compute_print_plan(order_qty=138, order_unit='roll', meter_per_roll=106, ups=3, wastage_pct=5)
        self.assertAlmostEqual(float(plan.base_meters), 4876.0, places=2)
        self.assertAlmostEqual(float(plan.order_rolls), 138.0, places=3)
        self.assertEqual(plan.net_qty, 138)

    def test_roll_orders_subtract_stock_in_rolls(self):
        plan = compute_print_plan(order_qty=20, stock_qty=14, order_unit='roll', meter_per_roll=137, ups=4, wastage_pct=5)
        self.assertAlmostEqual(float(plan.base_meters), 205.5, places=3)

    def test_stock_never_makes_net_negative(self):
        plan = compute_print_plan(order_qty=100, stock_qty=500, ups=1, meter_per_roll=100, pcs_per_roll=100)
        self.assertEqual((plan.net_qty, plan.print_meters), (0, 0))

    def test_missing_inputs_give_zero(self):
        self.assertEqual(compute_print_plan(order_qty=500, ups=None, meter_per_roll=100, pcs_per_roll=100).print_meters, 0)
        self.assertEqual(compute_print_plan(order_qty=500, ups=2, meter_per_roll=100).print_meters, 0)

    def test_pcs_per_print_meter_bridge(self):
        # 5,000 pcs per 142 m roll across 11 ups
        self.assertAlmostEqual(
            float(pcs_per_print_meter(ups=11, meter_per_roll=142, pcs_per_roll=5000)), 5000 * 11 / 142, places=3)


class FlexoFixtureMixin:
    def setUp(self):
        self.user = User.objects.create_user('planner', password='pw')
        self.machine = Machine.objects.create(name='Flexo-1', machine_type='flexo_printing')
        self.fl = self.make_recipe('LABEL-SKU-1', size_h_mm=Decimal('25.4'), die_code='FD1X101', pcs_per_roll=5000)
        self.satin = self.make_recipe('RIBBON-SKU-1', size_h_mm=Decimal('304.8'), die_code='2', ups=Decimal('1'),
                                     meter_per_roll=Decimal('200'))

    def make_recipe(self, sku, approved=True, **extra):
        data = dict(
            sku=sku, job_name=sku, material='TTH', size_w_mm=Decimal('25.4'),
            ups=Decimal('11'), roll_size=Decimal('316'), meter_per_roll=Decimal('142'), machine=self.machine,
            master_data_status='approved' if approved else 'draft',
        )
        data.update(extra)
        return FlexoSkuRecipe.objects.create(**data)

    def make_job(self, recipe=None, qty=18360, **kwargs):
        from datetime import date
        return services.create_job_from_recipe(
            recipe or self.fl, actor=self.user, po_number='WO-1', po_received_date=date(2026, 9, 1),
            order_qty=qty, **kwargs)


class JobCreationTests(FlexoFixtureMixin, TestCase):
    def test_job_created_from_recipe_calculates_meters(self):
        job = self.make_job()
        self.assertEqual(job.status, 'draft')
        self.assertTrue(job.jc_number.startswith('JC-FPP-'))
        self.assertAlmostEqual(float(job.print_meters), 49.772, places=3)
        self.assertEqual(job.balance_qty, 18360)
        self.assertEqual(job.die_code, 'FD1X101')

    def test_unapproved_sku_cannot_create_job(self):
        draft = self.make_recipe('DRAFT-SKU', approved=False, size_h_mm=Decimal('10'), pcs_per_roll=1000)
        with self.assertRaises(ValidationError):
            self.make_job(draft)

    def test_second_job_for_same_sku_is_repeat(self):
        self.make_job()
        self.assertEqual(self.make_job().repeat_flag, 'Repeat')

    def test_numbering_is_one_running_series(self):
        from datetime import date
        d = date(2026, 3, 10)
        numbers = [allocate_next_flexo_jc_number(for_date=d) for _ in range(3)]
        self.assertEqual(numbers, ['JC-FPP-03-26-01', 'JC-FPP-03-26-02', 'JC-FPP-03-26-03'])

    def test_numbering_resets_each_year_but_not_each_month(self):
        from datetime import date
        self.assertEqual(allocate_next_flexo_jc_number(for_date=date(2026, 12, 30)), 'JC-FPP-12-26-01')
        self.assertEqual(allocate_next_flexo_jc_number(for_date=date(2027, 1, 2)), 'JC-FPP-01-27-01')
        self.assertEqual(allocate_next_flexo_jc_number(for_date=date(2027, 2, 1)), 'JC-FPP-02-27-02')

    def test_numbering_continues_after_existing_jobs(self):
        job = self.make_job()                       # takes the next number for today's year
        nxt = allocate_next_flexo_jc_number()
        self.assertEqual(int(nxt.rsplit('-', 1)[1]), int(job.jc_number.rsplit('-', 1)[1]) + 1)

    def test_balance_follows_dispatch_runs(self):
        job = self.make_job(qty=1000)
        job.dispatch_runs.create(dispatch_index=1, dc_no='DC1', delivered_qty=400)
        job.save()
        self.assertEqual(job.total_dispatched, 400)
        self.assertEqual(job.balance_qty, 600)


class WorkflowTests(FlexoFixtureMixin, TestCase):
    def test_qc_flow_creates_job_card_and_releases(self):
        job = self.make_job()
        job.daily_demand = Decimal('1000')
        job.save()
        services.submit_to_qc(job, actor=self.user)
        job.refresh_from_db()
        self.assertEqual(job.status, 'pending_qc')
        card = FlexoJobCard.objects.get(planning_job=job)
        self.assertEqual(card.status, 'pending_qc')
        self.assertEqual(card.roll_meter_qty, job.print_meters)
        self.assertEqual(card.mtr_per_roll, '142')       # not '1.42E+2'
        self.assertEqual(card.days_plan, 19)              # ceil(18,360 pcs / 1,000 per day)

        services.approve_qc(job, actor=self.user)
        services.release_for_production(job, actor=self.user)
        job.refresh_from_db()
        card.refresh_from_db()
        self.assertEqual((job.status, card.status), ('released', 'released'))
        self.assertTrue(job.issued_to_production)
        self.assertEqual(job.events.count(), 4)

    def test_submit_blocked_without_machine(self):
        job = self.make_job()
        job.machine_name = None
        job.save()
        with self.assertRaises(ValidationError):
            services.submit_to_qc(job, actor=self.user)

    def test_reject_requires_reason_and_returns_to_draft(self):
        job = self.make_job()
        services.submit_to_qc(job, actor=self.user)
        with self.assertRaises(ValidationError):
            services.reject_qc(job, actor=self.user, reason='')
        services.reject_qc(job, actor=self.user, reason='Wrong die')
        job.refresh_from_db()
        self.assertEqual((job.status, job.rejection_reason), ('draft', 'Wrong die'))

    def test_invalid_transition_is_refused(self):
        job = self.make_job()
        with self.assertRaises(ValidationError):
            services.release_for_production(job, actor=self.user)

    def test_direct_cancel_pre_release_archives_job(self):
        job = self.make_job()
        services.cancel_job(job, actor=self.user, reason='Customer withdrew', reason_code='customer_cancelled')
        job.refresh_from_db()
        self.assertTrue(job.is_cancelled)
        self.assertFalse(job.is_active)
        self.assertEqual(job.effective_status, 'cancelled')

    def test_cancel_needs_valid_reason_code(self):
        job = self.make_job()
        with self.assertRaises(ValidationError):
            services.cancel_job(job, actor=self.user, reason='x', reason_code='bogus')

    def test_released_cancel_goes_through_change_request(self):
        job = self.make_job()
        services.submit_to_qc(job, actor=self.user)
        services.approve_qc(job, actor=self.user)
        services.release_for_production(job, actor=self.user)
        job.refresh_from_db()
        self.assertFalse(job.can_cancel_directly)
        change = services.request_cancellation(
            job, actor=self.user, reason='Customer withdrew', reason_code='customer_cancelled')
        with self.assertRaises(ValidationError):
            services.request_cancellation(job, actor=self.user, reason='again', reason_code='other')
        services.approve_change_request(change, actor=self.user)
        job.refresh_from_db()
        self.assertTrue(job.is_cancelled)

    def test_reopen_request_approve_returns_job_to_draft(self):
        job = self.make_job()
        services.submit_to_qc(job, actor=self.user)
        services.approve_qc(job, actor=self.user)
        change = services.request_reopen(job, actor=self.user, reason='Wrong qty')
        services.approve_change_request(change, actor=self.user)
        job.refresh_from_db()
        self.assertEqual(job.status, 'draft')
        self.assertFalse(job.issued_to_production)

    def test_reject_change_request_keeps_job_locked(self):
        job = self.make_job()
        services.submit_to_qc(job, actor=self.user)
        services.approve_qc(job, actor=self.user)
        change = services.request_reopen(job, actor=self.user, reason='Wrong qty')
        services.reject_change_request(change, actor=self.user, reason='No')
        job.refresh_from_db()
        self.assertEqual(job.status, 'qc_approved')

    def test_hold_archive_restore(self):
        job = self.make_job()
        with self.assertRaises(ValidationError):
            services.hold_job(job, actor=self.user, reason='')
        services.hold_job(job, actor=self.user, reason='Waiting for film')
        self.assertTrue(FlexoPlanningJob.objects.get(pk=job.pk).is_on_hold)
        services.release_hold(job, actor=self.user)
        services.archive_job(job, actor=self.user, reason='dup')
        self.assertFalse(FlexoPlanningJob.objects.get(pk=job.pk).is_active)
        services.restore_job(job, actor=self.user)
        self.assertTrue(FlexoPlanningJob.objects.get(pk=job.pk).is_active)


class SkuMasterWorkflowTests(FlexoFixtureMixin, TestCase):
    def test_missing_required_fields_reported(self):
        recipe = FlexoSkuRecipe(sku='X')
        missing = recipe.missing_required_fields()
        self.assertIn('Size H (mm)', missing)
        self.assertIn('Machine', missing)
        self.assertIn('Meter Per Roll', missing)
        self.assertNotIn('Pcs Per Roll', missing)   # optional: derived from piece length when blank

    def test_die_code_is_only_a_warning(self):
        recipe = self.make_recipe('NO-DIE', size_h_mm=Decimal('10'), pcs_per_roll=1000)
        self.assertEqual(recipe.missing_required_fields(), [])
        self.assertEqual(recipe.warning_fields(), ['Die Code / AWC'])


class FlexoViewSmokeTests(FlexoFixtureMixin, TestCase):
    def setUp(self):
        super().setUp()
        from core.models import Permission
        # A fresh test DB has no Permission rows (they come from seed_access_control),
        # and superusers receive every *existing* active permission.
        for code in (
            'action.view_jobcard', 'action.edit_jobcard', 'action.plan', 'action.view_approval_queue',
            'action.approve_qc', 'action.approve_pm', 'action.approve_sku_master_review',
            'action.cancel_planning_job',
        ):
            Permission.objects.get_or_create(code=code, defaults={'name': code})
        self.admin = User.objects.create_superuser('boss', 'boss@example.com', 'pw')
        self.client.force_login(self.admin)

    def test_pages_render(self):
        job = self.make_job()
        services.submit_to_qc(job, actor=self.user)
        card = job.job_card
        for name, args in [
            ('flexo:home', []), ('flexo:jobs', []), ('flexo:jobs_archived', []),
            ('flexo:job_detail', [job.pk]), ('flexo:approval_queue', []),
            ('flexo:sku_recipes', []), ('flexo:sku_recipes_approved', []),
            ('flexo:sku_recipe_add', []), ('flexo:sku_recipe_edit', [self.fl.pk]),
            ('flexo:manual_po_entry', []), ('flexo:job_card_detail', [card.pk]),
            ('flexo:job_card_print', [card.pk]),
        ]:
            response = self.client.get(reverse(name, args=args))
            self.assertEqual(response.status_code, 200, name)

    def test_manual_po_entry_creates_jobs(self):
        data = {
            'h-po_number': 'WO-77', 'h-po_received_date': '2026-09-01', 'h-destination': 'SITE-2', 'h-priority': '1',
            'l-TOTAL_FORMS': '5', 'l-INITIAL_FORMS': '0', 'l-MIN_NUM_FORMS': '0', 'l-MAX_NUM_FORMS': '1000',
            'l-0-sku': self.fl.pk, 'l-0-order_qty': '18360', 'l-0-stock_qty': '0',
            'l-1-sku': self.satin.pk, 'l-1-order_qty': '50000',
        }
        response = self.client.post(reverse('flexo:manual_po_entry'), data)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(FlexoPlanningJob.objects.filter(po_number='WO-77').count(), 2)

    def test_job_edit_locked_after_qc_approval(self):
        job = self.make_job()
        services.submit_to_qc(job, actor=self.user)
        services.approve_qc(job, actor=self.user)
        response = self.client.get(reverse('flexo:job_edit', args=[job.pk]))
        self.assertEqual(response.status_code, 302)

    def test_approval_queue_actions(self):
        job = self.make_job()
        services.submit_to_qc(job, actor=self.user)
        self.client.post(reverse('flexo:approval_queue'), {'action': 'approve_qc', 'planning_job_id': job.pk})
        job.refresh_from_db()
        self.assertEqual(job.status, 'qc_approved')
        self.client.post(reverse('flexo:approval_queue'),
                         {'action': 'release_for_production', 'planning_job_id': job.pk})
        job.refresh_from_db()
        self.assertEqual(job.status, 'released')

    def test_sku_approval_flow_and_draft_sync(self):
        recipe = self.make_recipe('FLOW-SKU', approved=False, size_h_mm=Decimal('10'), pcs_per_roll=1000)
        url = reverse('flexo:sku_recipe_edit', args=[recipe.pk])
        base = {
            'sku': recipe.sku, 'job_name': 'Flow', 'material': 'TTH',
            'size_w_mm': '25.4', 'size_h_mm': '10', 'ups': '2', 'roll_size': '316', 'meter_per_roll': '100', 'pcs_per_roll': '1000',
            'wastage_pct': '5', 'unit_type': 'piece', 'machine': self.machine.pk,
        }
        self.client.post(url, {**base, 'action': 'submit_review'})
        self.client.post(url, {**base, 'action': 'review'})
        self.client.post(url, {**base, 'action': 'approve'})
        recipe.refresh_from_db()
        self.assertEqual(recipe.master_data_status, 'approved')
        locked = self.client.post(url, {**base, 'job_name': 'Hacked'})
        self.assertEqual(locked.status_code, 200)
        recipe.refresh_from_db()
        self.assertEqual(recipe.job_name, 'Flow')


PO_PAYLOAD = {
    'po_number': 'PO-06-2026-172407', 'document_type': 'PO', 'po_date': '2026-06-29',
    'approval_date': '2026-06-29', 'department': 'PRODUCTION PRINTING',
    'delivery_location': 'UTOPIA INDUSTRIES SITE-3', 'grand_total': 9274250.0,
    'items': [
        {'line_no': 1, 'sku': 'BARCODELABEL4x4-WHITE', 'job_name': 'BARCODELABEL4x4-WHITE',
         'delivery_date': '2026-07-07', 'quantity': 138.0, 'unit': 'ROLL', 'unit_cost': 2050.0, 'remarks': ''},
        {'line_no': 3, 'sku': 'DIRECTTHERMALSTICKER2x3LABELROLL', 'job_name': 'DIRECT THERMAL STICKER 2x3 LABEL ROLL',
         'delivery_date': '2026-07-07', 'quantity': 1796.0, 'unit': 'ROLL', 'unit_cost': 2000.0, 'remarks': ''},
        {'line_no': 4, 'sku': 'UNKNOWN-NEW-SKU', 'job_name': 'UNKNOWN-NEW-SKU',
         'delivery_date': '2026-07-07', 'quantity': 10.0, 'unit': 'ROLL', 'unit_cost': 1.0, 'remarks': ''},
    ],
}
WO_PAYLOAD = {
    'po_number': 'WO-08-2026-07845', 'document_type': 'WO', 'po_date': '2026-08-28', 'approval_date': '2026-08-28',
    'department': 'PRODUCTION COMFORTER', 'delivery_location': 'UTOPIA INDUSTRIES SITE-6',
    'items': [{'line_no': 1, 'sku': 'LABELCAREB2B-BAREHOME-POLYESTERSATEEN-COMFORTER',
               'job_name': 'LABELCAREB2B-BAREHOME-POLYESTERSATEEN-COMFORTER', 'delivery_date': '2026-09-07',
               'quantity': 1000.0, 'unit': 'PIECE', 'unit_cost': 1.24,
               'remarks': '1.5X5.5 INCHES / Remarks: Material WO item'}],
}


class PoIntakeTests(FlexoFixtureMixin, TestCase):
    def setUp(self):
        super().setUp()
        from core.models import Department
        Department.objects.get_or_create(name='PRODUCTION PRINTING')
        self.label = self.make_recipe('BARCODELABEL4x4-WHITE', size_h_mm=Decimal('102'), unit_type='roll',
                                      ups=Decimal('3'), meter_per_roll=Decimal('106'))
        # matched through the job name, not the SKU (the PO SKU token has no spaces)
        self.thermal = self.make_recipe('DIRECT THERMAL STICKER 2x3 LABEL ROLL', size_h_mm=Decimal('50.8'),
                                        unit_type='roll', ups=Decimal('4'), meter_per_roll=Decimal('110'))

    def doc(self, payload=PO_PAYLOAD):
        from .po_intake import save_extracted_document
        return save_extracted_document(payload, actor=self.user)

    def test_lines_are_matched_to_sku_masters_ignoring_spaces_and_case(self):
        from .po_intake import analyze_lines
        states = {r['item']['line_no']: r['state'] for r in analyze_lines(self.doc())}
        self.assertEqual(states, {1: 'ready', 3: 'ready', 4: 'missing'})

    def test_unapproved_sku_blocks_the_line(self):
        from .po_intake import analyze_lines
        self.label.master_data_status = 'draft'
        self.label.save()
        states = {r['item']['line_no']: r['state'] for r in analyze_lines(self.doc())}
        self.assertEqual(states[1], 'unapproved')

    def test_jobs_created_in_po_unit_with_header_data(self):
        from .po_intake import analyze_lines, create_jobs_from_document
        doc = self.doc()
        created, skipped = create_jobs_from_document(doc, [r['index'] for r in analyze_lines(doc)], actor=self.user)
        self.assertEqual(len(created), 2)
        self.assertEqual(len(skipped), 1)                 # the SKU with no master
        job = next(j for j in created if j.sku == 'BARCODELABEL4x4-WHITE')
        self.assertEqual((job.po_number, job.unit_type, job.order_qty), ('PO-06-2026-172407', 'roll', 138))
        self.assertEqual(job.destination, 'SITE-3')
        self.assertEqual(str(job.po_received_date), '2026-06-29')
        self.assertEqual(str(job.delivery_date), '2026-07-07')
        self.assertEqual(job.department.name, 'PRODUCTION PRINTING')
        self.assertAlmostEqual(float(job.base_meters), 138 * 106 / 3, places=2)
        self.assertEqual(float(job.amount), 138 * 2050.0)

    def test_second_run_does_not_duplicate_jobs_and_draft_skus_are_created(self):
        from .po_intake import analyze_lines, create_draft_recipes, create_jobs_from_document
        doc = self.doc()
        create_jobs_from_document(doc, [r['index'] for r in analyze_lines(doc)], actor=self.user)
        again, _ = create_jobs_from_document(doc, [r['index'] for r in analyze_lines(doc)], actor=self.user)
        self.assertEqual(again, [])
        self.assertEqual(FlexoPlanningJob.objects.filter(po_number=doc.po_number).count(), 2)
        doc.refresh_from_db()
        self.assertEqual(doc.status, 'pending')           # the unknown SKU still has no job
        drafts = create_draft_recipes(doc, actor=self.user)
        self.assertEqual([d.sku for d in drafts], ['UNKNOWN-NEW-SKU'])
        self.assertEqual((drafts[0].master_data_status, drafts[0].unit_type), ('draft', 'roll'))

    def test_wo_piece_line_keeps_piece_unit_and_marks_document_processed(self):
        from .po_intake import create_jobs_from_document
        self.make_recipe('LABELCAREB2B-BAREHOME-POLYESTERSATEEN-COMFORTER', size_h_mm=Decimal('139.7'),
                         ups=Decimal('1'), meter_per_roll=Decimal('200'), pcs_per_roll=1000)
        doc = self.doc(WO_PAYLOAD)
        created, _ = create_jobs_from_document(doc, [0], actor=self.user)
        self.assertEqual((created[0].unit_type, created[0].order_qty), ('piece', 1000))
        self.assertEqual(float(created[0].order_roll_qty), 1.0)       # 1000 pcs / 1000 per roll
        doc.refresh_from_db()
        self.assertEqual(doc.status, 'processed')

    def test_payload_without_lines_or_number_is_rejected(self):
        from .po_intake import save_extracted_document
        with self.assertRaises(ValidationError):
            save_extracted_document({'po_number': 'PO-1', 'items': []}, actor=self.user)
        with self.assertRaises(ValidationError):
            save_extracted_document({'items': [{'sku': 'X'}]}, actor=self.user)

    def test_units_map_roll_and_piece_only(self):
        from .po_intake import map_unit
        self.assertEqual((map_unit('ROLL'), map_unit('piece'), map_unit('PCS'), map_unit('KG')),
                         ('roll', 'piece', 'piece', None))


class PoViewTests(FlexoFixtureMixin, TestCase):
    def setUp(self):
        super().setUp()
        from core.models import Permission
        for code in ('action.view_jobcard', 'action.edit_jobcard', 'action.plan'):
            Permission.objects.get_or_create(code=code, defaults={'name': code})
        self.client.force_login(User.objects.create_superuser('boss2', 'b@example.com', 'pw'))
        self.make_recipe('BARCODELABEL4x4-WHITE', size_h_mm=Decimal('102'), unit_type='roll', ups=Decimal('3'),
                         meter_per_roll=Decimal('106'))

    def test_upload_review_and_create(self):
        from unittest import mock
        from django.core.files.uploadedfile import SimpleUploadedFile
        from .models import FlexoPoDocument
        pdf = SimpleUploadedFile('PO-172407.pdf', b'%PDF-1.4 fake', content_type='application/pdf')
        with mock.patch('flexo.po_intake.extract_pdf', return_value=PO_PAYLOAD):
            response = self.client.post(reverse('flexo:po_upload'), {'po_file': pdf})
        self.assertEqual(response.status_code, 302)
        document = FlexoPoDocument.objects.get(po_number='PO-06-2026-172407')
        self.assertEqual(response.url, reverse('flexo:po_review', args=[document.pk]))
        page = self.client.get(response.url)
        self.assertContains(page, 'BARCODELABEL4x4-WHITE')
        self.assertContains(page, 'No SKU master')
        post = self.client.post(response.url, {'action': 'create_jobs', 'line': ['0']})
        self.assertEqual(post.status_code, 302)
        self.assertEqual(FlexoPlanningJob.objects.filter(po_number='PO-06-2026-172407').count(), 1)
        self.assertEqual(self.client.get(reverse('flexo:po_inbox')).status_code, 200)

    def test_non_pdf_and_unreadable_pdf_are_refused_politely(self):
        from unittest import mock
        from django.core.files.uploadedfile import SimpleUploadedFile
        txt = SimpleUploadedFile('x.txt', b'hello', content_type='text/plain')
        self.assertEqual(self.client.post(reverse('flexo:po_upload'), {'po_file': txt}).status_code, 200)
        pdf = SimpleUploadedFile('x.pdf', b'%PDF', content_type='application/pdf')
        with mock.patch('flexo.po_intake.extract_pdf', side_effect=RuntimeError('boom')):
            response = self.client.post(reverse('flexo:po_upload'), {'po_file': pdf})
        self.assertEqual(response.status_code, 200)


class SharedMasterGuardTests(FlexoFixtureMixin, TestCase):
    """Machines are one shared master. Flexo machines must stay out of offset-only
    screens/KPIs, offset presses must stay out of Flexo dropdowns, and a master that
    Flexo still uses cannot be deleted from Master Data."""

    def setUp(self):
        super().setUp()                                  # self.machine = Flexo-1 (flexo_printing)
        self.offset = Machine.objects.create(name='GTO Test', machine_type='offset_printing')
        self.digital = Machine.objects.create(name='Digital Test', machine_type='digital_printing')
        self.cutting = Machine.objects.create(name='Cutter Test', machine_type='cutting')

    def test_helpers_split_the_machine_list(self):
        from core.machine_scope import flexo_machines, offset_machines
        offset_names = set(offset_machines().values_list('name', flat=True))
        flexo_names = set(flexo_machines().values_list('name', flat=True))
        self.assertNotIn('Flexo-1', offset_names)
        self.assertTrue({'GTO Test', 'Digital Test', 'Cutter Test'} <= offset_names)
        self.assertEqual(flexo_names, {'Flexo-1', 'Cutter Test'})

    def test_floor_dashboard_utilisation_denominator_ignores_flexo_machines(self):
        from floor_dashboard.services import get_plant_overview
        with_flexo = get_plant_overview()['total_machines']
        Machine.objects.filter(machine_type='flexo_printing').delete()
        without_flexo = get_plant_overview()['total_machines']
        self.assertEqual(with_flexo, without_flexo)

    def test_planning_machine_picker_hides_flexo_machines(self):
        from planning.forms import SkuRecipeForm
        names = [value for value, _ in SkuRecipeForm().fields['machine_name'].widget.choices]
        self.assertIn('GTO Test', names)
        self.assertNotIn('Flexo-1', names)

    def test_flexo_forms_offer_only_flexo_cutting_and_generic_machines(self):
        from .forms import FlexoPlanningJobEditForm, FlexoSkuRecipeForm
        for form in (FlexoPlanningJobEditForm(), FlexoSkuRecipeForm()):
            field = form.fields.get('machine_name') or form.fields['machine']
            names = set(field.queryset.values_list('name', flat=True))
            self.assertTrue({'Flexo-1', 'Cutter Test'} <= names)
            self.assertFalse({'GTO Test', 'Digital Test'} & names)

    def test_saved_machine_stays_selectable_even_if_it_is_not_a_flexo_type(self):
        from .forms import FlexoPlanningJobEditForm
        job = self.make_job()
        job.machine_name = self.offset
        job.save()
        names = set(FlexoPlanningJobEditForm(instance=job).fields['machine_name'].queryset.values_list('name', flat=True))
        self.assertIn('GTO Test', names)

    def _master_delete(self, entity_type, record):
        from core.models import Permission
        Permission.objects.get_or_create(code='action.manage_masters', defaults={'name': 'Manage masters'})
        admin = User.objects.filter(username='boss3').first() or User.objects.create_superuser('boss3', 'b3@example.com', 'pw')
        self.client.force_login(admin)
        return self.client.post(reverse('master_data'), {
            'action': 'delete_master', 'entity_type': entity_type, 'record_id': record.pk})

    def test_machine_used_by_flexo_cannot_be_deleted_from_master_data(self):
        self.make_job()
        response = self._master_delete('machine', self.machine)
        self.assertEqual(response.status_code, 302)
        self.assertTrue(Machine.objects.filter(pk=self.machine.pk).exists())
        follow = self.client.get(reverse('master_data'))
        self.assertContains(follow, 'Linked Flexo records found')

    def test_unused_machine_can_still_be_deleted(self):
        self._master_delete('machine', self.cutting)
        self.assertFalse(Machine.objects.filter(pk=self.cutting.pk).exists())

    def test_department_and_operator_used_by_flexo_cannot_be_deleted(self):
        from core.models import Department, Operator
        department = Department.objects.create(name='FLEXO GUARD DEPT')
        job = self.make_job()
        job.department = department
        job.save()
        self._master_delete('department', department)
        self.assertTrue(Department.objects.filter(pk=department.pk).exists())
        operator = Operator.objects.create(name='Flexo Guard Operator')
        services.submit_to_qc(job, actor=self.user)
        job.job_card.printing_log_entries.create(operator=operator)
        self._master_delete('operator', operator)
        self.assertTrue(Operator.objects.filter(pk=operator.pk).exists())
