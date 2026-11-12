from django.contrib.auth import get_user_model
from django.core.exceptions import ValidationError
from django.test import TestCase

from core.models import Production, Sorter

from .models import PlanningJob, SkuRecipe
from .services import create_reprint_job, reprint_blockers, reprint_remaining_qty
from .sku_classification import classify_po_line, refresh_stale_repeat_flag
from .tests_merge import make_card, make_job


class ReprintTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user('planner', password='x')
        self.job = make_job('JC-11-26-PP-5000', 1000, repeat_flag='New', po_number='WO-1', department='KNITS', destination='SITE-2')
        self.card = make_card(self.job)
        self.card.status = 'in_production'
        self.card.is_print_job = False                       # no print-limit on the packing entries below
        self.card.save(update_fields=['status', 'is_print_job'])
        self.sorter = Sorter.objects.create(name='S1')

    def _pack(self, qty):
        Production.objects.create(job_card=self.card, entry_type='packing', date='2026-11-01', shift='A',
                                  packing_qty=qty, sorting_waste_qty=0, sorter=self.sorter)

    def test_remaining_is_order_minus_packed(self):
        self._pack(600)
        self.assertEqual(reprint_remaining_qty(self.job), 400)

    def test_reprint_makes_a_new_repeat_job_and_short_closes_the_original(self):
        self._pack(600)
        new = create_reprint_job(self.job, actor=self.user, reason='High wastage on press')
        self.assertEqual(new.jc_number, 'JC-11-26-PP-5000.1')
        self.assertEqual(new.order_qty, 400)
        self.assertEqual(new.repeat_flag, 'Repeat')
        self.assertEqual(new.status, 'draft')
        self.assertEqual(new.reprint_of_id, self.job.id)
        self.assertEqual((new.po_number, new.sku, new.department, new.destination),
                         (self.job.po_number, self.job.sku, 'KNITS', 'SITE-2'))
        self.assertIn('REPRINT of JC-11-26-PP-5000', new.requirement)
        self.card.refresh_from_db()
        self.assertEqual(self.card.workflow_status, 'closed')
        self.assertEqual(self.card.short_close_closed_qty, 400)
        self.assertIn(new.jc_number, self.card.short_close_close_reason)

    def test_reprint_needs_a_reason_and_cannot_be_done_twice(self):
        self._pack(600)
        with self.assertRaises(ValidationError):
            create_reprint_job(self.job, actor=self.user, reason='  ')
        create_reprint_job(self.job, actor=self.user, reason='High wastage')
        self.assertTrue(any('Already reprinted' in b for b in reprint_blockers(self.job)))
        with self.assertRaises(ValidationError):
            create_reprint_job(self.job, actor=self.user, reason='again')

    def test_quantity_cannot_exceed_what_is_missing(self):
        self._pack(600)
        with self.assertRaises(ValidationError):
            create_reprint_job(self.job, actor=self.user, reason='x', qty=401)
        new = create_reprint_job(self.job, actor=self.user, reason='x', qty=250)
        self.assertEqual(new.order_qty, 250)

    def test_nothing_to_reprint_when_fully_packed_or_not_started(self):
        self._pack(1000)
        self.assertTrue(reprint_blockers(self.job))
        other = make_job('JC-11-26-PP-5001', 500)
        other_card = make_card(other)                       # released, not yet in production
        self.assertTrue(any('in production' in b for b in reprint_blockers(other)))


class StaleRepeatFlagTests(TestCase):
    def test_a_draft_job_flagged_repeat_with_no_history_goes_back_to_new(self):
        job = make_job('JC-11-26-PP-6000', 1000, repeat_flag='Repeat', po_number='WO-9', sku='SKU-STALE')
        SkuRecipe.objects.create(sku='SKU-STALE', master_data_status='draft')
        self.assertTrue(refresh_stale_repeat_flag(job))
        job.refresh_from_db()
        self.assertEqual(job.repeat_flag, 'New')

    def test_a_job_with_an_approved_recipe_or_in_qc_keeps_its_flag(self):
        approved = make_job('JC-11-26-PP-6001', 100, repeat_flag='Repeat', po_number='WO-9', sku='SKU-OK')
        SkuRecipe.objects.create(sku='SKU-OK', master_data_status='approved')
        self.assertFalse(refresh_stale_repeat_flag(approved))
        in_qc = make_job('JC-11-26-PP-6002', 100, repeat_flag='Repeat', po_number='WO-9', sku='SKU-QC', status='pending_qc')
        self.assertFalse(refresh_stale_repeat_flag(in_qc))

    def test_a_terminated_earlier_wo_is_not_evidence_of_production(self):
        # the SKU appears on another, cancelled job only: it must not make a new WO a Repeat
        make_job('JC-11-26-PP-6003', 100, sku='SKU-TERM', po_number='WO-OLD', is_cancelled=True, is_active=False, status='cancelled')
        label, _ = classify_po_line('SKU-TERM', 'WO-NEW')
        self.assertEqual(label, 'New')
