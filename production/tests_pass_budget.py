from django.test import TestCase

from core.models import Production
from planning.models import MergeGroup, MergeGroupItem
from planning.tests_merge import make_card, make_job
from production.printing_entry_helpers import build_printing_job_card_maps
from production.printing_pass_helpers import (
    build_pass_tracking_info, get_per_pass_impression_budget, validate_print_pass_number,
)


class MergedLeadPassBudgetTests(TestCase):
    """JC-10-26-PP-4273 case: a 2-pass lead of a 5,000-sheet (+5 wastage) combined run, standalone plan only 1,563 sheets."""

    def setUp(self):
        self.job = make_job('LEAD1', 25000, ups=16, print_passes=2)
        self.card = make_card(self.job)
        self.card.total_sheet_quantity = 1563
        self.card.total_impressions_required = 3126
        self.card.production_tolerance_percent = 5
        self.card.save()
        self.group = MergeGroup.objects.create(code='MG-T', status='artwork_ready', total_sheet_ups=16, run_sheets=5000,
                                               combined_wastage_sheets=5, lead_job=self.job)
        MergeGroupItem.objects.create(merge_group=self.group, planning_job=self.job, allocated_ups=5, is_lead=True)

    def entry(self, pass_no, impressions, **extra):
        return Production.objects.create(job_card=self.card, entry_type='printing', date='2026-11-11', shift='B',
                                         print_pass_number=pass_no, impressions=impressions, intermediate_pass=pass_no < 2, **extra)

    def test_per_pass_budget_is_the_combined_run_not_the_standalone_plan(self):
        self.assertEqual(get_per_pass_impression_budget(self.card), 5005)

    def test_pass_one_with_3000_is_not_complete(self):
        self.entry(1, 3000)
        info = build_pass_tracking_info(self.card)
        row = info['pass_rows'][0]
        self.assertEqual((row['used'], row['budget'], row['remaining']), (3000, 5005, 2005))
        self.assertEqual(info['total_impressions_allowed'], int(round(5005 * 2 * 1.05)))
        self.assertEqual(info['suggested_pass'], 1)                         # keep logging pass 1
        with self.assertRaises(ValueError):                                 # (pass 1 is still open, so this is not "already complete")
            validate_print_pass_number(self.card, 2, new_impressions=5000)  # but pass 2 may not jump ahead of pass 1

    def test_pass_one_at_its_budget_is_complete(self):
        self.entry(1, 5005)
        with self.assertRaisesMessage(ValueError, 'already complete'):
            validate_print_pass_number(self.card, 1)

    def test_pass_two_cannot_exceed_what_pass_one_recorded(self):
        self.entry(1, 3000)
        with self.assertRaisesMessage(ValueError, 'Pass 1 only has 3,000 recorded'):
            validate_print_pass_number(self.card, 2, new_impressions=5000)
        validate_print_pass_number(self.card, 2, new_impressions=3000)       # equal to pass 1 is fine
        validate_print_pass_number(self.card, 2, new_impressions=3150)       # within the 5% tolerance

    def test_edits_are_never_blocked_by_the_pass_order_rule(self):
        self.entry(1, 3000)
        pass_two = self.entry(2, 5000)                       # the existing (legacy) pass-2 row stays correctable
        validate_print_pass_number(self.card, 2, exclude_production_id=pass_two.pk, new_impressions=5000)


class OrdinaryTwoPassJobTests(TestCase):
    def test_a_normal_job_is_unchanged_but_also_follows_the_pass_order_rule(self):
        job = make_job('NORM1', 1000, ups=4, print_passes=2)
        card = make_card(job)
        card.total_sheet_quantity = 260
        card.total_impressions_required = 520
        card.save()
        self.assertEqual(get_per_pass_impression_budget(card), 260)
        Production.objects.create(job_card=card, entry_type='printing', date='2026-11-11', shift='A', print_pass_number=1,
                                  impressions=260, intermediate_pass=True)
        validate_print_pass_number(card, 2, new_impressions=260)
        with self.assertRaises(ValueError):
            validate_print_pass_number(card, 2, new_impressions=300)


class MergedLeadSheetTargetTests(MergedLeadPassBudgetTests):
    def test_good_sheet_target_and_sheet_budget_are_the_combined_run(self):
        self.entry(1, 5000, waste_sheets=0, output_sheets=0)
        self.entry(2, 5000, output_sheets=4980, waste_sheets=20)
        info = build_pass_tracking_info(self.card)
        self.assertEqual(info['good_sheets_target'], 5000)                       # not the standalone ~1,562
        self.assertEqual(info['sheets_allowed'], int(round(5005 * 1.05)))        # 5,255
        self.assertEqual(info['sheets_used'], 5000)                              # good + waste actually run
        self.assertEqual(info['sheets_remaining'], 255)
