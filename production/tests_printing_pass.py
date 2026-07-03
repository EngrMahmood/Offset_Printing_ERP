from datetime import date

from django.contrib.auth import get_user_model
from django.core.exceptions import ValidationError
from django.test import TestCase

from core.models import JobCard, Machine, Production, Operator
from planning.models import PlanningJob
from production.printing_pass_helpers import (
    build_pass_tracking_info,
    get_job_card_pass_count,
    get_per_pass_impression_budget,
    get_suggested_print_pass,
    validate_print_pass_number,
)


class PrintingPassHelperTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(username='pass_tester', password='pass')
        self.machine = Machine.objects.create(name='Press Pass')
        self.operator = Operator.objects.create(name='Op A')
        self.planning_job = PlanningJob.objects.create(
            jc_number='JC-2PASS',
            sku='SKU-2P',
            job_name='2-pass job',
            order_qty=1000,
            ups=2,
            wastage_sheets=12,
            print_passes=2,
            status='draft',
            created_by=self.user,
        )
        self.job_card = JobCard.objects.create(
            job_card_no='JC-2PASS',
            SKU='SKU-2P',
            order_qty=1000,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=1024,
            total_sheet_quantity=512,
            total_colors=4,
            plate_set_no='PLATE-P',
            machine_name=self.machine,
            planning_job=self.planning_job,
        )

    def test_per_pass_budget_divides_total_impressions(self):
        self.assertEqual(get_per_pass_impression_budget(self.job_card), 512)

    def test_legacy_job_infers_two_passes_from_impressions_and_sheets(self):
        legacy_job = JobCard.objects.create(
            job_card_no='JC-LEGACY-2P',
            SKU='SKU-LEG',
            order_qty=1000,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=1024,
            total_sheet_quantity=512,
            total_colors=4,
            plate_set_no='PLATE-L',
            machine_name=self.machine,
        )
        from production.printing_pass_helpers import (
            infer_pass_count_from_impressions,
            passes_are_inferred,
        )
        self.assertEqual(infer_pass_count_from_impressions(legacy_job), 2)
        self.assertEqual(get_job_card_pass_count(legacy_job), 2)
        self.assertTrue(passes_are_inferred(legacy_job))
        tracking = build_pass_tracking_info(legacy_job)
        self.assertTrue(tracking['passes_inferred'])
        self.assertIn('inferred', tracking['legacy_notice'].lower())

    def test_legacy_single_pass_when_impressions_match_sheets(self):
        legacy_job = JobCard.objects.create(
            job_card_no='JC-LEGACY-1P',
            SKU='SKU-LEG1',
            order_qty=500,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=512,
            total_sheet_quantity=512,
            total_colors=4,
            plate_set_no='PLATE-L1',
            machine_name=self.machine,
        )
        from production.printing_pass_helpers import infer_pass_count_from_impressions
        self.assertEqual(infer_pass_count_from_impressions(legacy_job), 1)
        self.assertEqual(get_job_card_pass_count(legacy_job), 1)

    def test_inferred_two_pass_suggests_pass_two_after_pass_one_budget(self):
        legacy_job = JobCard.objects.create(
            job_card_no='JC-LEGACY-SUG',
            SKU='SKU-SUG',
            order_qty=1000,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=1024,
            total_sheet_quantity=512,
            total_colors=4,
            plate_set_no='PLATE-S',
            machine_name=self.machine,
        )
        Production.objects.create(
            job_card=legacy_job,
            entry_type='printing',
            print_pass_number=1,
            date=date(2026, 1, 2),
            shift='A',
            machine=self.machine,
            operator=self.operator,
            impressions=520,
            output_sheets=0,
            intermediate_pass=True,
            planned_time=30,
            run_time=30,
        )
        self.assertEqual(get_suggested_print_pass(legacy_job), 2)

    def test_multiple_pass_one_entries_allowed(self):
        Production.objects.create(
            job_card=self.job_card,
            entry_type='printing',
            print_pass_number=1,
            date=date(2026, 1, 2),
            shift='A',
            machine=self.machine,
            operator=self.operator,
            impressions=300,
            output_sheets=0,
            intermediate_pass=True,
            planned_time=30,
            run_time=30,
        )
        Production.objects.create(
            job_card=self.job_card,
            entry_type='printing',
            print_pass_number=1,
            date=date(2026, 1, 4),
            shift='B',
            machine=self.machine,
            operator=self.operator,
            impressions=220,
            output_sheets=0,
            intermediate_pass=True,
            planned_time=20,
            run_time=20,
        )
        tracking = build_pass_tracking_info(self.job_card)
        self.assertEqual(tracking['pass_usage'][1], 520)
        self.assertEqual(tracking['pass_usage'][2], 0)
        self.assertEqual(get_suggested_print_pass(self.job_card), 2)

    def test_cannot_start_pass_two_without_pass_one(self):
        with self.assertRaises(ValueError):
            validate_print_pass_number(self.job_card, 2)

    def test_final_pass_requires_good_sheets(self):
        Production.objects.create(
            job_card=self.job_card,
            entry_type='printing',
            print_pass_number=1,
            date=date(2026, 1, 2),
            shift='A',
            machine=self.machine,
            operator=self.operator,
            impressions=500,
            output_sheets=0,
            intermediate_pass=True,
            planned_time=30,
            run_time=30,
        )
        record = Production(
            job_card=self.job_card,
            entry_type='printing',
            print_pass_number=2,
            date=date(2026, 1, 3),
            shift='A',
            machine=self.machine,
            operator=self.operator,
            impressions=500,
            output_sheets=0,
            intermediate_pass=False,
            planned_time=30,
            run_time=30,
        )
        with self.assertRaises(ValidationError):
            record.full_clean()

    def test_final_pass_with_good_sheets_succeeds(self):
        Production.objects.create(
            job_card=self.job_card,
            entry_type='printing',
            print_pass_number=1,
            date=date(2026, 1, 2),
            shift='A',
            machine=self.machine,
            operator=self.operator,
            impressions=500,
            output_sheets=0,
            intermediate_pass=True,
            planned_time=30,
            run_time=30,
        )
        record = Production(
            job_card=self.job_card,
            entry_type='printing',
            print_pass_number=2,
            date=date(2026, 1, 3),
            shift='A',
            machine=self.machine,
            operator=self.operator,
            impressions=500,
            output_sheets=500,
            intermediate_pass=False,
            planned_time=30,
            run_time=30,
        )
        record.full_clean()
        record.save()
        self.assertEqual(self.job_card.total_printed_pcs, 5000)

    def test_legacy_sm74_four_colors_is_single_pass(self):
        sm74 = Machine.objects.create(name='SM74')
        legacy_job = JobCard.objects.create(
            job_card_no='JC-SM74-4C',
            SKU='SKU-SM74',
            order_qty=1000,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=512,
            total_sheet_quantity=512,
            total_colors=4,
            plate_set_no='PLATE-SM74',
            machine_name=sm74,
        )
        self.assertEqual(get_job_card_pass_count(legacy_job), 1)
        tracking = build_pass_tracking_info(legacy_job)
        self.assertEqual(tracking['pass_inference_reason'], 'SM74 with 4 colors')
        self.assertIn('SM74', tracking['legacy_notice'])

    def test_impressions_take_priority_over_sm74_rule(self):
        sm74 = Machine.objects.create(name='SM74')
        legacy_job = JobCard.objects.create(
            job_card_no='JC-SM74-2P',
            SKU='SKU-SM74-2P',
            order_qty=1000,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=1024,
            total_sheet_quantity=512,
            total_colors=4,
            plate_set_no='PLATE-SM74-2P',
            machine_name=sm74,
        )
        self.assertEqual(get_job_card_pass_count(legacy_job), 2)
        tracking = build_pass_tracking_info(legacy_job)
        self.assertEqual(tracking['pass_inference_source'], 'impressions')

    def test_legacy_gto_four_colors_is_two_pass(self):
        gto = Machine.objects.create(name='Heidelberg GTO 52')
        legacy_job = JobCard.objects.create(
            job_card_no='JC-GTO-2P',
            SKU='SKU-GTO',
            order_qty=500,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=512,
            total_sheet_quantity=512,
            total_colors=4,
            plate_set_no='PLATE-GTO',
            machine_name=gto,
        )
        self.assertEqual(get_job_card_pass_count(legacy_job), 2)
        tracking = build_pass_tracking_info(legacy_job)
        self.assertTrue(tracking['passes_inferred'])
        self.assertEqual(tracking['pass_inference_reason'], 'GTO with 3-4 colors')

    def test_legacy_gto_two_colors_is_single_pass(self):
        gto = Machine.objects.create(name='Heidelberg GTO 52')
        legacy_job = JobCard.objects.create(
            job_card_no='JC-GTO-1P',
            SKU='SKU-GTO-1P',
            order_qty=500,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=512,
            total_sheet_quantity=512,
            total_colors=2,
            plate_set_no='PLATE-GTO-1P',
            machine_name=gto,
        )
        self.assertEqual(get_job_card_pass_count(legacy_job), 1)
        tracking = build_pass_tracking_info(legacy_job)
        self.assertEqual(tracking['pass_inference_reason'], 'GTO with 1-2 colors')

    def test_legacy_colour_one_plus_one_is_two_pass(self):
        legacy_job = JobCard.objects.create(
            job_card_no='JC-1P1',
            SKU='SKU-1P1',
            order_qty=500,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=512,
            total_sheet_quantity=512,
            colour='1+1',
            total_colors=2,
            plate_set_no='PLATE-1P1',
            machine_name=self.machine,
        )
        self.assertEqual(get_job_card_pass_count(legacy_job), 2)
        tracking = build_pass_tracking_info(legacy_job)
        self.assertEqual(tracking['pass_inference_reason'], 'colour 1+1')

    def test_impressions_take_priority_over_one_plus_one_colour(self):
        legacy_job = JobCard.objects.create(
            job_card_no='JC-1P1-IMP',
            SKU='SKU-1P1-IMP',
            order_qty=500,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=1024,
            total_sheet_quantity=512,
            colour='1+1',
            total_colors=2,
            plate_set_no='PLATE-1P1-IMP',
            machine_name=self.machine,
        )
        self.assertEqual(get_job_card_pass_count(legacy_job), 2)
        tracking = build_pass_tracking_info(legacy_job)
        self.assertEqual(tracking['pass_inference_source'], 'impressions')
