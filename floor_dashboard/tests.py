from datetime import date, timedelta

from django.contrib.auth import get_user_model
from django.test import TestCase

from core.models import JobCard, Machine, Operator, Production, Sorter
from planning.models import PlanningJob
from . import services
from .models import DailyTarget

TEST_DATE = date(2026, 2, 1)


class FloorDashboardServicesTests(TestCase):
    def setUp(self):
        self.machine = Machine.objects.create(name='FD Test Machine', standard_impressions_per_hour=4000)
        self.operator = Operator.objects.create(name='FD Test Operator')
        self.sorter = Sorter.objects.create(name='FD Test Sorter')

        self.planning_job = PlanningJob.objects.create(
            jc_number='JC-FD-0001', order_qty=1000, status='released',
            plan_date=TEST_DATE, plan_month='February 2026', sku='SKU-FD-1',
        )
        self.job_card = JobCard.objects.create(
            job_card_no='JC-FD-0001', planning_job=self.planning_job, order_qty=1000,
            SKU='SKU-FD-1', ups=1, is_print_job=True, total_sheet_quantity=1000,
            total_colors=4, status='in_production', po_date=TEST_DATE,
            plate_set_no='PLATE-FD-1', machine_name=self.machine,
            total_impressions_required=1000,
        )
        Production.objects.create(
            job_card=self.job_card, entry_type='printing', date=TEST_DATE, shift='A',
            machine=self.machine, operator=self.operator,
            output_sheets=900, waste_sheets=50, impressions=950,
            planned_time=60, run_time=60,
        )
        Production.objects.create(
            job_card=self.job_card, entry_type='packing', date=TEST_DATE, shift='A',
            sorter=self.sorter, packing_qty=800, sorting_waste_qty=20,
        )

    def test_plant_overview_reflects_todays_activity(self):
        data = services.get_plant_overview(TEST_DATE)
        self.assertEqual(data['printed_pcs'], 900)
        self.assertEqual(data['packed_pcs'], 800)
        self.assertTrue(data['target_is_estimated'])
        self.assertEqual(data['target_qty'], 1000)

    def test_daily_target_overrides_estimate(self):
        DailyTarget.objects.create(date=TEST_DATE, target_qty=5000)
        data = services.get_plant_overview(TEST_DATE)
        self.assertFalse(data['target_is_estimated'])
        self.assertEqual(data['target_qty'], 5000)

    def test_printing_performance_leaderboards(self):
        data = services.get_printing_performance(TEST_DATE)
        self.assertEqual(data['output_sheets_total'], 900)
        self.assertEqual(data['waste_sheets_total'], 50)
        self.assertEqual(len(data['top_machines']), 1)
        self.assertEqual(data['top_machines'][0]['name'], 'FD Test Machine')
        self.assertEqual(len(data['top_operators']), 1)
        self.assertEqual(data['top_operators'][0]['name'], 'FD Test Operator')

    def test_packing_performance(self):
        data = services.get_packing_performance(TEST_DATE)
        self.assertEqual(data['packed_pcs'], 800)
        self.assertEqual(data['sorting_waste_pcs'], 20)
        self.assertEqual(len(data['top_sorters']), 1)
        self.assertEqual(data['top_sorters'][0]['name'], 'FD Test Sorter')

    def test_shift_comparison_has_both_shifts(self):
        data = services.get_shift_comparison(TEST_DATE)
        shift_codes = {s['shift'] for s in data['shifts']}
        self.assertEqual(shift_codes, {'A', 'B'})
        self.assertEqual(data['winning_shift'], 'Shift A')

    def test_machine_and_operator_leaderboards(self):
        machines = services.get_machine_leaderboard(TEST_DATE)
        self.assertEqual(machines['machine_of_the_day'], 'FD Test Machine')

        operators = services.get_operator_leaderboard(TEST_DATE)
        self.assertEqual(operators['operators'][0]['name'], 'FD Test Operator')
        self.assertEqual(operators['sorters'][0]['name'], 'FD Test Sorter')

    def test_wastage_quality(self):
        data = services.get_wastage_quality(TEST_DATE)
        self.assertEqual(data['printing_waste_pcs'], 50)
        self.assertEqual(data['sorting_waste_pcs'], 20)
        self.assertEqual(data['total_process_wastage_pcs'], 70)

    def test_dashboard_data_bundles_all_screens(self):
        data = services.get_dashboard_data(TEST_DATE)
        expected_keys = {
            'plant_overview', 'printing_performance', 'packing_performance',
            'dispatch_performance', 'shift_comparison', 'machine_leaderboard',
            'operator_leaderboard', 'wastage_quality', 'target_achievement',
            'recognition',
        }
        self.assertTrue(expected_keys.issubset(data.keys()))


class PeriodSummaryTests(TestCase):
    """REFERENCE_TODAY is a Wednesday well into its month, so 'yesterday',
    'earlier this week' (Monday), and 'earlier this month' (the 1st) are
    all distinct dates that don't collide with each other."""

    REFERENCE_TODAY = date(2026, 3, 18)  # Wednesday

    def setUp(self):
        assert self.REFERENCE_TODAY.weekday() == 2  # Wednesday
        self.monday = self.REFERENCE_TODAY - timedelta(days=self.REFERENCE_TODAY.weekday())
        self.yesterday = self.REFERENCE_TODAY - timedelta(days=1)
        self.first_of_month = self.REFERENCE_TODAY.replace(day=1)
        self.last_month_day = self.first_of_month - timedelta(days=1)

        self.sorter = Sorter.objects.create(name='Period Test Sorter')
        machine = Machine.objects.create(name='Period Test Machine')
        planning_job = PlanningJob.objects.create(
            jc_number='JC-FD-PERIOD', order_qty=1000, status='released',
            plan_date=self.REFERENCE_TODAY, plan_month='March 2026', sku='SKU-FD-PERIOD',
        )
        self.job_card = JobCard.objects.create(
            job_card_no='JC-FD-PERIOD', planning_job=planning_job, order_qty=1000,
            SKU='SKU-FD-PERIOD', ups=1, is_print_job=True, total_sheet_quantity=1000,
            total_colors=4, status='in_production', po_date=self.REFERENCE_TODAY,
            plate_set_no='PLATE-FD-PERIOD', total_impressions_required=1000,
            machine_name=machine,
        )

        Production.objects.create(
            job_card=self.job_card, entry_type='printing', date=self.REFERENCE_TODAY, shift='A',
            machine=machine, output_sheets=300, waste_sheets=0, impressions=300,
            planned_time=60, run_time=60,
        )
        self._pack(self.REFERENCE_TODAY, 100)
        self._pack(self.yesterday, 50)
        self._pack(self.monday, 30)
        self._pack(self.first_of_month, 20)
        self._pack(self.last_month_day, 10)

    def _pack(self, entry_date, qty):
        Production.objects.create(
            job_card=self.job_card, entry_type='packing', date=entry_date, shift='A',
            sorter=self.sorter, packing_qty=qty, sorting_waste_qty=0,
        )

    def _period(self, data, key):
        return next(p for p in data['periods'] if p['key'] == key)

    def test_today_only_includes_todays_entry(self):
        data = services.get_period_summary(self.REFERENCE_TODAY)
        self.assertEqual(self._period(data, 'today')['packed_pcs'], 100)

    def test_yesterday_only_includes_yesterdays_entry(self):
        data = services.get_period_summary(self.REFERENCE_TODAY)
        self.assertEqual(self._period(data, 'yesterday')['packed_pcs'], 50)

    def test_week_includes_monday_through_today_but_not_last_month(self):
        data = services.get_period_summary(self.REFERENCE_TODAY)
        # today (100) + yesterday (50) + monday (30) = 180; last month's 10 excluded.
        self.assertEqual(self._period(data, 'week')['packed_pcs'], 180)

    def test_month_includes_first_of_month_but_not_last_month(self):
        data = services.get_period_summary(self.REFERENCE_TODAY)
        # today (100) + yesterday (50) + monday (30) + 1st (20) = 200; last month's 10 excluded.
        self.assertEqual(self._period(data, 'month')['packed_pcs'], 200)

    def test_dashboard_data_includes_period_summary(self):
        data = services.get_dashboard_data(self.REFERENCE_TODAY)
        self.assertIn('period_summary', data)
        self.assertEqual(len(data['period_summary']['periods']), 4)


class FloorDashboardViewTests(TestCase):
    def test_dashboard_view_renders(self):
        response = self.client.get('/floor-dashboard/')
        self.assertEqual(response.status_code, 200)

    def test_dashboard_data_api_returns_json(self):
        response = self.client.get('/floor-dashboard/api/data/')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response['Content-Type'], 'application/json')
