from django.test import TestCase
from django.urls import reverse


class ReportsAppTests(TestCase):
    def test_reports_home_loads(self):
        response = self.client.get(reverse('reports:home'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Pending Jobs')
        self.assertContains(response, 'Balance Impressions')
        self.assertContains(response, 'Balance Dispatch')
        self.assertContains(response, 'Machine Planning')

    def test_reports_dashboard_period_filters(self):
        response = self.client.get(reverse('reports:home'), {'period': 'today'})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Today')

        response = self.client.get(reverse('reports:home'), {'period': 'week'})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'This Week')

        response = self.client.get(
            reverse('reports:home'),
            {'date_from': '2026-01-01', 'date_to': '2026-01-31'},
        )
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Jan')

    def test_machine_planning_report_loads(self):
        response = self.client.get(reverse('reports:detail', args=['machine-planning']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Machine Planning')

    def test_job_planning_report_loads(self):
        response = self.client.get(reverse('reports:detail', args=['job-planning']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Job Planning')

    def test_qc_approvals_report_loads(self):
        response = self.client.get(reverse('reports:detail', args=['qc-approvals']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'QC Approvals')

    def test_dispatch_tracking_report_loads(self):
        response = self.client.get(reverse('reports:detail', args=['dispatch-tracking']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Dispatch Tracking')

    def test_report_detail_404_for_invalid_slug(self):
        response = self.client.get(reverse('reports:detail', args=['invalid-report']))
        self.assertEqual(response.status_code, 404)
