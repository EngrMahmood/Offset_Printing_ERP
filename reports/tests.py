from django.test import TestCase
from django.urls import reverse
from django.contrib.auth import get_user_model

from core.models import UserProfile


class ReportsAppTests(TestCase):
    def setUp(self):
        User = get_user_model()
        self.user = User.objects.create_user(username='report_user', password='pass12345')
        profile = UserProfile.objects.get(user=self.user)
        profile.role = 'manager'
        profile.save(update_fields=['role'])

    def test_reports_home_loads(self):
        response = self.client.get(reverse('reports:home'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Available Reports')
        self.assertContains(response, 'Machine Planning')

    def test_reports_api_list_loads(self):
        response = self.client.get(reverse('reports:reports_api:list_reports'))
        self.assertEqual(response.status_code, 200)
        self.assertJSONEqual(
            response.content,
            {
                'ok': True,
                'count': len(response.json().get('items', [])),
                'items': response.json().get('items', []),
            },
        )

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

    def test_report_run_api_loads(self):
        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.json()['ok'])

    def test_report_export_csv_loads(self):
        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:export_report', args=['machine-planning']), {'type': 'csv'})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response['Content-Type'].split(';')[0], 'text/csv')
