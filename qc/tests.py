from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from core.models import UserProfile
from planning.models import PlanningJob


class QcWorkflowCompatibilityTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(username='qc_manager', password='testpass123')
        profile, _created = UserProfile.objects.get_or_create(user=self.user)
        profile.role = 'manager'
        profile.save(update_fields=['role'])
        self.client.force_login(self.user)

    def _create_pending_qc_job(self, jc_number='JC-QC-001', sku='QC-SKU-001'):
        return PlanningJob.objects.create(
            jc_number=jc_number,
            po_number='PO-QC-001',
            sku=sku,
            job_name=f'{sku} Job',
            status='pending_qc',
            repeat_flag='New',
            material='Paper',
            color_spec='4+0',
            application='UV',
            ups=2,
            print_sheet_size='25x36',
            purchase_sheet_size='25x36',
            machine_name='Machine A',
            wastage_sheets=10,
            plate_set_no='PLATE-1',
            order_qty=1000,
        )

    def test_qc_reject_returns_job_to_draft(self):
        job = self._create_pending_qc_job(jc_number='JC-QC-REJECT')

        response = self.client.post(
            reverse('qc:planning_job_status_update', args=[job.id]),
            {'transition': 'reject_qc', 'reason': 'Specs mismatch'},
        )

        job.refresh_from_db()
        self.assertEqual(response.status_code, 302)
        self.assertEqual(job.status, 'draft')

    def test_qc_approve_requires_existing_job_card(self):
        job = self._create_pending_qc_job(jc_number='JC-QC-APPROVE')

        response = self.client.post(
            reverse('qc:planning_job_status_update', args=[job.id]),
            {'transition': 'approve_qc'},
        )

        job.refresh_from_db()
        self.assertEqual(response.status_code, 302)
        self.assertEqual(job.status, 'pending_qc')
