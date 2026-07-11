from datetime import date

from django.contrib.auth import get_user_model
from django.db import connection
from django.test import TestCase
from django.test.utils import CaptureQueriesContext
from django.urls import reverse

from core.models import JobCard, Machine, UserProfile
from planning.models import PlanningJob
from printing_plates.models import PlateRequest
from production.released_jobs import PAGE_SIZE, released_print_jobs_queryset


class ReleasedJobsPerformanceTests(TestCase):
    def setUp(self):
        User = get_user_model()
        self.user = User.objects.create_user(username='prod_released', password='pass')
        profile, _ = UserProfile.objects.get_or_create(user=self.user)
        profile.role = 'production'
        profile.save(update_fields=['role'])
        self.client.force_login(self.user)
        self.machine = Machine.objects.create(name='GTO Released Test')

        for i in range(PAGE_SIZE + 5):
            planning_job = PlanningJob.objects.create(
                jc_number=f'JC-REL-{i:03d}',
                po_number=f'PO-REL-{i:03d}',
                sku=f'SKU-REL-{i:03d}',
                status='released',
                planning_stage='planning_done',
                plate_set_no=f'SET-{i:03d}',
                machine_name=self.machine.name,
            )
            JobCard.objects.create(
                job_card_no=f'JC-REL-{i:03d}',
                planning_job=planning_job,
                SKU=f'SKU-REL-{i:03d}',
                order_qty=1000,
                is_print_job=True,
                status='released',
                po_date=date(2026, 1, 1),
                total_impressions_required=100,
                total_sheet_quantity=100,
                total_colors=4,
                plate_set_no=f'SET-{i:03d}',
                machine_name=self.machine,
            )

        waiting_card = JobCard.objects.get(job_card_no='JC-REL-000')
        PlateRequest.objects.create(
            job_card=waiting_card,
            planning_job=waiting_card.planning_job,
            source=PlateRequest.SOURCE_REPLACEMENT,
            replacement_reason=PlateRequest.REASON_DAMAGED_DURING_RUN,
            status=PlateRequest.STATUS_SENT,
            damaged_colors='C',
            set_no='SET-000',
        )

    def test_pagination_returns_page_size(self):
        response = self.client.get(reverse('released_jobs'))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(len(response.context['rows']), PAGE_SIZE)
        self.assertTrue(response.context['page_obj'].has_next())

    def test_page_two_preserves_filters(self):
        response = self.client.get(
            reverse('released_jobs'),
            {'plate_status': 'ready', 'page': '2'},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context['plate_status'], 'ready')
        self.assertIn('plate_status=ready', response.context['filter_query'])
        for row in response.context['rows']:
            self.assertFalse(row['waiting_for_plate'])

    def test_plate_status_waiting_filter_uses_annotation(self):
        response = self.client.get(reverse('released_jobs'), {'plate_status': 'waiting'})
        self.assertEqual(response.status_code, 200)
        self.assertGreaterEqual(response.context['waiting_count'], 1)
        self.assertTrue(all(row['waiting_for_plate'] for row in response.context['rows']))

    def test_list_query_count_stays_bounded(self):
        with CaptureQueriesContext(connection) as ctx:
            response = self.client.get(reverse('released_jobs'))
        self.assertEqual(response.status_code, 200)
        # Auth/session + annotated page query + recipe/history batch — not per-row N+1.
        self.assertLess(len(ctx), 35)

    def test_queryset_annotations_present(self):
        job = released_print_jobs_queryset().first()
        self.assertIsNotNone(job)
        self.assertTrue(hasattr(job, 'waiting_for_plate'))
        self.assertTrue(hasattr(job, 'remake_count'))
        self.assertTrue(hasattr(job, 'has_printing_entry'))
