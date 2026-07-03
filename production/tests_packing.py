from datetime import date

from django.contrib.auth import get_user_model
from django.core.exceptions import ValidationError
from django.test import Client, TestCase
from django.urls import reverse

from core.models import JobCard, Machine, Production, Sorter
from production.packing_entry import _packing_eligible_job_cards_queryset


class PackingProductionValidationTests(TestCase):
    def setUp(self):
        self.machine = Machine.objects.create(name='Press 1')
        self.sorter = Sorter.objects.create(name='Sorter A')
        self.job_card = JobCard.objects.create(
            job_card_no='JC-PACK-001',
            SKU='SKU-PACK',
            order_qty=1000,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=100,
            total_sheet_quantity=100,
            total_colors=4,
            plate_set_no='PLATE-1',
            machine_name=self.machine,
        )
        Production.objects.create(
            job_card=self.job_card,
            entry_type='printing',
            date=date(2026, 1, 2),
            shift='A',
            machine=self.machine,
            output_sheets=100,
            impressions=100,
            planned_time=60,
            run_time=60,
        )

    def test_packing_within_printed_limit_succeeds(self):
        record = Production(
            job_card=self.job_card,
            entry_type='packing',
            date=date(2026, 1, 3),
            shift='A',
            packing_qty=800,
            sorting_waste_qty=200,
            sorter=self.sorter,
        )
        record.full_clean()
        record.save()
        self.assertEqual(self.job_card.total_packed_pcs, 800)
        self.assertEqual(self.job_card.total_packing_used_pcs, 1000)

    def test_packing_exceeding_printed_limit_fails(self):
        record = Production(
            job_card=self.job_card,
            entry_type='packing',
            date=date(2026, 1, 3),
            shift='A',
            packing_qty=900,
            sorting_waste_qty=200,
            sorter=self.sorter,
        )
        with self.assertRaises(ValidationError):
            record.full_clean()

    def test_cut_and_pack_uses_order_qty_limit(self):
        cut_pack_job = JobCard.objects.create(
            job_card_no='JC-CUT-001',
            SKU='SKU-CUT',
            order_qty=500,
            is_print_job=False,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_sheet_quantity=0,
            total_colors=0,
            plate_set_no='N/A',
            machine_name=self.machine,
        )
        record = Production(
            job_card=cut_pack_job,
            entry_type='packing',
            date=date(2026, 1, 3),
            shift='B',
            packing_qty=400,
            sorting_waste_qty=100,
            sorter=self.sorter,
        )
        record.full_clean()
        record.save()
        self.assertEqual(cut_pack_job.total_packed_pcs, 400)


class PackingJobCardSearchTests(TestCase):
    def setUp(self):
        User = get_user_model()
        self.user = User.objects.create_user(username='packer', password='pass')
        self.client = Client()
        self.client.force_login(self.user)
        self.machine = Machine.objects.create(name='Press 2')
        self.sorter = Sorter.objects.create(name='Sorter B')
        self.print_job = JobCard.objects.create(
            job_card_no='JC-PRINT-SEARCH',
            SKU='SKU-PRINT',
            order_qty=1000,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=100,
            total_sheet_quantity=100,
            total_colors=4,
            plate_set_no='PLATE-2',
            machine_name=self.machine,
        )
        self.no_print_job = JobCard.objects.create(
            job_card_no='JC-NO-PRINT',
            SKU='SKU-NOPRINT',
            order_qty=500,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=100,
            total_sheet_quantity=100,
            total_colors=4,
            plate_set_no='PLATE-3',
            machine_name=self.machine,
        )
        self.cut_pack_job = JobCard.objects.create(
            job_card_no='JC-CUT-SEARCH',
            SKU='SKU-CUT',
            order_qty=300,
            is_print_job=False,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_sheet_quantity=0,
            total_colors=0,
            plate_set_no='N/A',
            machine_name=self.machine,
        )
        Production.objects.create(
            job_card=self.print_job,
            entry_type='printing',
            date=date(2026, 1, 2),
            shift='A',
            machine=self.machine,
            output_sheets=50,
            impressions=50,
            planned_time=30,
            run_time=30,
        )

    def test_eligible_queryset_includes_print_jobs_with_printing_entry(self):
        ids = set(_packing_eligible_job_cards_queryset().values_list('id', flat=True))
        self.assertIn(self.print_job.id, ids)
        self.assertIn(self.cut_pack_job.id, ids)
        self.assertNotIn(self.no_print_job.id, ids)

    def test_packing_search_returns_matching_jobs(self):
        url = reverse('packing_job_card_search')
        response = self.client.get(url, {'q': 'JC-PRINT'})
        self.assertEqual(response.status_code, 200)
        data = response.json()
        job_ids = {item['id'] for item in data['results']}
        self.assertIn(self.print_job.id, job_ids)
        self.assertNotIn(self.no_print_job.id, job_ids)

    def test_packing_search_requires_two_chars(self):
        url = reverse('packing_job_card_search')
        response = self.client.get(url, {'q': 'J'})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()['results'], [])


class PrintingJobCardSearchTests(TestCase):
    def setUp(self):
        User = get_user_model()
        self.user = User.objects.create_user(username='printer', password='pass')
        self.client = Client()
        self.client.force_login(self.user)
        self.machine = Machine.objects.create(name='Press 3')
        self.print_job = JobCard.objects.create(
            job_card_no='JC-PRINT-ONLY',
            SKU='SKU-ONLY',
            order_qty=1000,
            ups=10,
            is_print_job=True,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_impressions_required=100,
            total_sheet_quantity=100,
            total_colors=4,
            plate_set_no='PLATE-4',
            machine_name=self.machine,
        )

    def test_printing_search_returns_matching_jobs(self):
        url = reverse('printing_job_card_search')
        response = self.client.get(url, {'q': 'JC-PRINT'})
        self.assertEqual(response.status_code, 200)
        job_ids = {item['id'] for item in response.json()['results']}
        self.assertIn(self.print_job.id, job_ids)
