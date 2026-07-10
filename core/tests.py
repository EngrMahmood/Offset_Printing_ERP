from datetime import date, timedelta

from django.core.exceptions import ValidationError
from django.test import TestCase, Client
from django.contrib.auth import get_user_model
from django.urls import reverse
from django.utils import timezone

from .jc_numbering import allocate_next_jc_number
from .models import DeliveryLocation, JobCard, Machine, Production, Dispatch, SequenceCounter
from .services import sync_delivery_locations_from_planning
from .views import _dispatch_remaining_badge
from planning.models import PlanningJob


class DispatchValidationTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(username='tester', password='testpass')
        self.machine = Machine.objects.create(name='Test Machine')

    def _ensure_sorter(self):
        from core.models import Sorter
        sorter, _ = Sorter.objects.get_or_create(name='Test Sorter')
        return sorter

    def test_print_job_dispatch_within_produced_pieces_succeeds(self):
        job_card = JobCard.objects.create(
            job_card_no='JC-01-26-1071',
            SKU='SKU-001',
            order_qty=100,
            ups=10,
            is_print_job=True,
            total_impressions_required=100,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_sheet_quantity=10,
            total_colors=4,
            plate_set_no='PLATE-1',
            machine_name=self.machine,
        )
        Production.objects.create(
            job_card=job_card,
            entry_type='printing',
            date='2026-01-01',
            shift='A',
            machine=self.machine,
            output_sheets=10,
            waste_sheets=0,
            impressions=100,
            planned_time=60,
            run_time=60,
        )
        Production.objects.create(
            job_card=job_card,
            entry_type='packing',
            date='2026-01-01',
            shift='A',
            packing_qty=50,
            sorting_waste_qty=0,
            sorter=self._ensure_sorter(),
        )

        dispatch = Dispatch(
            job_card=job_card,
            dc_no='DC-TEST-001',
            dispatch_date='2026-01-02',
            dispatch_qty=50,
            created_by=self.user,
        )
        dispatch.save()

        self.assertEqual(job_card.total_dispatch, 50)

    def test_print_job_dispatch_exceeding_production_raises(self):
        job_card = JobCard.objects.create(
            job_card_no='JC-02-26-0001',
            SKU='SKU-001',
            order_qty=100,
            ups=10,
            is_print_job=True,
            total_impressions_required=50,
            status='in_production',
            po_date=date(2026, 1, 1),
            total_sheet_quantity=10,
            total_colors=4,
            plate_set_no='PLATE-1',
            machine_name=self.machine,
        )
        Production.objects.create(
            job_card=job_card,
            entry_type='printing',
            date='2026-01-01',
            shift='A',
            machine=self.machine,
            output_sheets=5,
            waste_sheets=0,
            impressions=50,
            planned_time=60,
            run_time=60,
        )
        Production.objects.create(
            job_card=job_card,
            entry_type='packing',
            date='2026-01-01',
            shift='A',
            packing_qty=50,
            sorting_waste_qty=0,
            sorter=self._ensure_sorter(),
        )

        dispatch = Dispatch(
            job_card=job_card,
            dc_no='DC-TEST-002',
            dispatch_date='2026-01-02',
            dispatch_qty=60,
            created_by=self.user,
        )
        with self.assertRaises(ValidationError):
            dispatch.save()

    def test_allocate_next_jc_number_always_includes_pp(self):
        SequenceCounter.objects.all().delete()
        job_card_no = allocate_next_jc_number(date(2026, 6, 26))
        self.assertRegex(job_card_no, r'^JC-06-26-PP-\d{4}$')


class DispatchFeatureTests(TestCase):
    def setUp(self):
        self.client = Client()
        self.user = get_user_model().objects.create_user(username='dispatch_user', password='testpass')
        self.user.profile.role = 'dispatch'
        self.user.profile.save()
        self.machine = Machine.objects.create(name='Dispatch Test Machine')
        self.job_card = JobCard.objects.create(
            job_card_no='JC-03-26-0001',
            SKU='SKU-DISP-1',
            order_qty=100,
            ups=10,
            is_print_job=True,
            total_impressions_required=100,
            status='in_production',
            po_date=date(2026, 1, 1),
            PO_No='PO-1001',
            destination='Acme Corp',
            total_sheet_quantity=10,
            total_colors=4,
            plate_set_no='PLATE-1',
            machine_name=self.machine,
        )
        Production.objects.create(
            job_card=self.job_card,
            date='2026-01-01',
            shift='A',
            machine=self.machine,
            output_sheets=10,
            waste_sheets=0,
            impressions=100,
            planned_time=60,
            run_time=60,
        )
        from core.models import Sorter
        sorter, _ = Sorter.objects.get_or_create(name='Test Sorter')
        Production.objects.create(
            job_card=self.job_card,
            entry_type='packing',
            date='2026-01-01',
            shift='A',
            packing_qty=100,
            sorting_waste_qty=0,
            sorter=sorter,
        )

    def test_dispatch_job_card_search_returns_matching_job_card(self):
        self.client.login(username='dispatch_user', password='testpass')
        response = self.client.get(reverse('dispatch_job_card_search'), {'q': 'SKU-DISP'})
        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(len(payload['results']), 1)
        self.assertEqual(payload['results'][0]['job_card_no'], 'JC-03-26-0001')
        self.assertEqual(payload['results'][0]['remaining'], 100)

    def test_dispatch_dc_duplicate_check_flags_same_job_card(self):
        Dispatch.objects.create(
            job_card=self.job_card,
            dc_no='DC-100',
            dispatch_date=timezone.now().date(),
            dispatch_qty=20,
            created_by=self.user,
        )
        self.client.login(username='dispatch_user', password='testpass')
        response = self.client.get(reverse('dispatch_dc_duplicate_check'), {
            'dc_no': 'DC-100',
            'job_card_id': self.job_card.id,
        })
        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertTrue(payload['same_jc_duplicate'])
        self.assertTrue(payload['blocking'])

    def test_dispatch_dc_duplicate_check_allows_other_job_card(self):
        other_job = JobCard.objects.create(
            job_card_no='JC-03-26-0002',
            SKU='SKU-DISP-2',
            order_qty=50,
            ups=5,
            is_print_job=False,
            status='in_production',
            po_date=date(2026, 1, 1),
            PO_No='PO-1002',
            destination='Other Corp',
            total_sheet_quantity=10,
            total_colors=4,
            plate_set_no='PLATE-2',
            machine_name=self.machine,
        )
        Dispatch.objects.create(
            job_card=other_job,
            dc_no='DC-200',
            dispatch_date=timezone.now().date() - timedelta(days=2),
            dispatch_qty=10,
            created_by=self.user,
        )
        self.client.login(username='dispatch_user', password='testpass')
        response = self.client.get(reverse('dispatch_dc_duplicate_check'), {
            'dc_no': 'DC-200',
            'job_card_id': self.job_card.id,
        })
        payload = response.json()
        self.assertFalse(payload['same_jc_duplicate'])
        self.assertFalse(payload['blocking'])
        self.assertEqual(len(payload['same_dc_entries']), 1)
        self.assertEqual(payload['same_dc_entries'][0]['job_card_no'], 'JC-03-26-0002')
        self.assertEqual(payload['same_dc_sku_count'], 1)
        self.assertEqual(payload['same_dc_line_count'], 1)

    def test_dispatch_entry_blocks_duplicate_dc_for_same_job_card(self):
        Dispatch.objects.create(
            job_card=self.job_card,
            dc_no='DC-300',
            dispatch_date=timezone.now().date(),
            dispatch_qty=10,
            created_by=self.user,
        )
        self.client.login(username='dispatch_user', password='testpass')
        response = self.client.post(reverse('dispatch_entry'), {
            'job_card': self.job_card.id,
            'dc_no': 'DC-300',
            'dispatch_date': '2026-01-03',
            'dispatch_qty': 5,
        })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(Dispatch.objects.filter(job_card=self.job_card, dc_no='DC-300').count(), 1)

    def test_dispatch_entry_requires_dc_no(self):
        self.client.login(username='dispatch_user', password='testpass')
        response = self.client.post(reverse('dispatch_entry'), {
            'job_card': self.job_card.id,
            'dispatch_date': '2026-01-03',
            'dispatch_qty': 5,
        })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(Dispatch.objects.filter(job_card=self.job_card).count(), 0)

    def test_dispatch_remaining_badge_states(self):
        complete = _dispatch_remaining_badge(100, 100)
        partial = _dispatch_remaining_badge(100, 40)
        not_started = _dispatch_remaining_badge(100, 0)
        self.assertEqual(complete['label'], 'Complete')
        self.assertEqual(complete['badge_class'], 'erp-badge-completed')
        self.assertEqual(partial['label'], '60 left')
        self.assertEqual(partial['badge_class'], 'erp-badge-pending')
        self.assertEqual(not_started['label'], 'Not dispatched')
        self.assertEqual(not_started['badge_class'], 'erp-badge-draft')


class DeliveryLocationSyncTests(TestCase):
    def test_sync_creates_delivery_location_from_planning_job_destination(self):
        user = get_user_model().objects.create_user(username='delivery_sync', password='testpass')
        PlanningJob.objects.create(
            jc_number='JC-DL-001',
            sku='SKU-DL-1',
            destination='Main Warehouse',
            status='draft',
            created_by=user,
        )

        created = sync_delivery_locations_from_planning()

        self.assertEqual(created, 1)
        self.assertTrue(DeliveryLocation.objects.filter(name__iexact='Main Warehouse').exists())

    def test_sync_skips_existing_delivery_location(self):
        user = get_user_model().objects.create_user(username='delivery_sync2', password='testpass')
        DeliveryLocation.objects.create(name='Main Warehouse')
        PlanningJob.objects.create(
            jc_number='JC-DL-002',
            sku='SKU-DL-2',
            destination='Main Warehouse',
            status='draft',
            created_by=user,
        )

        created = sync_delivery_locations_from_planning()

        self.assertEqual(created, 0)
        self.assertEqual(DeliveryLocation.objects.filter(name__iexact='Main Warehouse').count(), 1)


class PrintColorNormalizationTests(TestCase):
    def test_sheet_decimal_notation_maps_to_plus_form(self):
        from core.print_colors import normalize_color_spec_value

        self.assertEqual(normalize_color_spec_value('4.0'), '4+0')
        self.assertEqual(normalize_color_spec_value('1.0'), '1+0')
        self.assertEqual(normalize_color_spec_value(4.0), '4+0')
        self.assertEqual(normalize_color_spec_value('4+0'), '4+0')

    def test_mangled_legacy_values_repair_to_plus_form(self):
        from core.print_colors import normalize_color_spec_value, repair_mangled_decimal_color_spec

        self.assertEqual(repair_mangled_decimal_color_spec('40'), '4+0')
        self.assertEqual(repair_mangled_decimal_color_spec('10'), '1+0')
        self.assertEqual(normalize_color_spec_value('40'), '4+0')
        self.assertEqual(normalize_color_spec_value('10'), '1+0')

    def test_legacy_color_text_still_normalizes(self):
        from core.print_colors import normalize_color_spec_value

        self.assertEqual(normalize_color_spec_value('4 color'), '4')
        self.assertEqual(normalize_color_spec_value('4C'), '4')

