from datetime import date, timedelta

from django.core.exceptions import ValidationError
from django.test import TestCase, Client
from django.contrib.auth import get_user_model
from django.urls import reverse
from django.utils import timezone

from .jc_numbering import allocate_next_jc_number
from .jobcard_service import execute_job_card_action
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

    def test_print_job_dispatch_exceeding_production_is_allowed(self):
        """Dispatch is bound to order qty, not packed qty — an operator's
        under-logged production entry must not block a valid dispatch."""
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
        dispatch.save()

        self.assertEqual(job_card.total_dispatch, 60)

    def test_print_job_dispatch_exceeding_order_qty_raises(self):
        job_card = JobCard.objects.create(
            job_card_no='JC-02-26-0002',
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
            packing_qty=100,
            sorting_waste_qty=0,
            sorter=self._ensure_sorter(),
        )

        dispatch = Dispatch(
            job_card=job_card,
            dc_no='DC-TEST-003',
            dispatch_date='2026-01-02',
            dispatch_qty=101,
            created_by=self.user,
        )
        with self.assertRaises(ValidationError):
            dispatch.save()

    def test_allocate_next_jc_number_always_includes_pp(self):
        SequenceCounter.objects.all().delete()
        job_card_no = allocate_next_jc_number(date(2026, 6, 26))
        self.assertRegex(job_card_no, r'^JC-06-26-PP-\d{4}$')


class ExcessProductionToStockTests(TestCase):
    """Once packed qty passes order qty, the excess should be carried
    forward as stock for the next repeat run of the same SKU."""

    def setUp(self):
        self.user = get_user_model().objects.create_user(username='stock_tester', password='testpass')
        self.machine = Machine.objects.create(name='Stock Test Machine')

    def _ensure_sorter(self):
        from core.models import Sorter
        sorter, _ = Sorter.objects.get_or_create(name='Stock Test Sorter')
        return sorter

    def _make_job(self, jc_number, order_qty):
        planning_job = PlanningJob.objects.create(
            jc_number=jc_number, order_qty=order_qty, status='released',
            plan_date=date.today(), plan_month='July 2026', sku='SKU-STOCK-1',
        )
        printed_qty = order_qty + 15
        job_card = JobCard.objects.create(
            job_card_no=jc_number, planning_job=planning_job, order_qty=order_qty,
            SKU='SKU-STOCK-1', ups=1, is_print_job=True, total_sheet_quantity=printed_qty,
            total_colors=4, status='in_production', po_date=date(2026, 1, 1),
            plate_set_no='PLATE-1', machine_name=self.machine,
            total_impressions_required=printed_qty,
        )
        Production.objects.create(
            job_card=job_card, entry_type='printing', date='2026-01-01', shift='A',
            machine=self.machine, output_sheets=printed_qty, waste_sheets=0,
            impressions=printed_qty, planned_time=60, run_time=60,
        )
        return planning_job, job_card

    def test_packing_below_order_qty_leaves_stock_at_zero(self):
        planning_job, job_card = self._make_job('JC-STOCK-0001', 100)
        Production.objects.create(
            job_card=job_card, entry_type='packing', date='2026-01-01', shift='A',
            packing_qty=80, sorting_waste_qty=0, sorter=self._ensure_sorter(),
        )
        planning_job.refresh_from_db()
        self.assertEqual(int(planning_job.stock_qty or 0), 0)

    def test_packing_beyond_order_qty_carries_excess_to_stock(self):
        planning_job, job_card = self._make_job('JC-STOCK-0002', 100)
        Production.objects.create(
            job_card=job_card, entry_type='packing', date='2026-01-01', shift='A',
            packing_qty=100, sorting_waste_qty=0, sorter=self._ensure_sorter(),
        )
        Production.objects.create(
            job_card=job_card, entry_type='packing', date='2026-01-02', shift='A',
            packing_qty=15, sorting_waste_qty=0, sorter=self._ensure_sorter(),
        )
        planning_job.refresh_from_db()
        self.assertEqual(int(planning_job.stock_qty or 0), 15)


class StockAwareCompletionBlockersTests(TestCase):
    """Reproduces two real stuck-job patterns reported against the Job Card
    Finalization queue: JC-08-26-PP-1683 (printing fell short of dispatch
    because part of it was covered by existing stock) and the fully-stock
    case where a job is dispatched entirely from stock with no printing or
    packing entries of its own at all."""

    def setUp(self):
        self.machine = Machine.objects.create(name='Blocker Test Machine')

    def _make_job_card(self, jc_number, order_qty, *, dispatched, printed, packed, stock_qty):
        from core.jobcard_service import job_card_completion_blockers
        self._blockers = job_card_completion_blockers  # exposed for readability below

        planning_job = PlanningJob.objects.create(
            jc_number=jc_number, order_qty=order_qty, status='in_production',
            plan_date=date.today(), plan_month='August 2026', sku=f'SKU-{jc_number}',
            stock_qty=stock_qty,
        )
        job_card = JobCard.objects.create(
            job_card_no=jc_number, planning_job=planning_job, order_qty=order_qty,
            SKU=f'SKU-{jc_number}', ups=1, is_print_job=True,
            total_sheet_quantity=max(printed, 1), total_colors=4, status='in_production',
            po_date=date(2026, 1, 1), plate_set_no='PLATE-1', machine_name=self.machine,
            total_impressions_required=max(printed, 1),
        )
        if printed:
            Production.objects.create(
                job_card=job_card, entry_type='printing', date='2026-08-24', shift='A',
                machine=self.machine, output_sheets=printed, waste_sheets=0,
                impressions=printed, planned_time=60, run_time=60,
            )
        if packed:
            from core.models import Sorter
            sorter, _ = Sorter.objects.get_or_create(name='Blocker Test Sorter')
            Production.objects.create(
                job_card=job_card, entry_type='packing', date='2026-08-24', shift='A',
                packing_qty=packed, sorting_waste_qty=0, sorter=sorter,
            )
        if dispatched:
            Dispatch.objects.create(
                job_card=job_card, dc_no=f'DC-{jc_number}', dispatch_date=date(2026, 8, 27),
                dispatch_qty=dispatched,
            )
        return planning_job, job_card

    def test_partial_stock_without_stock_qty_entered_stays_blocked(self):
        """JC-1683 as found: printed/packed short of the full dispatch, and
        nobody had told the system about the stock that covered the rest —
        must stay blocked (this is not a blanket bypass)."""
        from core.jobcard_service import job_card_completion_blockers
        _planning_job, job_card = self._make_job_card(
            'JC-BLOCK-1683', order_qty=3600, dispatched=3600, printed=2420, packed=2400, stock_qty=0,
        )
        blockers = job_card_completion_blockers(job_card)
        self.assertEqual(len(blockers), 1)
        self.assertIn('Packed + stock (2400) is less than dispatched (3600)', blockers[0])

    def test_partial_stock_with_stock_qty_entered_clears(self):
        """Same job as above, but with the actual stock on hand recorded —
        matches what the Job Card Finalization 'Set Stock' action now writes."""
        from core.jobcard_service import job_card_completion_blockers
        _planning_job, job_card = self._make_job_card(
            'JC-BLOCK-1683B', order_qty=3600, dispatched=3600, printed=2420, packed=2400, stock_qty=1200,
        )
        self.assertEqual(job_card_completion_blockers(job_card), [])

    def test_fully_stock_fulfilled_job_needs_no_printing_or_packing(self):
        """JC-2015-style: the whole order was already in stock, so nothing
        was printed or packed for this run at all — only dispatch happened.
        Before this fix, 'No printing entries logged' blocked closing even
        though stock fully explained the shortfall."""
        from core.jobcard_service import job_card_completion_blockers
        _planning_job, job_card = self._make_job_card(
            'JC-BLOCK-2015', order_qty=6037, dispatched=6037, printed=0, packed=0, stock_qty=6037,
        )
        self.assertEqual(job_card_completion_blockers(job_card), [])

    def test_stock_short_of_dispatch_still_blocks_both_printing_and_packing(self):
        from core.jobcard_service import job_card_completion_blockers
        _planning_job, job_card = self._make_job_card(
            'JC-BLOCK-2015B', order_qty=6037, dispatched=6037, printed=0, packed=0, stock_qty=6000,
        )
        blockers = job_card_completion_blockers(job_card)
        self.assertEqual(len(blockers), 2)
        self.assertTrue(any('No printing entries' in b for b in blockers))
        self.assertTrue(any('Packed + stock (6000) is less than dispatched (6037)' in b for b in blockers))


class JobCardFinalizationSetStockViewTests(TestCase):
    def setUp(self):
        self.machine = Machine.objects.create(name='Set Stock Test Machine')
        self.user = get_user_model().objects.create_user(username='finalizer', password='testpass123')
        from core.models import Permission, UserPermissionOverride
        # The soft-coded access-control system (core/permissions.py) has no
        # seed data in a fresh test DB (Role/Permission rows come from the
        # seed_access_control management command, not a migration), so grant
        # this test user the specific permission the view requires.
        permission, _ = Permission.objects.get_or_create(
            code='action.finalize_job_card', defaults={'name': 'Finalize Job Card'},
        )
        UserPermissionOverride.objects.get_or_create(
            user=self.user, permission=permission, defaults={'granted': True},
        )
        self.client.force_login(self.user)

        self.planning_job = PlanningJob.objects.create(
            jc_number='JC-SETSTOCK-0001', order_qty=1000, status='in_production',
            plan_date=date.today(), plan_month='August 2026', sku='SKU-SETSTOCK-1',
        )
        self.job_card = JobCard.objects.create(
            job_card_no='JC-SETSTOCK-0001', planning_job=self.planning_job, order_qty=1000,
            SKU='SKU-SETSTOCK-1', ups=1, is_print_job=True, total_sheet_quantity=1,
            total_colors=4, status='in_production', po_date=date(2026, 1, 1),
            plate_set_no='PLATE-1', machine_name=self.machine, total_impressions_required=1,
        )

    def test_set_stock_updates_planning_job_and_logs_change(self):
        from core.models import ChangeLog
        response = self.client.post(
            reverse('job_card_finalization_set_stock'),
            {'job_card_id': self.job_card.id, 'stock_qty': '1200'},
        )
        self.assertRedirects(response, reverse('job_card_finalization_queue'))
        self.planning_job.refresh_from_db()
        self.assertEqual(int(self.planning_job.stock_qty), 1200)
        self.assertTrue(
            ChangeLog.objects.filter(
                entity_type='job_card', record_id=self.job_card.id, action='update',
            ).exists()
        )

    def test_set_stock_rejects_negative_value(self):
        response = self.client.post(
            reverse('job_card_finalization_set_stock'),
            {'job_card_id': self.job_card.id, 'stock_qty': '-5'},
        )
        self.assertRedirects(response, reverse('job_card_finalization_queue'))
        self.planning_job.refresh_from_db()
        self.assertIsNone(self.planning_job.stock_qty)

    def test_set_stock_requires_planning_job(self):
        orphan_card = JobCard.objects.create(
            job_card_no='JC-SETSTOCK-ORPHAN', order_qty=100, SKU='SKU-ORPHAN', ups=1,
            is_print_job=True, total_sheet_quantity=1, total_colors=4, status='in_production',
            po_date=date(2026, 1, 1), plate_set_no='PLATE-1', machine_name=self.machine,
            total_impressions_required=1,
        )
        response = self.client.post(
            reverse('job_card_finalization_set_stock'),
            {'job_card_id': orphan_card.id, 'stock_qty': '50'},
        )
        self.assertRedirects(response, reverse('job_card_finalization_queue'))


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


class JobCardPlanningValidationTests(TestCase):
    def setUp(self):
        self.machine = Machine.objects.create(name='Validation Test Machine')

    def test_plate_set_no_is_optional_for_planning_approval(self):
        job_card = JobCard.objects.create(
            job_card_no='JC-PLATE-OPTIONAL-001',
            SKU='SKU-OPTIONAL-PLATE',
            order_qty=100,
            ups=10,
            is_print_job=True,
            total_impressions_required=100,
            status='draft',
            po_date=date(2026, 1, 1),
            total_sheet_quantity=10,
            total_colors=4,
            wastage=0,
            machine_name=self.machine,
            plate_set_no='',
        )

        self.assertNotIn('Plate Set', job_card.planning_missing_fields())
        execute_job_card_action(job_card, 'approve_planning', actor=None)
        job_card.refresh_from_db()
        self.assertEqual(job_card.workflow_status, 'planning_approved')


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


class MasterDataAddNewEntryTests(TestCase):
    """Admin can add new Machines/Application Types (and other master data)
    directly from the Master Data screen, not just edit existing rows."""

    def setUp(self):
        User = get_user_model()
        self.admin = User.objects.create_user(username='md_admin_add', password='pass12345')
        from .models import UserProfile
        profile = UserProfile.objects.get(user=self.admin)
        profile.role = 'admin'
        profile.save(update_fields=['role'])
        self.client.login(username='md_admin_add', password='pass12345')

    def test_admin_can_create_new_machine_from_master_data(self):
        response = self.client.post(reverse('master_data'), {
            'entity_type': 'machine',
            'action': 'create_master',
            'new_name': 'Brand New GTO',
        })
        self.assertEqual(response.status_code, 302)
        self.assertTrue(Machine.objects.filter(name='Brand New GTO').exists())

    def test_admin_can_create_new_application_type_from_master_data(self):
        from .models import ApplicationType
        response = self.client.post(reverse('master_data'), {
            'entity_type': 'application',
            'action': 'create_master',
            'new_name': 'Spot UV',
        })
        self.assertEqual(response.status_code, 302)
        self.assertTrue(ApplicationType.objects.filter(name='Spot UV').exists())

    def test_create_master_rejects_duplicate_name(self):
        from .models import ApplicationType
        ApplicationType.objects.create(name='Foil Stamp')
        response = self.client.post(reverse('master_data'), {
            'entity_type': 'application',
            'action': 'create_master',
            'new_name': 'foil stamp',  # case-insensitive duplicate
        })
        self.assertEqual(response.status_code, 302)
        self.assertEqual(ApplicationType.objects.filter(name__iexact='foil stamp').count(), 1)


class QuickAddMasterTests(TestCase):
    """quick_add_master powers the '+' buttons on the SKU master-entry form
    (Application, Machine, Material, Product Type)."""

    def setUp(self):
        User = get_user_model()
        self.planner = User.objects.create_user(username='qam_planner', password='pass12345')
        from .models import UserProfile
        UserProfile.objects.filter(user=self.planner).update(role='planner')

    def test_planner_can_quick_add_application_type(self):
        from .models import ApplicationType
        self.client.login(username='qam_planner', password='pass12345')
        response = self.client.post(reverse('quick_add_master'), {
            'type': 'application',
            'name': 'Spot UV',
        })
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data['ok'])
        self.assertTrue(data['created'])
        self.assertTrue(ApplicationType.objects.filter(name='Spot UV').exists())

    def test_quick_add_application_reuses_existing_case_insensitively(self):
        from .models import ApplicationType
        ApplicationType.objects.create(name='Foil Stamp')
        self.client.login(username='qam_planner', password='pass12345')
        response = self.client.post(reverse('quick_add_master'), {
            'type': 'application',
            'name': 'foil stamp',
        })
        data = response.json()
        self.assertTrue(data['ok'])
        self.assertFalse(data['created'])
        self.assertEqual(ApplicationType.objects.filter(name__iexact='foil stamp').count(), 1)

    def test_quick_add_application_requires_planner_or_admin(self):
        User = get_user_model()
        from .models import UserProfile
        qc_user = User.objects.create_user(username='qam_qc', password='pass12345')
        UserProfile.objects.filter(user=qc_user).update(role='qc')
        self.client.login(username='qam_qc', password='pass12345')
        response = self.client.post(reverse('quick_add_master'), {
            'type': 'application',
            'name': 'Spot UV',
        })
        self.assertEqual(response.status_code, 403)

    def test_planner_can_quick_add_machine(self):
        self.client.login(username='qam_planner', password='pass12345')
        response = self.client.post(reverse('quick_add_master'), {
            'type': 'machine',
            'name': 'Quick Add Machine',
        })
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data['ok'])
        self.assertTrue(Machine.objects.filter(name='Quick Add Machine').exists())


class MasterDataMachinePrintSizeTests(TestCase):
    """Machine print size is entered in inches on the Master Data screen
    (matching how planners record sheet sizes) and stored internally in
    mm - verifies the conversion round-trips correctly."""

    def setUp(self):
        User = get_user_model()
        self.admin = User.objects.create_user(username='md_admin', password='pass12345')
        from .models import UserProfile
        profile = UserProfile.objects.get(user=self.admin)
        profile.role = 'admin'
        profile.save(update_fields=['role'])
        self.machine = Machine.objects.create(name='Test SM74', is_active=True)

    def test_inches_input_converts_to_mm_on_save(self):
        self.client.login(username='md_admin', password='pass12345')
        response = self.client.post(reverse('master_data'), {
            'entity_type': 'machine',
            'action': 'edit_master',
            'record_id': self.machine.id,
            'new_name': self.machine.name,
            'standard_impressions_per_hour': '4000',
            'standard_setup_minutes_per_color': '15',
            'plate_life_impressions': '25000',
            'max_print_length_in': '29.13',  # ~740mm
            'max_print_width_in': '41.34',   # ~1050mm
        })
        self.assertEqual(response.status_code, 302)
        self.machine.refresh_from_db()
        self.assertAlmostEqual(float(self.machine.max_print_length_mm), 740.0, delta=0.5)
        self.assertAlmostEqual(float(self.machine.max_print_width_mm), 1050.0, delta=0.5)


class SkuPreferredMachineLearningTests(TestCase):
    """Part C: actual production machine should write back to the SKU
    master's preferred machine, unless explicitly locked."""

    def setUp(self):
        self.machine = Machine.objects.create(name='GTO 2A', is_active=True)
        self.other_machine = Machine.objects.create(name='SM 74', is_active=True)

        from planning.models import PlanningJob, SkuRecipe
        self.recipe = SkuRecipe.objects.create(sku='SKU-LEARN-1', machine_name='', is_active=True)
        self.pj = PlanningJob.objects.create(
            jc_number='JC-LEARN-1', order_qty=100, status='released',
            plan_date=date.today(), plan_month='July 2026', sku='SKU-LEARN-1',
        )
        self.jc = JobCard.objects.create(
            job_card_no='JC-LEARN-1', planning_job=self.pj, order_qty=100,
            total_sheet_quantity=100, status='in_production', is_active=True,
            SKU='SKU-LEARN-1', po_date=date.today(), machine_name=self.machine,
            total_impressions_required=100, total_colors=4,
        )

    def test_production_save_writes_back_actual_machine_to_sku_master(self):
        from planning.models import SkuRecipe

        Production.objects.create(
            entry_type='printing', job_card=self.jc, date=date.today(), shift='A',
            output_sheets=100, status='completed', machine=self.machine,
        )
        self.recipe.refresh_from_db()
        self.assertEqual(self.recipe.machine_name, 'GTO 2A')

    def test_locked_sku_master_is_not_overwritten(self):
        from planning.models import SkuRecipe

        self.recipe.machine_name = 'SM 74'
        self.recipe.machine_name_locked = True
        self.recipe.save(update_fields=['machine_name', 'machine_name_locked'])

        Production.objects.create(
            entry_type='printing', job_card=self.jc, date=date.today(), shift='A',
            output_sheets=100, status='completed', machine=self.machine,
        )
        self.recipe.refresh_from_db()
        self.assertEqual(self.recipe.machine_name, 'SM 74')


class MachineRoutingTests(TestCase):
    def _make_fleet(self):
        gto1a = Machine.objects.create(name='GTO 1A', machine_type='offset_printing', machine_group_code='GTO1', default_colors=1, operational_colors=1)
        gto1b = Machine.objects.create(name='GTO 1B', machine_type='offset_printing', machine_group_code='GTO1', default_colors=1, operational_colors=1)
        gto2a = Machine.objects.create(name='GTO 2A', machine_type='offset_printing', machine_group_code='GTO2', default_colors=2, operational_colors=2)
        gto2b = Machine.objects.create(name='GTO 2B', machine_type='offset_printing', machine_group_code='GTO2', default_colors=2, operational_colors=2)
        gto2c = Machine.objects.create(name='GTO 2C', machine_type='offset_printing', machine_group_code='GTO2', default_colors=2, operational_colors=2)
        sm74 = Machine.objects.create(
            name='SM 74', machine_type='offset_printing', machine_group_code='SM74', default_colors=5, operational_colors=5,
            max_print_length_mm=740, max_print_width_mm=1050,
        )
        for m in (gto1a, gto1b, gto2a, gto2b, gto2c):
            m.max_print_length_mm = 520
            m.max_print_width_mm = 740
            m.save()
        return [gto1a, gto1b, gto2a, gto2b, gto2c, sm74]

    def test_color_class_treats_symmetric_plus_as_single_colour(self):
        from core.machine_routing import color_class

        self.assertEqual(color_class('1+1'), 1)
        self.assertEqual(color_class('4+4'), 4)
        self.assertEqual(color_class('4'), 4)

    def test_one_color_job_routes_to_gto1_pool_named(self):
        from core.machine_routing import build_pools, route_job

        machines = self._make_fleet()
        pools = build_pools(machines)
        result = route_job('1', '18*25', pools, size_gate_machine_code='SM74')
        self.assertEqual(result['pool_key'], 'GTO1')
        self.assertIn('GTO 1A', result['pool_label'])
        self.assertIn('GTO 1B', result['pool_label'])
        self.assertEqual(result['passes'], 1)

    def test_three_color_job_routes_to_gto2_with_two_passes(self):
        from core.machine_routing import build_pools, route_job

        machines = self._make_fleet()
        pools = build_pools(machines)
        result = route_job('3', '18*25', pools, size_gate_machine_code='SM74')
        self.assertEqual(result['pool_key'], 'GTO2')
        self.assertEqual(result['passes'], 2)

    def test_oversized_job_routes_to_sm74_regardless_of_color(self):
        from core.machine_routing import build_pools, route_job

        machines = self._make_fleet()
        pools = build_pools(machines)
        # 30x40 inches -> 762x1016mm, exceeds the GTO groups' 520x740mm max.
        result = route_job('1', '30*40', pools, size_gate_machine_code='SM74')
        self.assertEqual(result['pool_key'], 'SM74')

    def test_front_back_1plus1_routes_to_gto1(self):
        from core.machine_routing import build_pools, route_job

        machines = self._make_fleet()
        pools = build_pools(machines)
        result = route_job('1+1', '18*25', pools, size_gate_machine_code='SM74')
        self.assertEqual(result['pool_key'], 'GTO1')

    def test_degraded_two_color_machine_joins_single_color_pool(self):
        from core.machine_routing import build_pools

        machines = self._make_fleet()
        gto2a = next(m for m in machines if m.name == 'GTO 2A')
        gto2a.operational_colors = 1
        gto2a.save()
        pools = build_pools(machines)
        gto1_pool = pools['GTO1']
        self.assertIn('GTO 2A', [m.name for m in gto1_pool.members])

    def test_maintenance_machine_excluded_from_pools(self):
        from core.machine_routing import build_pools

        machines = self._make_fleet()
        gto1b = next(m for m in machines if m.name == 'GTO 1B')
        gto1b.operational_colors = 0
        gto1b.save()
        pools = build_pools(machines)
        gto1_pool = pools['GTO1']
        self.assertNotIn('GTO 1B', [m.name for m in gto1_pool.members])
        self.assertIn('GTO 1B', [m.name for m in gto1_pool.maintenance_members])

