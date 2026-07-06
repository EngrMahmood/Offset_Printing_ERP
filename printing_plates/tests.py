from django.test import TestCase
from django.contrib.auth import get_user_model
from django.urls import reverse
from django.utils import timezone
from core.models import UserProfile, Vendor, Machine, Department, Material, JobCard
from planning.models import PlanningJob
from printing_plates.models import PlateRequest

User = get_user_model()

class PlateWorkflowTestCase(TestCase):
    def setUp(self):
        # Create users with different roles
        self.admin_user = User.objects.create_user(username='admin', email='admin@erp.com', password='password')
        self.admin_profile = self.admin_user.profile
        self.admin_profile.role = 'admin'
        self.admin_profile.save()
        
        self.designer_user = User.objects.create_user(username='designer', email='designer@erp.com', password='password')
        self.designer_profile = self.designer_user.profile
        self.designer_profile.role = 'graphics_designer'
        self.designer_profile.save()
        
        self.operator_user = User.objects.create_user(username='operator', email='operator@erp.com', password='password')
        self.operator_profile = self.operator_user.profile
        self.operator_profile.role = 'operator'
        self.operator_profile.save()
        
        # Setup master data
        self.machine = Machine.objects.create(name='KBA Rapida', standard_impressions_per_hour=10000, standard_setup_minutes_per_color=15)
        self.department = Department.objects.create(name='Offset Printing')
        self.material = Material.objects.create(name='Art Paper 300gsm')
        self.vendor = Vendor.objects.create(name='Dot Max')
        
        # Setup PlanningJob and JobCard
        self.planning_job = PlanningJob.objects.create(
            jc_number='JC-0001',
            planning_stage='new_plate_making'
        )

        self.jobcard = JobCard.objects.create(
            job_card_no='JC-0001',
            planning_job=self.planning_job,
            SKU='SKU-001',
            order_qty=5000,
            total_impressions_required=5000,
            machine_name=self.machine,
            department=self.department,
            material=self.material
        )
        
        # Setup PlateRequest
        self.plate_request = PlateRequest.objects.create(
            planning_job=self.planning_job,
            job_card=self.jobcard,
            machine=self.machine,
            department=self.department,
            requested_by=self.admin_user
        )

    def test_vendor_creation_and_quick_add(self):
        # Test AJAX quick-add vendor
        self.client.login(username='designer', password='password')
        url = reverse('quick_add_master')
        response = self.client.post(url, {
            'type': 'vendor',
            'name': 'Ali Print Pack'
        })
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data['ok'])
        self.assertEqual(data['name'], 'Ali Print Pack')
        
        # Check Vendor model
        self.assertTrue(Vendor.objects.filter(name='Ali Print Pack').exists())
        
        # Test duplicate vendor addition
        response = self.client.post(url, {
            'type': 'vendor',
            'name': 'Ali Print Pack'
        })
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data['ok'])
        self.assertFalse(data['created'])  # should return existing without creating duplicate

    def test_role_based_workflow_permission(self):
        # Operator tries to send to vendor (should be denied/redirected)
        self.client.login(username='operator', password='password')
        url = reverse('printing_plates:request_action', kwargs={'pk': self.plate_request.pk})
        
        response = self.client.post(url, {
            'action': 'send_to_vendor',
            'vendor': 'Dot Max',
            'set_no': 'Set 1'
        })
        # The mixin redirects or denies (returns 403 Forbidden)
        self.assertEqual(response.status_code, 403)
        # Check database: status should still be draft
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_DRAFT)

    def test_workflow_transitions(self):
        # Login as authorized designer
        self.client.login(username='designer', password='password')
        url = reverse('printing_plates:request_action', kwargs={'pk': self.plate_request.pk})
        
        # 1. Send to Vendor
        response = self.client.post(url, {
            'action': 'send_to_vendor',
            'vendor': 'Dot Max',
            'set_no': 'Set-01',
            'new_set_no': 'New-Set-01',
            'awc_no': '973',
            'print_color': '4',
            'plate_color': 'Black, Special 1',
            'sets_required': '2',
            'plate_quantity': '4',
            'remarks': 'First print plates',
            'size_w_mm': '100',
            'size_h_mm': '200',
            'print_sheet_size': '28x40',
            'ups': '4',
            'purchase_sheet_size': '30x42',
            'purchase_sheet_ups': '2',
            'die_cutting': 'NO',
        })
        self.assertEqual(response.status_code, 302)
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_SENT)
        self.assertEqual(self.plate_request.vendor, 'Dot Max')
        self.assertEqual(self.plate_request.set_no, 'Set-01')
        self.assertEqual(self.plate_request.new_set_no, 'New-Set-01')
        self.assertEqual(self.plate_request.awc_no, '973')
        self.assertEqual(self.plate_request.plate_color, 'Black, Special 1')
        self.assertEqual(self.plate_request.sets_required, 2)
        self.assertEqual(self.plate_request.plate_quantity, 4)
        self.assertEqual(self.plate_request.plate_quantity_display, '4 (2 sets)')
        self.assertEqual(self.plate_request.sent_by, self.designer_user)
        self.assertIsNotNone(self.plate_request.sent_at)
        self.planning_job.refresh_from_db()
        self.assertEqual(self.planning_job.color_spec, '4')
        self.assertEqual(self.planning_job.total_colors, 4)
        
        # 2. Receive from Vendor
        response = self.client.post(url, {
            'action': 'receive_from_vendor',
            'challan': 'CH-9988',
            'box': 'BOX-A',
            'remarks': 'Received safely'
        })
        self.assertEqual(response.status_code, 302)
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_RECEIVED)
        self.assertEqual(self.plate_request.challan, 'CH-9988')
        self.assertEqual(self.plate_request.box, 'BOX-A')
        self.assertEqual(self.plate_request.received_by, self.designer_user)
        self.assertIsNotNone(self.plate_request.received_at)
        
        # 3. Issue to Production
        response = self.client.post(url, {
            'action': 'issue_to_production'
        })
        self.assertEqual(response.status_code, 302)
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_AVAILABLE)
        
        # Check that PlanningJob stage automatically transitioned to 'plate_received'
        self.planning_job.refresh_from_db()
        self.assertEqual(self.planning_job.planning_stage, 'plate_received')

    def test_detail_view_safe_lookups(self):
        # Create a plate request with no planning job or job card
        request_no_planning = PlateRequest.objects.create(
            planning_job=None,
            job_card=None,
            machine=self.machine,
            department=self.department,
            requested_by=self.admin_user
        )
        self.client.login(username='designer', password='password')
        url = reverse('printing_plates:request_detail', kwargs={'pk': request_no_planning.pk})
        response = self.client.get(url)
        self.assertEqual(response.status_code, 200)

    def test_stage_transition_validation_on_planning_job(self):
        self.client.login(username='admin', password='password')
        job = PlanningJob.objects.create(
            jc_number='JC-TEST-02',
            planning_stage='jc_ready',
            material='',
            application='',
            machine_name=''
        )
        url = reverse('planning:jobs')
        
        # 1. Try to transition to plate_making without planner details (should fail)
        response = self.client.post(url, {
            'action': 'update_planning_stage',
            'job_id': job.id,
            'planning_stage': 'plate_making'
        })
        self.assertEqual(response.status_code, 302)
        job.refresh_from_db()
        self.assertEqual(job.planning_stage, 'jc_ready')
        
        # 2. Fill in planner details and transition (should succeed)
        job.material = 'Cardboard 300gsm'
        job.application = 'Packaging'
        job.machine_name = 'KBA'
        job.save()
        
        response = self.client.post(url, {
            'action': 'update_planning_stage',
            'job_id': job.id,
            'planning_stage': 'plate_making'
        })
        self.assertEqual(response.status_code, 302)
        job.refresh_from_db()
        self.assertIn(job.planning_stage, ['new_plate_making', 'repeat_plate_making'])

    def test_designer_layout_submission_updates_recipe_and_job(self):
        from planning.models import SkuRecipe
        self.client.login(username='designer', password='password')
        
        # Setup SkuRecipe
        recipe = SkuRecipe.objects.create(
            sku='SKU-NEW-DESIGN',
            job_name='New Test Box',
            master_data_status='draft',
            is_active=True
        )
        
        # Setup PlanningJob with repeat_flag = 'New'
        job = PlanningJob.objects.create(
            jc_number='JC-NEW-01',
            sku='SKU-NEW-DESIGN',
            planning_stage='new_plate_making',
            repeat_flag='New'
        )
        
        # Setup PlateRequest
        req = PlateRequest.objects.create(
            planning_job=job,
            sku_recipe=recipe,
            requested_by=self.admin_user
        )
        
        url = reverse('printing_plates:request_action', kwargs={'pk': req.pk})
        response = self.client.post(url, {
            'action': 'send_to_vendor',
            'vendor': 'Dot Max',
            'set_no': 'Set-A',
            'new_set_no': 'New-Set-A',
            'awc_no': 'AWC-777',
            'print_color': '2',
            'plate_color': 'Cyan, Yellow',
            'plate_quantity': '2',
            'awc_no': 'AWC-777',
            'remarks': 'Layout standard prep',
            
            # Designer layout specifications
            'size_w_mm': '250.50',
            'size_h_mm': '350.75',
            'print_sheet_size': '28x40',
            'ups': '4',
            'purchase_sheet_size': '30x42',
            'purchase_sheet_ups': '2',
            'die_cutting': 'NO'
        })
        
        self.assertEqual(response.status_code, 302)
        
        # Verify SkuRecipe was updated and status set to pending_review
        recipe.refresh_from_db()
        self.assertEqual(recipe.size_w_mm, 250)
        self.assertEqual(recipe.size_h_mm, 351)
        self.assertEqual(recipe.print_sheet_size, '28x40')
        self.assertEqual(recipe.ups, 4)
        self.assertEqual(recipe.purchase_sheet_size, '30x42')
        self.assertEqual(recipe.purchase_sheet_ups, 2)
        self.assertEqual(recipe.plate_set_no, 'Set-A')
        self.assertEqual(recipe.awc_no, 'AWC-777')
        self.assertEqual(recipe.die_cutting, 'NO')
        # Plate ink chips must not overwrite production print color master field.
        self.assertEqual(recipe.color_spec, '2')
        self.assertNotEqual(recipe.color_spec, 'Cyan, Yellow')
        
        # Verify current PlanningJob has layout specs copied
        job.refresh_from_db()
        self.assertEqual(job.size_w_mm, 250)
        self.assertEqual(job.size_h_mm, 351)
        self.assertEqual(job.print_sheet_size, '28x40')
        self.assertEqual(job.color_spec, '2')
        self.assertEqual(job.total_colors, 2)
        self.assertNotEqual(job.color_spec, 'Cyan, Yellow')
        self.assertEqual(job.ups, 4.0)
        self.assertEqual(job.purchase_sheet_size, '30x42')
        self.assertEqual(job.purchase_sheet_ups, 2.0)
        self.assertEqual(job.plate_set_no, 'Set-A')
        self.assertEqual(job.remarks, 'Layout standard prep')
        self.assertEqual(recipe.notes, 'Layout standard prep')

    def test_plate_request_creation_copies_remarks(self):
        job = PlanningJob.objects.create(
            jc_number='JC-NEW-REMARKS',
            sku='SKU-NEW-REMARKS',
            planning_stage='new_plate_making',
            remarks='Original PO Remarks'
        )
        from printing_plates.services import create_or_get_plate_request_from_planning_job
        req = create_or_get_plate_request_from_planning_job(job, self.admin_user)
        self.assertEqual(req.remarks, 'Original PO Remarks')

    def test_plate_queue_view_renders_columns(self):
        self.planning_job.sku = 'SKU-QUEUE-001'
        self.planning_job.job_name = 'Queue Test Job'
        self.planning_job.color_spec = '4 color'
        self.planning_job.repeat_flag = 'New'
        self.planning_job.save(update_fields=['sku', 'job_name', 'color_spec', 'repeat_flag', 'updated_at'])
        self.jobcard.SKU = 'SKU-QUEUE-001'
        self.jobcard.save(update_fields=['SKU'])
        self.plate_request.progress = 'Layout in progress'
        self.plate_request.requested_at = timezone.now()
        self.plate_request.save(update_fields=['progress', 'requested_at', 'updated_at'])

        self.client.login(username='designer', password='password')
        response = self.client.get(reverse('printing_plates:queue'))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'plate-queue-table')
        self.assertContains(response, 'erp-plate-type-flag')
        self.assertContains(response, 'JC-0001')
        self.assertContains(response, 'SKU-QUEUE-001')
        self.assertContains(response, 'Queue Test Job')
        self.assertContains(response, 'Layout in progress')
        self.assertEqual(response.context['queue_count'], 1)
        self.assertEqual(response.context['plate_requests'][0].plate_request_type, 'New Artwork')

    def test_plate_request_list_renders_columns(self):
        self.plate_request.progress = 'Layout in progress'
        self.plate_request.requested_at = timezone.now()
        self.plate_request.save(update_fields=['progress', 'requested_at', 'updated_at'])

        self.client.login(username='designer', password='password')
        response = self.client.get(reverse('printing_plates:request_list'))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'plate-requests-table')
        self.assertContains(response, 'erp-plate-type-flag')
        self.assertContains(response, 'JC-0001')
        self.assertContains(response, 'Layout in progress')
        self.assertEqual(response.context['list_count'], 1)

    def test_plate_request_detail_renders_sections(self):
        self.client.login(username='designer', password='password')
        response = self.client.get(reverse('printing_plates:request_detail', kwargs={'pk': self.plate_request.pk}))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'plate-detail-layout')
        self.assertContains(response, 'Design & Layout Specifications')
        self.assertContains(response, 'Action Panel')
        self.assertContains(response, 'JC-0001')

    def test_request_plate_remake_requires_issued_plates(self):
        from django.core.exceptions import ValidationError
        from printing_plates.services import request_plate_remake

        self.plate_request.status = PlateRequest.STATUS_DRAFT
        self.plate_request.save(update_fields=['status'])
        self.planning_job.planning_stage = 'new_plate_making'
        self.planning_job.plate_set_no = ''
        self.planning_job.save(update_fields=['planning_stage', 'plate_set_no'])
        self.jobcard.plate_set_no = ''
        self.jobcard.save(update_fields=['plate_set_no'])
        with self.assertRaises(ValidationError):
            request_plate_remake(
                self.jobcard,
                actor=self.admin_user,
                reason=PlateRequest.REASON_DAMAGED_DURING_RUN,
                damaged_colors='Cyan',
                notes='Cracked',
            )

    def test_request_plate_remake_requires_damaged_colors(self):
        from django.core.exceptions import ValidationError
        from printing_plates.services import request_plate_remake

        self.plate_request.status = PlateRequest.STATUS_AVAILABLE
        self.plate_request.save(update_fields=['status'])
        with self.assertRaises(ValidationError):
            request_plate_remake(
                self.jobcard,
                actor=self.admin_user,
                reason=PlateRequest.REASON_DAMAGED_DURING_RUN,
                damaged_colors='',
                notes='Cracked',
            )

    def test_request_plate_remake_allows_legacy_plate_received_jobs(self):
        from printing_plates.services import request_plate_remake

        self.plate_request.delete()
        self.planning_job.planning_stage = 'plate_received'
        self.planning_job.plate_set_no = 'LEGACY-SET'
        self.planning_job.save(update_fields=['planning_stage', 'plate_set_no'])
        self.jobcard.plate_set_no = 'LEGACY-SET'
        self.jobcard.save(update_fields=['plate_set_no'])

        replacement = request_plate_remake(
            self.jobcard,
            actor=self.admin_user,
            reason=PlateRequest.REASON_DAMAGED_BEFORE_PRINTING,
            damaged_colors='Black',
            notes='Old job plates damaged',
        )
        self.assertEqual(replacement.set_no, 'LEGACY-SET')
        self.assertEqual(replacement.damaged_colors, 'Black')

    def test_request_plate_remake_creates_replacement_and_keeps_history(self):
        from printing_plates.services import (
            get_plate_remake_count,
            job_is_waiting_for_plates,
            request_plate_remake,
        )

        self.plate_request.status = PlateRequest.STATUS_AVAILABLE
        self.plate_request.set_no = 'SET-OLD'
        self.plate_request.awc_no = 'AWC-1'
        self.plate_request.save(update_fields=['status', 'set_no', 'awc_no'])

        replacement = request_plate_remake(
            self.jobcard,
            actor=self.admin_user,
            reason=PlateRequest.REASON_DAMAGED_DURING_RUN,
            damaged_colors='Cyan, Black',
            notes='Plate cracked on press',
        )

        self.plate_request.refresh_from_db()
        self.planning_job.refresh_from_db()

        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_ARCHIVED)
        self.assertEqual(replacement.source, PlateRequest.SOURCE_REPLACEMENT)
        self.assertEqual(replacement.replacement_reason, PlateRequest.REASON_DAMAGED_DURING_RUN)
        self.assertEqual(replacement.damaged_colors, 'Cyan, Black')
        self.assertEqual(replacement.replaces_request_id, self.plate_request.pk)
        self.assertEqual(replacement.awc_no, 'AWC-1')
        self.assertEqual(self.planning_job.planning_stage, 'repeat_plate_making')
        self.assertTrue(job_is_waiting_for_plates(self.jobcard))
        self.assertEqual(get_plate_remake_count(self.jobcard), 1)

    def test_replacement_filter_on_plate_request_list(self):
        self.plate_request.status = PlateRequest.STATUS_AVAILABLE
        self.plate_request.save(update_fields=['status'])
        from printing_plates.services import request_plate_remake
        request_plate_remake(
            self.jobcard,
            actor=self.admin_user,
            reason=PlateRequest.REASON_WORN_OUT,
            damaged_colors='Magenta',
            notes='Worn',
        )

        self.client.login(username='designer', password='password')
        response = self.client.get(reverse('printing_plates:request_list') + '?type=replacement')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context['request_type'], 'replacement')
        self.assertGreaterEqual(response.context['replacement_count'], 1)
        self.assertGreaterEqual(response.context['list_count'], 1)

    def test_awc_no_cannot_be_reused_on_different_sku(self):
        from planning.models import SkuRecipe

        SkuRecipe.objects.create(
            sku='SKU-OTHER',
            job_name='Other Design',
            awc_no='AWC-UNIQUE-1',
            master_data_status='approved',
            is_active=True,
        )
        self.planning_job.sku = 'SKU-001'
        self.planning_job.save(update_fields=['sku', 'updated_at'])
        self.jobcard.SKU = 'SKU-001'
        self.jobcard.save(update_fields=['SKU'])

        self.client.login(username='designer', password='password')
        url = reverse('printing_plates:request_action', kwargs={'pk': self.plate_request.pk})
        response = self.client.post(url, {
            'action': 'send_to_vendor',
            'vendor': 'Dot Max',
            'print_color': '4',
            'plate_color': 'Black',
            'awc_no': 'AWC-UNIQUE-1',
            'sets_required': '1',
            'plate_quantity': '1',
        })
        self.assertEqual(response.status_code, 302)
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_DRAFT)
        self.assertNotEqual(self.plate_request.awc_no, 'AWC-UNIQUE-1')

    def test_awc_no_can_be_reused_on_same_sku(self):
        from planning.models import SkuRecipe

        recipe = SkuRecipe.objects.create(
            sku='SKU-001',
            job_name='Same Design',
            awc_no='AWC-SAME-1',
            master_data_status='approved',
            is_active=True,
        )
        self.planning_job.sku = 'SKU-001'
        self.planning_job.save(update_fields=['sku', 'updated_at'])
        self.jobcard.SKU = 'SKU-001'
        self.jobcard.save(update_fields=['SKU'])
        self.plate_request.sku_recipe = recipe
        self.plate_request.save(update_fields=['sku_recipe'])

        self.client.login(username='designer', password='password')
        url = reverse('printing_plates:request_action', kwargs={'pk': self.plate_request.pk})
        response = self.client.post(url, {
            'action': 'send_to_vendor',
            'vendor': 'Dot Max',
            'print_color': '4',
            'plate_color': 'Black',
            'awc_no': 'AWC-SAME-1',
            'set_no': 'Set-1',
            'sets_required': '1',
            'plate_quantity': '1',
            'size_w_mm': '100',
            'size_h_mm': '200',
            'print_sheet_size': '28x40',
            'ups': '4',
            'purchase_sheet_size': '30x42',
            'purchase_sheet_ups': '2',
            'die_cutting': 'NO',
        })
        self.assertEqual(response.status_code, 302)
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_SENT)
        self.assertEqual(self.plate_request.awc_no, 'AWC-SAME-1')

    def test_send_to_vendor_blocked_when_designer_fields_missing(self):
        self.client.login(username='designer', password='password')
        url = reverse('printing_plates:request_action', kwargs={'pk': self.plate_request.pk})
        response = self.client.post(url, {
            'action': 'send_to_vendor',
            'vendor': 'Dot Max',
            'print_color': '4',
            'plate_color': 'Black',
            'set_no': 'Set-1',
            'awc_no': 'AWC-1',
            'size_w_mm': '100',
            'size_h_mm': '200',
            'print_sheet_size': '28x40',
            'ups': '4',
            'purchase_sheet_size': '30x42',
            'purchase_sheet_ups': '2',
            # die_cutting intentionally missing
        })
        self.assertEqual(response.status_code, 302)
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_DRAFT)

    def test_multi_set_quantity_suggested_from_impressions_and_inks(self):
        from printing_plates.plate_set_helpers import build_plate_set_suggestion, suggest_sets_required

        self.machine.plate_life_impressions = 25000
        self.machine.save(update_fields=['plate_life_impressions'])
        self.jobcard.total_impressions_required = 100000
        self.jobcard.save(update_fields=['total_impressions_required'])
        self.planning_job.color_spec = '4'
        self.planning_job.save(update_fields=['color_spec', 'updated_at'])

        self.assertEqual(suggest_sets_required(100000, 25000), 4)
        suggestion = build_plate_set_suggestion(self.plate_request, plate_color='Cyan, Magenta, Yellow, Black')
        self.assertEqual(suggestion['sets_required'], 4)
        self.assertEqual(suggestion['plate_quantity'], 16)

        self.client.login(username='designer', password='password')
        url = reverse('printing_plates:request_action', kwargs={'pk': self.plate_request.pk})
        response = self.client.post(url, {
            'action': 'send_to_vendor',
            'vendor': 'Dot Max',
            'print_color': '4',
            'plate_color': 'Cyan, Magenta, Yellow, Black',
            'awc_no': 'AWC-MULTI-1',
            'set_no': 'Set-M',
            'size_w_mm': '100',
            'size_h_mm': '200',
            'print_sheet_size': '28x40',
            'ups': '4',
            'purchase_sheet_size': '30x42',
            'purchase_sheet_ups': '2',
            'die_cutting': 'NO',
            # omit sets/quantity — server should suggest 4 sets × 4 inks = 16
        })
        self.assertEqual(response.status_code, 302)
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.sets_required, 4)
        self.assertEqual(self.plate_request.plate_quantity, 16)
        self.assertEqual(self.plate_request.plate_quantity_display, '16 (4 sets)')


class StalePlateRequestVisibilityTests(TestCase):
    def test_open_request_visible_after_job_moves_to_planning_done(self):
        from printing_plates.services import plate_request_active_queryset, plate_request_is_stale_open

        job = PlanningJob.objects.create(
            jc_number='JC-STALE-1',
            sku='SKU-STALE',
            status='released',
            planning_stage='planning_done',
            repeat_flag='New',
        )
        plate_request = PlateRequest.objects.create(
            planning_job=job,
            status=PlateRequest.STATUS_DRAFT,
        )
        self.assertTrue(plate_request_is_stale_open(plate_request))
        self.assertTrue(plate_request_active_queryset().filter(pk=plate_request.pk).exists())


class PlanningPlateRequestCancelTests(TestCase):
    def setUp(self):
        self.planner = User.objects.create_user(username='planner', password='password')
        self.planner.profile.role = 'planner'
        self.planner.profile.save()

        self.job = PlanningJob.objects.create(
            jc_number='JC-CANCEL-1',
            sku='SKU-CANCEL',
            status='released',
            planning_stage='planning_done',
            repeat_flag='New',
        )
        self.plate_request = PlateRequest.objects.create(
            planning_job=self.job,
            status=PlateRequest.STATUS_DRAFT,
        )

    def test_planner_can_cancel_open_plate_request_from_planning(self):
        self.client.login(username='planner', password='password')
        url = reverse('planning:job_plate_request_cancel', args=[self.job.id])
        response = self.client.post(url, {
            'plate_request_id': self.plate_request.pk,
            'cancel_reason': 'Plates already issued outside workflow',
        })
        self.assertEqual(response.status_code, 302)
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_ARCHIVED)
        self.assertTrue(self.plate_request.is_cancelled)

    def test_stale_open_plate_requests_queryset_matches_helper(self):
        from printing_plates.services import plate_request_is_stale_open, stale_open_plate_requests_queryset

        qs_ids = set(stale_open_plate_requests_queryset().values_list('pk', flat=True))
        self.assertIn(self.plate_request.pk, qs_ids)
        self.assertTrue(plate_request_is_stale_open(self.plate_request))


class PlateMakingPreventionTests(TestCase):
    def setUp(self):
        self.admin = User.objects.create_user(username='admin', password='password')
        self.admin.profile.role = 'admin'
        self.admin.profile.save()

        self.job = PlanningJob.objects.create(
            jc_number='JC-BLOCK-1',
            sku='SKU-BLOCK',
            status='released',
            planning_stage='planning_done',
            repeat_flag='New',
            material='Paper',
            application='Label',
            machine_name='KBA',
            plate_set_no='SET-999',
        )

    def test_planning_stage_update_blocked_for_released_job(self):
        self.client.login(username='admin', password='password')
        response = self.client.post(reverse('planning:jobs'), {
            'action': 'update_planning_stage',
            'job_id': self.job.id,
            'planning_stage': 'plate_making',
        })
        self.assertEqual(response.status_code, 302)
        self.job.refresh_from_db()
        self.assertEqual(self.job.planning_stage, 'planning_done')

    def test_bulk_cancel_stale_archives_open_requests(self):
        plate_request = PlateRequest.objects.create(
            planning_job=self.job,
            status=PlateRequest.STATUS_DRAFT,
        )
        from printing_plates.services import bulk_cancel_stale_open_plate_requests

        result = bulk_cancel_stale_open_plate_requests(actor=self.admin, dry_run=False)
        self.assertGreaterEqual(result['cancelled'], 1)
        plate_request.refresh_from_db()
        self.assertEqual(plate_request.status, PlateRequest.STATUS_ARCHIVED)
        self.assertTrue(plate_request.is_cancelled)


class ReleaseBlockedByOpenPlateRequestTests(TestCase):
    def setUp(self):
        self.admin = User.objects.create_user(username='adminrel', password='password')
        self.admin.profile.role = 'admin'
        self.admin.profile.save()

        self.planning_job = PlanningJob.objects.create(
            jc_number='JC-REL-BLOCK',
            sku='SKU-REL-BLOCK',
            status='qc_approved',
            planning_stage='new_plate_making',
            repeat_flag='New',
        )
        self.machine = Machine.objects.create(name='KBA Block Test')
        self.department = Department.objects.create(name='Offset Block Test')
        self.material = Material.objects.create(name='Art Paper Block')
        self.job_card = JobCard.objects.create(
            job_card_no='JC-REL-BLOCK',
            planning_job=self.planning_job,
            SKU='SKU-REL-BLOCK',
            order_qty=5000,
            total_impressions_required=5000,
            total_sheet_quantity=500,
            total_colors=4,
            plate_set_no='SET-BLOCK',
            po_date=timezone.now().date(),
            machine_name=self.machine,
            department=self.department,
            material=self.material,
            status='production_approved',
        )
        self.plate_request = PlateRequest.objects.create(
            planning_job=self.planning_job,
            job_card=self.job_card,
            status=PlateRequest.STATUS_DRAFT,
            source=PlateRequest.SOURCE_PLANNING,
        )

    def test_release_blocked_while_open_plate_request_exists(self):
        from django.core.exceptions import ValidationError
        from core.jobcard_service import release_to_production

        with self.assertRaises(ValidationError):
            release_to_production(self.job_card, actor=self.admin)

        self.job_card.refresh_from_db()
        self.assertEqual(self.job_card.workflow_status, 'production_approved')
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_DRAFT)

    def test_release_allowed_after_plate_request_cancelled(self):
        from core.jobcard_service import release_to_production
        from printing_plates.services import cancel_plate_request

        cancel_plate_request(
            self.plate_request,
            actor=self.admin,
            reason='Plates issued outside workflow',
        )
        release_to_production(self.job_card, actor=self.admin)
        self.job_card.refresh_from_db()
        self.assertEqual(self.job_card.workflow_status, 'released')
