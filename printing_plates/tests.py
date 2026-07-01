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
            'plate_color': 'Black, Special',
            'plate_quantity': '4',
            'remarks': 'First print plates'
        })
        self.assertEqual(response.status_code, 302)
        self.plate_request.refresh_from_db()
        self.assertEqual(self.plate_request.status, PlateRequest.STATUS_SENT)
        self.assertEqual(self.plate_request.vendor, 'Dot Max')
        self.assertEqual(self.plate_request.set_no, 'Set-01')
        self.assertEqual(self.plate_request.new_set_no, 'New-Set-01')
        self.assertEqual(self.plate_request.awc_no, '973')
        self.assertEqual(self.plate_request.plate_color, 'Black, Special')
        self.assertEqual(self.plate_request.plate_quantity, 4)
        self.assertEqual(self.plate_request.sent_by, self.designer_user)
        self.assertIsNotNone(self.plate_request.sent_at)
        
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
            'plate_color': 'Cyan, Yellow',
            'plate_quantity': '2',
            'remarks': 'Layout standard prep',
            
            # Designer layout specifications
            'size_w_mm': '250.50',
            'size_h_mm': '350.75',
            'print_sheet_size': '28x40',
            'ups': '4',
            'purchase_sheet_size': '30x42',
            'purchase_sheet_ups': '2',
            'die_cutting': 'Die-Standard'
        })
        
        self.assertEqual(response.status_code, 302)
        
        # Verify SkuRecipe was updated and status set to pending_review
        recipe.refresh_from_db()
        self.assertEqual(recipe.size_w_mm, 250.50)
        self.assertEqual(recipe.size_h_mm, 350.75)
        self.assertEqual(recipe.print_sheet_size, '28x40')
        self.assertEqual(recipe.ups, 4.0)
        self.assertEqual(recipe.purchase_sheet_size, '30x42')
        self.assertEqual(recipe.purchase_sheet_ups, 2.0)
        self.assertEqual(recipe.color_spec, 'Cyan, Yellow')
        self.assertEqual(recipe.awc_no, 'AWC-777')
        self.assertEqual(recipe.die_cutting, 'Die-Standard')
        self.assertEqual(recipe.master_data_status, 'pending_review')
        
        # Verify current PlanningJob has layout specs copied
        job.refresh_from_db()
        self.assertEqual(job.size_w_mm, 250.50)
        self.assertEqual(job.size_h_mm, 350.75)
        self.assertEqual(job.print_sheet_size, '28x40')
        self.assertEqual(job.ups, 4.0)
        self.assertEqual(job.purchase_sheet_size, '30x42')
        self.assertEqual(job.purchase_sheet_ups, 2.0)
        self.assertEqual(job.color_spec, 'Cyan, Yellow')
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

