from datetime import date
from django.contrib.auth import get_user_model
from django.test import TestCase
from core.models import JobCard, Machine, Operator, Sorter, Production, Dispatch, JobCardWipStatus, ProductionWipStatus, ChangeLog, UserProfile
from planning.models import PlanningJob
from production.wip_service import evaluate_and_update_job_wip_status, update_wip_status_for_job, get_system_calculated_status_name

class WipAutomationTests(TestCase):
    def setUp(self):
        User = get_user_model()
        self.user = User.objects.create_user(username='wip_test_user', password='pass')
        self.profile, _ = UserProfile.objects.get_or_create(user=self.user)
        self.profile.role = 'admin'
        self.profile.save(update_fields=['role'])
        
        self.machine = Machine.objects.create(name='Test Machine', standard_impressions_per_hour=1000)
        self.operator = Operator.objects.create(name='Test Operator', is_active=True)
        self.sorter = Sorter.objects.create(name='Test Sorter', is_active=True)
        
        self.planning_job = PlanningJob.objects.create(
            jc_number='JC-WIP-001',
            po_number='PO-WIP-001',
            sku='SKU-WIP-001',
            status='released',
            planning_stage='planning_done',
            machine_name=self.machine.name,
            front_pass=1,
            back_pass=0,
        )
        
        self.job_card = JobCard.objects.create(
            job_card_no='JC-WIP-001',
            planning_job=self.planning_job,
            SKU='SKU-WIP-001',
            order_qty=1000,
            ups=1,
            is_print_job=True,
            status='released',
            po_date=date(2026, 1, 1),
            total_impressions_required=1000,
            total_sheet_quantity=1000,
            total_colors=4,
            machine_name=self.machine,
        )

    def test_calculated_vs_manual_status(self):
        """Dynamic system status correctly computes from logs without overwriting manual status."""
        # 1. Start printing
        Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=100,
            output_sheets=100,
            print_pass_number=1,
            created_by=self.user
        )
        
        # System calculates 'Printing'
        self.assertEqual(get_system_calculated_status_name(self.job_card), 'Printing')
        
        # 2. Supervisor sets status to 'Ready for Dispatch' manually
        update_wip_status_for_job(self.job_card, 'Ready for Dispatch', user=self.user, is_manual=True)
        
        wip_status = JobCardWipStatus.objects.get(job_card=self.job_card)
        self.assertEqual(wip_status.status.name, 'Ready for Dispatch')
        self.assertTrue(wip_status.is_manual)
        
        # Auto evaluation should not overwrite the manual status
        evaluate_and_update_job_wip_status(self.job_card)
        wip_status.refresh_from_db()
        self.assertEqual(wip_status.status.name, 'Ready for Dispatch')
        
        # But system calculated status still correctly reflects 'Printing' based on actual logs
        self.assertEqual(get_system_calculated_status_name(self.job_card), 'Printing')

    def test_printing_production_transition(self):
        """Creating a printing production record transitions WIP status to Printing."""
        prod = Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=100,
            output_sheets=100,
            print_pass_number=1,
            created_by=self.user
        )
        
        wip_status = JobCardWipStatus.objects.filter(job_card=self.job_card).first()
        self.assertIsNotNone(wip_status)
        self.assertEqual(wip_status.status.name, 'Printing')
        self.assertFalse(wip_status.is_manual)

    def test_packing_production_transition(self):
        """Creating a packing record transitions WIP to Sorting / Packing."""
        # Print first so there is a packing limit allowance
        Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=1000,
            output_sheets=1000,
            print_pass_number=1,
            created_by=self.user
        )
        
        prod = Production.objects.create(
            entry_type='packing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            sorter=self.sorter,
            shift='A',
            date=date(2026, 1, 1),
            packing_qty=100,
            created_by=self.user
        )
        
        wip_status = JobCardWipStatus.objects.filter(job_card=self.job_card).first()
        self.assertEqual(wip_status.status.name, 'Sorting / Packing')

    def test_packing_completed_transition(self):
        """When packed quantity >= order quantity, status becomes Ready for Dispatch."""
        # Print first so there is a packing limit allowance
        Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=1000,
            output_sheets=1000,
            print_pass_number=1,
            created_by=self.user
        )
        
        prod = Production.objects.create(
            entry_type='packing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            sorter=self.sorter,
            shift='A',
            date=date(2026, 1, 1),
            packing_qty=1000,
            created_by=self.user
        )
        
        wip_status = JobCardWipStatus.objects.filter(job_card=self.job_card).first()
        self.assertEqual(wip_status.status.name, 'Ready for Dispatch')

    def test_dispatch_transitions(self):
        """Dispatch logs trigger transition to Partial Dispatch and then Completed."""
        # Print first so there is a packing limit allowance
        Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=1000,
            output_sheets=1000,
            print_pass_number=1,
            created_by=self.user
        )
        
        # Setup packing first to allow dispatching up to order qty
        Production.objects.create(
            entry_type='packing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            sorter=self.sorter,
            shift='A',
            date=date(2026, 1, 1),
            packing_qty=1000,
            created_by=self.user
        )
        
        # Transition job card to in_production to allow dispatch
        self.job_card.status = 'in_production'
        self.job_card.save(update_fields=['status'])
        
        # Log partial dispatch
        disp1 = Dispatch.objects.create(
            job_card=self.job_card,
            dc_no='DC-1',
            dispatch_date=date(2026, 1, 2),
            dispatch_qty=400,
            created_by=self.user
        )
        wip_status = JobCardWipStatus.objects.filter(job_card=self.job_card).first()
        self.assertEqual(wip_status.status.name, 'Partial Dispatch')
        
        # Log remaining dispatch
        disp2 = Dispatch.objects.create(
            job_card=self.job_card,
            dc_no='DC-2',
            dispatch_date=date(2026, 1, 2),
            dispatch_qty=600,
            created_by=self.user
        )
        wip_status.refresh_from_db()
        self.assertEqual(wip_status.status.name, 'Completed')
