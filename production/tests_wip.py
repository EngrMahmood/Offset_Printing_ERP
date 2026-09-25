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
            print_passes=1,
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
        # Setup multi-pass
        self.planning_job.print_passes = 2
        self.planning_job.save()
        self.job_card.total_impressions_required = 2000
        self.job_card.save()

        # 1. Start printing (intermediate pass, output_sheets=0 is mandatory)
        Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=100,
            output_sheets=0,
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

    def test_stock_fully_covered_job_shows_ready_for_dispatch(self):
        """A print job with zero printing/packing activity but stock_qty
        covering the whole order must not fall through to 'Printing' just
        because workflow_status is 'released' — see JC-09-26-PP-2362."""
        self.planning_job.stock_qty = 1000
        self.planning_job.save(update_fields=['stock_qty'])

        self.assertEqual(self.job_card.stock_covered_qty, 1000)
        self.assertEqual(get_system_calculated_status_name(self.job_card), 'Ready for Dispatch')

    def test_stock_partially_covering_job_still_shows_printing(self):
        """Partial stock coverage (less than order_qty) must not short-circuit
        the normal Printing status — only full coverage does."""
        self.planning_job.stock_qty = 400
        self.planning_job.save(update_fields=['stock_qty'])

        self.assertEqual(get_system_calculated_status_name(self.job_card), 'Printing')

    def test_printing_production_transition(self):
        """Creating an intermediate printing production record transitions WIP status to Printing."""
        # Setup multi-pass
        self.planning_job.print_passes = 2
        self.planning_job.save()
        self.job_card.total_impressions_required = 2000
        self.job_card.save()

        prod = Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=100,
            output_sheets=0,
            print_pass_number=1,
            created_by=self.user
        )
        
        wip_status = JobCardWipStatus.objects.filter(job_card=self.job_card).first()
        self.assertIsNotNone(wip_status)
        self.assertEqual(wip_status.status.name, 'Printing')
        self.assertFalse(wip_status.is_manual)

    def test_printing_completed_transition(self):
        """Logging the final print pass transitions WIP to Printing Completed."""
        # Setup multi-pass
        self.planning_job.print_passes = 2
        self.planning_job.save()
        self.job_card.total_impressions_required = 2000
        self.job_card.save()

        # Log first pass (intermediate, output_sheets=0)
        Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=1000,
            output_sheets=0,
            print_pass_number=1,
            created_by=self.user
        )
        self.assertEqual(get_system_calculated_status_name(self.job_card), 'Printing')

        # Log second pass (final, output_sheets > 0 allowed)
        Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=1000,
            output_sheets=1000,
            print_pass_number=2,
            created_by=self.user
        )
        self.assertEqual(get_system_calculated_status_name(self.job_card), 'Printing Completed')

    def test_final_pass_started_but_short_of_order_qty_is_partial_printing(self):
        """Logging SOME output on the final pass doesn't mean the run is
        done — an operator can log a partial entry on the final pass before
        reaching the order quantity. Must show 'Partial Printing', not
        'Printing Completed', until produced pcs actually reach order_qty."""
        self.planning_job.print_passes = 2
        self.planning_job.save()
        self.job_card.total_impressions_required = 2000
        self.job_card.save()

        Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=1000,
            output_sheets=0,
            print_pass_number=1,
            created_by=self.user
        )
        # Final pass started, but only 400 of the 1000-pcs order produced so far.
        Production.objects.create(
            entry_type='printing',
            job_card=self.job_card,
            machine=self.machine,
            operator=self.operator,
            shift='A',
            date=date(2026, 1, 1),
            impressions=400,
            output_sheets=400,
            print_pass_number=2,
            created_by=self.user
        )
        self.assertEqual(get_system_calculated_status_name(self.job_card), 'Partial Printing')

    def test_packing_production_transition(self):
        """Creating a packing record transitions WIP to Sorting / Packing."""
        # Print first (1 pass is final, output_sheets > 0 is allowed)
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
        # Print first
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
        # Print first
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
        
        # Setup packing first
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


class ProductionWipPageRenderTests(TestCase):
    """Exercises the actual /production-wip/ view+template (not just
    wip_service in isolation) to verify the Supervisor Status dash-unless-
    manual behaviour, the Partial Printing badge, and the JC hyperlink."""

    def setUp(self):
        from core.models import Permission, UserPermissionOverride

        User = get_user_model()
        self.user = User.objects.create_user(username='wip_page_user', password='pass')
        self.profile, _ = UserProfile.objects.get_or_create(user=self.user)
        self.profile.role = 'admin'
        self.profile.save(update_fields=['role'])
        # Fresh test DB has no seeded Role/Permission rows (they come from
        # the seed_access_control management command, not a migration) —
        # grant the one permission this view needs directly, same pattern
        # as core.tests.JobCardFinalizationSetStockViewTests.
        permission, _ = Permission.objects.get_or_create(
            code='action.view_production_wip', defaults={'name': 'View Production WIP'},
        )
        UserPermissionOverride.objects.get_or_create(
            user=self.user, permission=permission, defaults={'granted': True},
        )
        self.client.force_login(self.user)

        self.machine = Machine.objects.create(name='WIP Page Test Machine')
        self.planning_job = PlanningJob.objects.create(
            jc_number='JC-WIPPAGE-001', order_qty=1000, ups=1, status='in_production',
            plan_date=date(2026, 1, 1), plan_month='January 2026', print_passes=2,
        )
        self.job_card = JobCard.objects.create(
            job_card_no='JC-WIPPAGE-001', planning_job=self.planning_job, order_qty=1000, ups=1,
            SKU='SKU-WIPPAGE-001', is_print_job=True, total_sheet_quantity=1000, total_colors=4,
            status='in_production', po_date=date(2026, 1, 1), total_impressions_required=2000,
            machine_name=self.machine,
        )

    def test_supervisor_status_blank_when_auto_manual_when_overridden(self):
        # Give the job SOME loggable activity so a WIP status row actually
        # gets auto-created on first render (a job with nothing logged yet
        # computes 'Not Set', which never gets a stored row at all — see
        # evaluate_and_update_job_wip_status).
        Production.objects.create(
            entry_type='printing', job_card=self.job_card, machine=self.machine,
            date=date(2026, 1, 1), shift='A', impressions=100, output_sheets=0,
            print_pass_number=1, created_by=self.user,
        )

        # Auto (never manually touched) — should render as a dash, not "Printing".
        response = self.client.get('/production-wip/')
        self.assertEqual(response.status_code, 200)
        html = response.content.decode()
        self.assertIn('JC-WIPPAGE-001', html)

        wip_status = JobCardWipStatus.objects.get(job_card=self.job_card)
        self.assertFalse(wip_status.is_manual)

        # Manual override — should now show the status name + "Manual" badge.
        update_wip_status_for_job(self.job_card, 'Ready for Dispatch', user=self.user, is_manual=True)
        response = self.client.get('/production-wip/')
        html = response.content.decode()
        self.assertIn('Manual', html)

    def test_page_does_not_crash_for_job_with_no_wip_status_row_at_all(self):
        """Regression: a job card with nothing logged yet computes 'Not Set',
        which never gets a JobCardWipStatus row created (see
        evaluate_and_update_job_wip_status). The view used to crash on such
        a row with `getattr(job.production_wip_status, 'is_manual', False)`
        — that only catches AttributeError, not Django's
        RelatedObjectDoesNotExist raised by a missing reverse OneToOne."""
        self.assertFalse(JobCardWipStatus.objects.filter(job_card=self.job_card).exists())

        response = self.client.get('/production-wip/')

        self.assertEqual(response.status_code, 200)
        html = response.content.decode()
        self.assertIn('JC-WIPPAGE-001', html)

    def test_jc_number_links_to_history_report(self):
        response = self.client.get('/production-wip/')
        html = response.content.decode()
        self.assertIn(f'/planning/job/{self.planning_job.id}/history-report/pdf/', html)

    def test_partial_printing_badge_renders(self):
        Production.objects.create(
            entry_type='printing', job_card=self.job_card, machine=self.machine,
            date=date(2026, 1, 1), shift='A', impressions=1000, output_sheets=0,
            print_pass_number=1, created_by=self.user,
        )
        Production.objects.create(
            entry_type='printing', job_card=self.job_card, machine=self.machine,
            date=date(2026, 1, 1), shift='A', impressions=400, output_sheets=400,
            print_pass_number=2, created_by=self.user,
        )
        response = self.client.get('/production-wip/')
        html = response.content.decode()
        self.assertIn('Partial Printing', html)

    def test_partial_printing_is_a_selectable_manual_status(self):
        """Regression: 'Partial Printing' was added as a calculated-status
        value but the master ProductionWipStatus list (which drives the
        Supervisor Status filter and the per-row manual override dropdown)
        was never given a matching entry, so supervisors had no way to
        manually set a job to it."""
        response = self.client.get('/production-wip/')
        self.assertEqual(response.status_code, 200)
        status = ProductionWipStatus.objects.get(name='Partial Printing')
        html = response.content.decode()
        self.assertIn(f'<option value="{status.id}"', html)
