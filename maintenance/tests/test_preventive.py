from datetime import date, timedelta
from io import StringIO

from django.core.management import call_command
from django.urls import reverse
from django.utils import timezone

from maintenance.models import MaintenanceActivityLog, MaintenanceRecord, PreventiveMaintenancePlan
from maintenance.services import (
    generate_due_pm_records, machine_impression_total, pm_plan_status, remind_overdue_pm,
)
from maintenance.workflow import TransitionError, transition_record

from .base import MaintenanceTestCase


def log_impressions(machine, count):
    """Record ``count`` impressions against the machine as a Production entry (bypassing job-card plumbing)."""
    from core.models import JobCard, Production

    card, _ = JobCard.objects.get_or_create(
        job_card_no='JC-MNT-TEST', defaults=dict(
            SKU='TEST', order_qty=1000, ups=1, total_sheet_quantity=1000, total_impressions_required=1000000,
            total_colors=4, wastage=0, po_date=timezone.localdate(),
        ),
    )
    Production.objects.bulk_create([Production(
        job_card=card, machine=machine, entry_type='printing', impressions=count, is_active=True,
        date=timezone.localdate(),
    )])


class PMGenerationTests(MaintenanceTestCase):
    def make_plan(self, **overrides):
        values = dict(
            machine=self.machine, title='Monthly lubrication', interval_type='DAYS', interval_value=30,
            next_due_at=timezone.localdate() - timedelta(days=1), checklist='Clean rollers\nCheck guards',
            instructions='Follow the machine manual.', assigned_to=self.engineer,
        )
        values.update(overrides)
        return PreventiveMaintenancePlan.objects.create(**values)

    def test_due_plan_creates_one_ticket_with_its_checklist(self):
        plan = self.make_plan()
        result = generate_due_pm_records(actor=self.manager)
        self.assertEqual(len(result['created']), 1)
        record = result['created'][0]
        self.assertEqual((record.maintenance_type, record.status, record.pm_plan), ('PREVENTIVE', 'DIAGNOSED', plan))
        self.assertEqual(record.assigned_to, self.engineer)
        self.assertEqual(list(record.checklist_items.values_list('text', flat=True)), ['Clean rollers', 'Check guards'])

    def test_running_repeatedly_does_not_duplicate_tickets(self):
        self.make_plan()
        for _ in range(3):
            generate_due_pm_records(actor=self.manager)
        self.assertEqual(MaintenanceRecord.objects.filter(pm_plan__isnull=False).count(), 1)

    def test_a_plan_that_is_not_due_creates_nothing(self):
        self.make_plan(next_due_at=timezone.localdate() + timedelta(days=5))
        result = generate_due_pm_records(actor=self.manager)
        self.assertEqual((len(result['created']), result['skipped']), (0, 1))

    def test_next_due_date_advances_from_the_due_date_past_today(self):
        today = timezone.localdate()
        plan = self.make_plan(next_due_at=today - timedelta(days=65), interval_value=30)
        generate_due_pm_records(actor=self.manager)
        plan.refresh_from_db()
        self.assertEqual(plan.next_due_at, today - timedelta(days=65) + timedelta(days=90))
        self.assertGreater(plan.next_due_at, today)

    def test_same_occurrence_cannot_be_created_twice_even_if_state_is_rewound(self):
        plan = self.make_plan()
        generate_due_pm_records(actor=self.manager)
        due = (timezone.localdate() - timedelta(days=1))
        PreventiveMaintenancePlan.objects.filter(pk=plan.pk).update(next_due_at=due)
        result = generate_due_pm_records(actor=self.manager)
        self.assertEqual(len(result['created']), 0)
        self.assertEqual(MaintenanceRecord.objects.filter(pm_plan=plan).count(), 1)

    def test_disabled_plan_is_ignored(self):
        self.make_plan(is_active=False)
        self.assertEqual(len(generate_due_pm_records(actor=self.manager)['created']), 0)

    def test_one_failing_plan_does_not_stop_the_others(self):
        good = self.make_plan(title='Good plan')
        bad = self.make_plan(title='Bad plan', machine=self.other_machine)
        original = PreventiveMaintenancePlan.save

        def flaky(plan, *args, **kwargs):
            if plan.pk == bad.pk:
                raise RuntimeError('boom')
            return original(plan, *args, **kwargs)

        PreventiveMaintenancePlan.save = flaky
        try:
            result = generate_due_pm_records(actor=self.manager)
        finally:
            PreventiveMaintenancePlan.save = original
        self.assertEqual([r.pm_plan for r in result['created']], [good])
        self.assertEqual([p.pk for p, _ in result['failed']], [bad.pk])
        self.assertFalse(MaintenanceRecord.objects.filter(pm_plan=bad).exists())   # rolled back

    def test_management_command_reports_a_summary(self):
        self.manager.is_superuser = True
        self.manager.save()
        self.make_plan()
        out = StringIO()
        call_command('generate_due_pm_records', stdout=out)
        self.assertIn('1 ticket(s) created', out.getvalue())
        out = StringIO()
        call_command('generate_due_pm_records', stdout=out)
        self.assertIn('0 ticket(s) created', out.getvalue())


class ImpressionPlanTests(MaintenanceTestCase):
    def make_plan(self, **overrides):
        values = dict(machine=self.machine, title='Every 100k', interval_type='IMPRESSIONS', interval_value=100000)
        values.update(overrides)
        return PreventiveMaintenancePlan.objects.create(**values)

    def test_counter_is_the_sum_of_logged_production_impressions(self):
        log_impressions(self.machine, 40000)
        log_impressions(self.machine, 15000)
        log_impressions(self.other_machine, 999)
        self.assertEqual(machine_impression_total(self.machine), 55000)

    def test_no_baseline_means_no_invented_ticket(self):
        log_impressions(self.machine, 250000)
        plan = self.make_plan()
        result = generate_due_pm_records(actor=self.manager)
        plan.refresh_from_db()
        self.assertEqual(len(result['created']), 0)
        self.assertEqual(plan.next_due_impressions, 350000)      # baseline = current total + interval

    def test_due_when_threshold_reached_once_per_occurrence(self):
        log_impressions(self.machine, 10000)
        plan = self.make_plan(next_due_impressions=100000)
        self.assertEqual(len(generate_due_pm_records(actor=self.manager)['created']), 0)
        log_impressions(self.machine, 95000)                      # total 105,000
        self.assertEqual(len(generate_due_pm_records(actor=self.manager)['created']), 1)
        self.assertEqual(len(generate_due_pm_records(actor=self.manager)['created']), 0)
        plan.refresh_from_db()
        self.assertEqual(plan.next_due_impressions, 200000)

    def test_plan_status_labels(self):
        today = date(2026, 6, 15)
        plan = self.make_plan(next_due_impressions=100000)
        self.assertEqual(pm_plan_status(plan, today, 120000)[0], 'overdue')
        self.assertEqual(pm_plan_status(plan, today, 95000)[0], 'soon')
        self.assertEqual(pm_plan_status(plan, today, 10000)[0], 'upcoming')
        days = PreventiveMaintenancePlan.objects.create(
            machine=self.machine, title='d', interval_type='DAYS', interval_value=7, next_due_at=today + timedelta(days=3),
        )
        self.assertEqual(pm_plan_status(days, today)[0], 'soon')
        days.next_due_at = today - timedelta(days=2)
        self.assertEqual(pm_plan_status(days, today)[0], 'overdue')


class PMTicketWorkflowTests(MaintenanceTestCase):
    def raise_ticket(self, **plan_values):
        plan = PreventiveMaintenancePlan.objects.create(
            machine=self.machine, title='Weekly', interval_type='DAYS', interval_value=7,
            next_due_at=timezone.localdate(), checklist='Check belts\nCheck guards', **plan_values,
        )
        return generate_due_pm_records(actor=self.manager)['created'][0], plan

    def test_checklist_must_be_finished_before_completion(self):
        record, _plan = self.raise_ticket()
        record = transition_record(record, 'IN_PROGRESS', self.engineer)
        record.repair_action = 'Inspected'
        record.save()
        with self.assertRaises(TransitionError):
            transition_record(record, 'COMPLETED', self.engineer)
        record.checklist_items.update(is_done=True)
        self.assertEqual(transition_record(record, 'COMPLETED', self.engineer).status, 'COMPLETED')

    def test_pm_needs_no_root_cause(self):
        record, _plan = self.raise_ticket()
        record = transition_record(record, 'IN_PROGRESS', self.engineer)
        record.checklist_items.update(is_done=True)
        record.repair_action = 'Inspected'
        record.save()
        self.assertEqual(transition_record(record, 'COMPLETED', self.engineer).status, 'COMPLETED')

    def test_verification_records_last_done_on_the_plan(self):
        record, plan = self.raise_ticket()
        record = transition_record(record, 'IN_PROGRESS', self.engineer)
        record.checklist_items.update(is_done=True)
        record.repair_action, record.test_result = 'Inspected', 'PASS'
        record.save()
        record = transition_record(record, 'COMPLETED', self.engineer)
        transition_record(record, 'VERIFIED', self.engineer)
        plan.refresh_from_db()
        self.assertEqual(plan.last_done_at, timezone.localdate())

    def test_checklist_results_are_saved_with_who_and_when(self):
        record, _plan = self.raise_ticket()
        item = record.checklist_items.first()
        self.client.force_login(self.engineer)
        self.client.post(reverse('maintenance:record_checklist_update', args=[record.pk]), {
            f'done_{item.pk}': 'on', f'finding_{item.pk}': 'Belt worn, replaced',
        })
        item.refresh_from_db()
        self.assertTrue(item.is_done)
        self.assertEqual((item.finding, item.done_by), ('Belt worn, replaced', self.engineer))

    def test_operator_cannot_edit_the_checklist(self):
        record, _plan = self.raise_ticket()
        item = record.checklist_items.first()
        self.client.force_login(self.operator)
        response = self.client.post(reverse('maintenance:record_checklist_update', args=[record.pk]), {f'done_{item.pk}': 'on'})
        self.assertEqual(response.status_code, 403)
        item.refresh_from_db()
        self.assertFalse(item.is_done)


class OverdueReminderTests(MaintenanceTestCase):
    def make_overdue(self):
        plan = PreventiveMaintenancePlan.objects.create(
            machine=self.machine, title='Weekly', interval_type='DAYS', interval_value=7,
            next_due_at=timezone.localdate() - timedelta(days=3), assigned_to=self.engineer,
        )
        return generate_due_pm_records(actor=self.manager)['created'][0], plan

    def test_reminder_is_sent_once_per_week_per_ticket(self):
        record, _plan = self.make_overdue()
        later = timezone.localdate() + timedelta(days=1)
        self.assertEqual(len(remind_overdue_pm(as_of=later, actor=self.manager)), 1)
        self.assertEqual(len(remind_overdue_pm(as_of=later, actor=self.manager)), 0)
        self.assertEqual(MaintenanceActivityLog.objects.filter(record=record, action='Overdue reminder').count(), 1)

    def test_closed_tickets_are_not_reminded(self):
        record, _plan = self.make_overdue()
        MaintenanceRecord.objects.filter(pk=record.pk).update(status='CLOSED')
        self.assertEqual(len(remind_overdue_pm(as_of=timezone.localdate() + timedelta(days=1), actor=self.manager)), 0)


class PMScreenPermissionTests(MaintenanceTestCase):
    def test_engineer_can_create_edit_and_toggle_plans(self):
        self.client.force_login(self.engineer)
        response = self.client.post(reverse('maintenance:pm_plan_create'), {
            'machine': self.machine.pk, 'title': 'Weekly check', 'interval_type': 'DAYS', 'interval_value': 7,
            'checklist': 'Check A\nCheck B', 'is_active': 'on',
        })
        self.assertEqual(response.status_code, 302)
        plan = PreventiveMaintenancePlan.objects.get()
        self.assertEqual(plan.next_due_at, timezone.localdate() + timedelta(days=7))     # defaulted
        self.client.post(reverse('maintenance:pm_plan_toggle', args=[plan.pk]))
        plan.refresh_from_db()
        self.assertFalse(plan.is_active)
        self.assertEqual(self.client.get(reverse('maintenance:pm_plan_list')).status_code, 200)
        self.assertEqual(self.client.get(reverse('maintenance:pm_plan_edit', args=[plan.pk])).status_code, 200)

    def test_impression_plan_gets_a_baseline_on_save(self):
        log_impressions(self.machine, 5000)
        self.client.force_login(self.engineer)
        self.client.post(reverse('maintenance:pm_plan_create'), {
            'machine': self.machine.pk, 'title': 'Run based', 'interval_type': 'IMPRESSIONS', 'interval_value': 100000,
            'is_active': 'on',
        })
        self.assertEqual(PreventiveMaintenancePlan.objects.get().next_due_impressions, 105000)

    def test_operator_cannot_touch_plans(self):
        plan = PreventiveMaintenancePlan.objects.create(
            machine=self.machine, title='x', interval_type='DAYS', interval_value=7, next_due_at=timezone.localdate(),
        )
        self.client.force_login(self.operator)
        for name in ('maintenance:pm_plan_edit', 'maintenance:pm_plan_toggle'):
            self.assertEqual(self.client.post(reverse(name, args=[plan.pk])).status_code, 403)
        self.assertEqual(self.client.post(reverse('maintenance:pm_plan_create'), {}).status_code, 403)
