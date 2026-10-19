from datetime import timedelta
from io import StringIO

from django.core.management import call_command
from django.urls import reverse
from django.utils import timezone

from core.models import Permission, Role
from maintenance.models import FaultCategory, MachineDowntime, MaintenanceActivityLog, MaintenanceSparePart
from maintenance.reporting import build_report
from maintenance.workflow import transition_record
from supply_chain.models import ItemRequest, ItemRequestType

from .base import MaintenanceTestCase


class ReportTests(MaintenanceTestCase):
    def setUp(self):
        super().setUp()
        self.today = timezone.localdate()

    def breakdown(self, machine=None, hours=2, days_ago=3, status='CLOSED', category=None):
        record = self.make_record(machine=machine, status=status, maintenance_type='BREAKDOWN', stopped=True)
        MaintenanceRecord_update = type(record).objects.filter(pk=record.pk)
        MaintenanceRecord_update.update(reported_date=self.today - timedelta(days=days_ago))
        interval = MachineDowntime.objects.filter(record=record).get()
        start = timezone.now() - timedelta(days=days_ago)
        interval.start_at, interval.end_at = start, start + timedelta(hours=hours)
        interval.save()
        if category:
            record.fault_types.add(category)
        return record

    def test_downtime_hours_mttr_and_breakdown_counts(self):
        self.breakdown(hours=2)
        self.breakdown(machine=self.other_machine, hours=4)
        report = build_report(self.today - timedelta(days=30), self.today)
        self.assertEqual(report['breakdown_tickets'], 2)
        self.assertEqual(report['total_downtime_hours'], 6.0)
        self.assertAlmostEqual(report['overall_mttr_minutes'], 180.0, places=0)
        names = {row['machine']: row for row in report['by_machine']}
        self.assertEqual(names['SM74-A']['downtime_hours'], 2.0)

    def test_date_range_excludes_older_tickets(self):
        self.breakdown(days_ago=60)
        report = build_report(self.today - timedelta(days=30), self.today)
        self.assertEqual(report['total_tickets'], 0)

    def test_cancelled_tickets_are_excluded(self):
        self.make_record(status='CANCELLED')
        self.assertEqual(build_report(self.today - timedelta(days=1), self.today)['total_tickets'], 0)

    def test_ongoing_downtime_is_counted_and_flagged_but_not_in_mttr(self):
        record = self.make_record(status='IN_PROGRESS', maintenance_type='BREAKDOWN', stopped=True)
        MachineDowntime.objects.filter(record=record).update(start_at=timezone.now() - timedelta(hours=5))
        report = build_report(self.today - timedelta(days=1), self.today)
        row = report['by_machine'][0]
        self.assertTrue(row['ongoing'])
        self.assertGreaterEqual(row['downtime_hours'], 4.9)
        self.assertIsNone(report['overall_mttr_minutes'])

    def test_repeat_faults_need_two_tickets_with_the_same_category(self):
        category = FaultCategory.objects.create(name='Feeder-test')
        self.breakdown(category=category)
        self.breakdown(category=category, hours=1)
        self.breakdown(machine=self.other_machine, category=category)
        report = build_report(self.today - timedelta(days=30), self.today)
        self.assertEqual([(r['machine__name'], r['n']) for r in report['repeat_faults']], [('SM74-A', 2)])

    def test_repair_and_waiting_time_come_from_status_history(self):
        record = self.make_record(status='DIAGNOSED')
        base = timezone.now() - timedelta(hours=10)
        for status, hours in (('IN_PROGRESS', 0), ('AWAITING_PARTS', 3), ('IN_PROGRESS', 7)):
            log = MaintenanceActivityLog.objects.create(record=record, action='Status changed', to_status=status)
            MaintenanceActivityLog.objects.filter(pk=log.pk).update(created_at=base + timedelta(hours=hours))
        report = build_report(self.today - timedelta(days=1), self.today)
        self.assertEqual(report['waiting_hours'], 4.0)
        self.assertEqual(report['repair_hours'], 6.0)       # 0-3h, then 7h-now (3h)

    def test_pm_compliance(self):
        from maintenance.models import PreventiveMaintenancePlan
        from maintenance.services import generate_due_pm_records

        for index in range(2):
            PreventiveMaintenancePlan.objects.create(
                machine=self.machine, title=f'Plan {index}', interval_type='DAYS', interval_value=7, next_due_at=self.today,
            )
        records = generate_due_pm_records(actor=self.manager)['created']
        type(records[0]).objects.filter(pk=records[0].pk).update(status='CLOSED')
        report = build_report(self.today - timedelta(days=1), self.today)
        self.assertEqual((report['pm_total'], report['pm_done'], report['pm_compliance_pct']), (2, 1, 50.0))

    def test_report_page_renders_for_staff_and_filters_by_date(self):
        self.breakdown(days_ago=2)
        self.client.force_login(self.engineer)
        url = reverse('maintenance:reports')
        self.assertContains(self.client.get(url), 'SM74-A')
        old = (self.today - timedelta(days=400)).isoformat()
        response = self.client.get(url, {'date_from': old, 'date_to': (self.today - timedelta(days=300)).isoformat()})
        self.assertNotContains(response, 'SM74-A')


class DashboardAndListTests(MaintenanceTestCase):
    def test_operator_dashboard_shows_own_problems_in_plain_words_only(self):
        self.make_record(reporter=self.operator, status='AWAITING_PARTS')
        self.make_record(reporter=self.operator2, fault_description='Somebody else')
        self.client.force_login(self.operator)
        response = self.client.get(reverse('maintenance:dashboard'))
        self.assertContains(response, 'Waiting for spare parts')
        self.assertNotContains(response, 'Somebody else')
        self.assertNotContains(response, 'Open Tickets')

    def test_engineer_dashboard_shows_attention_items(self):
        critical = self.make_record(status='REPORTED', stopped=True, priority='CRITICAL')
        type(critical).objects.filter(pk=critical.pk).update(created_at=timezone.now() - timedelta(hours=6))
        self.client.force_login(self.engineer)
        response = self.client.get(reverse('maintenance:dashboard'))
        self.assertContains(response, 'Critical Faults')
        self.assertContains(response, 'Needs Attention')
        self.assertContains(response, 'Machines Currently Stopped')

    def test_list_filters_combine(self):
        mine = self.make_record(status='IN_PROGRESS', assigned_to=self.engineer, fault_description='Alpha')
        other = self.make_record(status='IN_PROGRESS', assigned_to=self.manager, fault_description='Beta')
        old = self.make_record(status='IN_PROGRESS', assigned_to=self.engineer, fault_description='Gamma')
        type(old).objects.filter(pk=old.pk).update(reported_date=timezone.localdate() - timedelta(days=40))
        self.client.force_login(self.manager)
        response = self.client.get(reverse('maintenance:record_list'), {
            'assigned_to': self.engineer.pk, 'status': 'IN_PROGRESS',
            'date_from': (timezone.localdate() - timedelta(days=7)).isoformat(),
        })
        records = list(response.context['records'])
        self.assertEqual(records, [mine])
        self.assertNotIn(other, records)


class DemandTests(MaintenanceTestCase):
    def test_raise_demand_is_safe_to_repeat(self):
        ItemRequestType.objects.get_or_create(code='MNT', defaults={'name': 'Maintenance', 'is_active': True})
        record = self.make_record(status='AWAITING_PARTS', spare_parts_needed=True)
        MaintenanceSparePart.objects.create(record=record, description='Feeder belt', quantity=2, uom='pcs')
        self.client.force_login(self.engineer)
        url = reverse('maintenance:raise_demand', args=[record.pk])
        self.client.post(url)
        self.client.post(url)
        self.assertEqual(ItemRequest.objects.filter(machine=self.machine).count(), 1)
        line = record.spare_parts.get()
        self.assertIsNotNone(line.item_request)


class ItemRequestAccessTests(MaintenanceTestCase):
    def test_every_role_including_custom_ones_can_access_item_requests(self):
        Role.objects.create(slug='custom-tester', display_name='Custom tester')
        call_command('seed_access_control', stdout=StringIO())
        permission = Permission.objects.get(code='nav.item_request')
        missing = [role.slug for role in Role.objects.all() if not role.permissions.filter(pk=permission.pk).exists()]
        self.assertEqual(missing, [])
        self.assertTrue(Role.objects.get(slug='maintenance_engineer').permissions.filter(pk=permission.pk).exists())
        self.assertTrue(Role.objects.get(slug='viewer').permissions.filter(pk=permission.pk).exists())


class SchedulerTests(MaintenanceTestCase):
    def test_run_once_generates_due_tickets_and_is_idempotent(self):
        from maintenance.models import PreventiveMaintenancePlan
        from maintenance.scheduler import run_once

        self.manager.is_superuser = True
        self.manager.save()
        PreventiveMaintenancePlan.objects.create(
            machine=self.machine, title='Sched', interval_type='DAYS', interval_value=7, next_due_at=timezone.localdate(),
        )
        self.assertEqual(len(run_once()['created']), 1)
        self.assertEqual(len(run_once()['created']), 0)
