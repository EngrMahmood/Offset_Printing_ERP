from datetime import date, datetime, timedelta

from django.urls import reverse
from django.utils import timezone

from maintenance.downtime_analysis import (
    TEAM_MAINTENANCE, TEAM_SUPPLIER, TEAM_SUPPLY, TEAM_VENDOR, build_downtime_analysis, fmt_duration,
    machine_overview, parts_stage_segments,
)
from maintenance.models import DowntimePhase, MachineDowntime, MaintenanceActivityLog, MaintenanceSparePart
from maintenance.signals import notify_parts_received_for
from maintenance.workflow import TransitionError, transition_record
from supply_chain.models import ItemProcurementTimeline, ItemRequest, ItemRequestType

from .base import MaintenanceTestCase


def aware(year, month, day, hour=0):
    return timezone.make_aware(datetime(year, month, day, hour))


def current_phase(machine):
    interval = MachineDowntime.objects.filter(machine=machine, end_at__isnull=True).first()
    return interval.phases.filter(ended_at__isnull=True).first() if interval else None


class PhaseTrackingTests(MaintenanceTestCase):
    def test_stopped_complaint_starts_in_waiting_for_maintenance(self):
        self.make_record(stopped=True)
        self.assertEqual(current_phase(self.machine).category, DowntimePhase.WAIT_RESPONSE)

    def test_ticket_journey_is_recorded_step_by_step(self):
        record = self.make_record(status='DIAGNOSED', stopped=True)
        record = transition_record(record, 'IN_PROGRESS', self.engineer)
        self.assertEqual(current_phase(self.machine).category, DowntimePhase.REPAIR)

        MaintenanceSparePart.objects.create(record=record, description='Feeder belt', quantity=2, uom='pcs')
        record = transition_record(record, 'AWAITING_PARTS', self.engineer)
        self.assertEqual(current_phase(self.machine).category, DowntimePhase.WAIT_PARTS)

        record = transition_record(record, 'IN_PROGRESS', self.engineer)
        self.assertEqual(current_phase(self.machine).category, DowntimePhase.REPAIR)

        record.root_cause, record.repair_action, record.test_result = 'worn', 'replaced', 'PASS'
        record.save()
        record = transition_record(record, 'COMPLETED', self.engineer)
        self.assertEqual(current_phase(self.machine).category, DowntimePhase.TESTING)

        transition_record(record, 'VERIFIED', self.engineer)
        self.assertIsNone(current_phase(self.machine))
        interval = MachineDowntime.objects.get(machine=self.machine)
        self.assertIsNotNone(interval.end_at)
        phases = list(interval.phases.all())
        self.assertTrue(all(p.ended_at for p in phases))
        self.assertEqual(
            [p.category for p in phases],
            ['WAIT_START', 'REPAIR', 'WAIT_PARTS', 'REPAIR', 'TESTING'],
        )

    def test_phases_are_contiguous_and_end_with_the_interval(self):
        record = self.make_record(status='DIAGNOSED', stopped=True)
        record = transition_record(record, 'IN_PROGRESS', self.engineer)
        record.root_cause, record.repair_action, record.test_result = 'a', 'b', 'PASS'
        record.save()
        record = transition_record(record, 'COMPLETED', self.engineer)
        transition_record(record, 'VERIFIED', self.engineer)
        interval = MachineDowntime.objects.get(machine=self.machine)
        phases = list(interval.phases.all())
        for earlier, later in zip(phases, phases[1:]):
            self.assertEqual(earlier.ended_at, later.started_at)
        self.assertEqual(phases[-1].ended_at, interval.end_at)

    def test_repeating_a_sync_does_not_create_extra_phases(self):
        from maintenance.services import sync_downtime_phase

        record = self.make_record(stopped=True)
        for _ in range(3):
            sync_downtime_phase(record)
        self.assertEqual(DowntimePhase.objects.count(), 1)

    def test_waiting_for_parts_needs_a_part_and_raises_the_request(self):
        record = self.make_record(status='IN_PROGRESS', stopped=True)
        with self.assertRaises(TransitionError):
            transition_record(record, 'AWAITING_PARTS', self.engineer)
        MaintenanceSparePart.objects.create(record=record, description='Gear', quantity=1, uom='pcs')
        transition_record(record, 'AWAITING_PARTS', self.engineer)
        self.assertEqual(ItemRequest.objects.filter(machine=self.machine).count(), 1)
        self.assertTrue(record.spare_parts.get().item_request_id)

    def test_waiting_for_vendor_needs_a_service_job(self):
        from maintenance.models import MaintenanceServiceJob

        record = self.make_record(status='IN_PROGRESS', stopped=True)
        with self.assertRaises(TransitionError):
            transition_record(record, 'AWAITING_VENDOR', self.engineer)
        MaintenanceServiceJob.objects.create(record=record, vendor='ACME', scope='Rewind motor')
        transition_record(record, 'AWAITING_VENDOR', self.engineer)
        self.assertEqual(current_phase(self.machine).category, DowntimePhase.WAIT_VENDOR)

    def test_downtime_handed_to_another_ticket_keeps_counting_in_that_tickets_step(self):
        first = self.make_record(status='COMPLETED', stopped=True)
        second = self.make_record(status='IN_PROGRESS', stopped=True, fault_description='Gripper')
        first.root_cause, first.repair_action, first.test_result = 'a', 'b', 'PASS'
        first.save()
        transition_record(first, 'VERIFIED', self.engineer)
        phase = current_phase(self.machine)
        self.assertEqual((phase.category, phase.record_id), (DowntimePhase.REPAIR, second.pk))

    def test_reopening_a_stopped_ticket_restarts_downtime(self):
        record = self.make_record(status='CLOSED', stopped=True)
        MachineDowntime.objects.all().delete()
        transition_record(record, 'IN_PROGRESS', self.manager, note='Fault is back')
        self.assertEqual(self.open_downtimes().count(), 1)
        self.assertEqual(current_phase(self.machine).category, DowntimePhase.REPAIR)


class StageSplitTests(MaintenanceTestCase):
    """Waiting for parts is split by what Supply Chain was doing."""

    def build_wait(self, request_date, pr=None, po=None, received=None, phase_start=None, phase_end=None):
        record = self.make_record(status='AWAITING_PARTS', stopped=True)
        from maintenance.services import _maintenance_department

        request_type = ItemRequestType.objects.get(code='MNT')
        request = ItemRequest.objects.create(
            department=_maintenance_department(), raised_by=self.engineer, request_type=request_type, request_date=request_date, item_title='Feeder belt', machine=self.machine,
            request_no='IR-TEST-1',
        )
        ItemProcurementTimeline.objects.create(request=request, pr_date=pr, po_date=po, received_date=received, supplier='ACME Spares')
        MaintenanceSparePart.objects.create(record=record, description='Feeder belt', quantity=1, uom='pcs', item_request=request)
        interval = MachineDowntime.objects.get(record=record)
        DowntimePhase.objects.filter(downtime=interval).delete()
        phase = DowntimePhase.objects.create(
            downtime=interval, record=record, category=DowntimePhase.WAIT_PARTS,
            started_at=phase_start, ended_at=phase_end,
        )
        return record, phase

    def test_parts_wait_is_split_by_procurement_stage(self):
        _record, phase = self.build_wait(
            date(2026, 3, 2), pr=date(2026, 3, 3), po=date(2026, 3, 4), received=date(2026, 3, 8),
            phase_start=aware(2026, 3, 2), phase_end=aware(2026, 3, 10),
        )
        days = {key: (end - start).days for key, start, end in parts_stage_segments(phase)}
        self.assertEqual(days, {'PARTS_APPROVAL': 1, 'PARTS_PURCHASING': 1, 'PARTS_DELIVERY': 4, 'PARTS_NOT_FITTED': 2})

    def test_time_before_the_request_is_attributed_to_maintenance(self):
        _record, phase = self.build_wait(
            date(2026, 3, 4), phase_start=aware(2026, 3, 2), phase_end=aware(2026, 3, 6),
        )
        pieces = {key: (end - start).days for key, start, end in parts_stage_segments(phase)}
        self.assertEqual(pieces, {'PARTS_BEFORE_REQUEST': 2, 'PARTS_APPROVAL': 2})

    def test_missing_dates_extend_the_previous_stage(self):
        _record, phase = self.build_wait(
            date(2026, 3, 2), po=date(2026, 3, 5), phase_start=aware(2026, 3, 2), phase_end=aware(2026, 3, 9),
        )
        pieces = {key: (end - start).days for key, start, end in parts_stage_segments(phase)}
        self.assertEqual(pieces, {'PARTS_APPROVAL': 3, 'PARTS_DELIVERY': 4})

    def test_wait_with_no_linked_request_is_flagged(self):
        record = self.make_record(status='AWAITING_PARTS', stopped=True)
        interval = MachineDowntime.objects.get(record=record)
        phase = DowntimePhase.objects.create(
            downtime=interval, record=record, category=DowntimePhase.WAIT_PARTS,
            started_at=aware(2026, 3, 2), ended_at=aware(2026, 3, 4),
        )
        self.assertEqual([p[0] for p in parts_stage_segments(phase)], ['PARTS_UNLINKED'])

    def test_analysis_attributes_time_to_the_owning_team(self):
        _record, _phase = self.build_wait(
            date(2026, 3, 2), pr=date(2026, 3, 3), po=date(2026, 3, 4), received=date(2026, 3, 8),
            phase_start=aware(2026, 3, 2), phase_end=aware(2026, 3, 10),
        )
        interval = MachineDowntime.objects.get(machine=self.machine)
        DowntimePhase.objects.create(
            downtime=interval, record=_record, category=DowntimePhase.REPAIR,
            started_at=aware(2026, 3, 10), ended_at=aware(2026, 3, 10, 6),
        )
        vendor_record = self.make_record(machine=self.other_machine, status='AWAITING_VENDOR', stopped=True)
        vendor_interval = MachineDowntime.objects.get(record=vendor_record)
        DowntimePhase.objects.filter(downtime=vendor_interval).delete()
        DowntimePhase.objects.create(
            downtime=vendor_interval, record=vendor_record, category=DowntimePhase.WAIT_VENDOR,
            started_at=aware(2026, 3, 3), ended_at=aware(2026, 3, 5),
        )
        MachineDowntime.objects.all().update(end_at=aware(2026, 3, 11))
        report = build_downtime_analysis(date(2026, 3, 1), date(2026, 3, 31))
        seconds = {team['key']: team['seconds'] for team in report['teams']}
        day = 86400
        self.assertEqual(seconds[TEAM_SUPPLY], 2 * day)             # approval 1d + purchasing 1d
        self.assertEqual(seconds[TEAM_SUPPLIER], 4 * day)           # PO -> receipt
        self.assertEqual(seconds[TEAM_VENDOR], 2 * day)
        self.assertEqual(seconds[TEAM_MAINTENANCE], 2 * day + 6 * 3600)   # received-not-fitted 2d + repair 6h
        steps = {step['key']: step['duration'] for step in report['steps']}
        self.assertEqual(steps['PARTS_DELIVERY'], '4d 0h')
        self.assertEqual(steps['REPAIR'], '6h 0m')

    def test_date_window_clips_phases(self):
        _record, _phase = self.build_wait(
            date(2026, 3, 2), pr=date(2026, 3, 3), po=date(2026, 3, 4), received=date(2026, 3, 8),
            phase_start=aware(2026, 3, 2), phase_end=aware(2026, 3, 10),
        )
        MachineDowntime.objects.all().update(end_at=aware(2026, 3, 11))
        report = build_downtime_analysis(date(2026, 3, 5), date(2026, 3, 6))
        self.assertEqual(report['total_seconds'], 2 * 86400)       # only 5th-6th, all delivery

    def test_ongoing_wait_is_listed_with_its_request_and_supplier(self):
        _record, _phase = self.build_wait(
            timezone.localdate() - timedelta(days=4), po=timezone.localdate() - timedelta(days=2),
            phase_start=timezone.now() - timedelta(days=4), phase_end=None,
        )
        report = build_downtime_analysis(timezone.localdate() - timedelta(days=10), timezone.localdate())
        wait = report['ongoing_waits'][0]
        self.assertEqual((wait['request_no'], wait['supplier'], wait['kind']), ('IR-TEST-1', 'ACME Spares', 'Spare parts'))
        self.assertEqual(wait['stage'], 'supplier delivery (PO to receipt)')

    def test_fmt_duration(self):
        self.assertEqual(fmt_duration(2 * 86400 + 4 * 3600), '2d 4h')
        self.assertEqual(fmt_duration(3 * 3600 + 15 * 60), '3h 15m')
        self.assertEqual(fmt_duration(12 * 60), '12m')


class MachineBoardAndDowntimeScreenTests(MaintenanceTestCase):
    def test_machine_board_lists_tickets_per_machine_with_current_state(self):
        stopped = self.make_record(machine=self.machine, stopped=True)
        self.make_record(machine=self.machine, status='CLOSED', fault_description='Old problem')
        self.make_record(machine=self.other_machine, status='IN_PROGRESS', fault_description='Lamp')
        self.client.force_login(self.engineer)
        response = self.client.get(reverse('maintenance:machine_board'))
        self.assertEqual(response.status_code, 200)
        rows = {row['machine'].name: row for row in response.context['rows']}
        self.assertTrue(rows['SM74-A']['stopped'])
        self.assertEqual(len(rows['SM74-A']['tickets']), 2)
        self.assertEqual(len(rows['SM74-A']['open_tickets']), 1)
        self.assertFalse(rows['GTO-B']['stopped'])
        self.assertEqual(response.context['rows'][0]['machine'].name, 'SM74-A')       # stopped machines first
        self.assertContains(response, stopped.record_no)

    def test_operator_sees_machine_status_but_only_their_own_tickets(self):
        self.make_record(machine=self.machine, reporter=self.operator, stopped=True)
        theirs = self.make_record(machine=self.machine, reporter=self.operator2, fault_description='Theirs')
        self.client.force_login(self.operator)
        response = self.client.get(reverse('maintenance:machine_board'))
        row = response.context['rows'][0]
        self.assertTrue(row['stopped'])
        self.assertEqual(len(row['tickets']), 1)
        self.assertNotContains(response, theirs.record_no)

    def test_downtime_window_is_staff_only_and_renders_for_engineers(self):
        self.make_record(stopped=True)
        self.client.force_login(self.operator)
        self.assertEqual(self.client.get(reverse('maintenance:downtime_list')).status_code, 403)
        self.client.force_login(self.engineer)
        response = self.client.get(reverse('maintenance:downtime_list'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'By process step')
        self.assertContains(response, 'Each stoppage')

    def test_downtime_window_filters_by_machine(self):
        self.make_record(stopped=True)
        self.make_record(machine=self.other_machine, stopped=True)
        self.client.force_login(self.engineer)
        response = self.client.get(reverse('maintenance:downtime_list'), {'machine': self.other_machine.pk})
        self.assertEqual({row['interval'].machine for row in response.context['intervals']}, {self.other_machine})

    def test_ticket_detail_shows_where_the_downtime_went(self):
        record = self.make_record(stopped=True)
        self.client.force_login(self.engineer)
        self.assertContains(self.client.get(reverse('maintenance:record_detail', args=[record.pk])), 'where the time went')

    def test_subnav_hides_staff_screens_from_operators(self):
        self.client.force_login(self.operator)
        body = self.client.get(reverse('maintenance:dashboard')).content.decode()
        self.assertIn(reverse('maintenance:machine_board'), body)
        self.assertNotIn(reverse('maintenance:downtime_list'), body)
        self.assertNotIn(reverse('maintenance:reports'), body)


class PartsReceivedNotificationTests(MaintenanceTestCase):
    def setUp(self):
        super().setUp()
        self.record = self.make_record(status='IN_PROGRESS', stopped=True)
        MaintenanceSparePart.objects.create(record=self.record, description='Belt', quantity=1, uom='pcs')
        self.record = transition_record(self.record, 'AWAITING_PARTS', self.engineer)
        self.request = self.record.spare_parts.get().item_request

    def received_logs(self):
        return MaintenanceActivityLog.objects.filter(record=self.record, action='Parts received').count()

    def test_receipt_is_noticed_once(self):
        procurement, _ = ItemProcurementTimeline.objects.get_or_create(request=self.request)
        procurement.received_date = timezone.localdate()
        procurement.save()
        procurement.save()
        self.assertEqual(self.received_logs(), 1)

    def test_nothing_happens_before_receipt(self):
        ItemProcurementTimeline.objects.get_or_create(request=self.request)
        notify_parts_received_for(self.request)
        self.assertEqual(self.received_logs(), 0)

    def test_a_second_wait_can_be_notified_again(self):
        procurement, _ = ItemProcurementTimeline.objects.get_or_create(request=self.request)
        procurement.received_date = timezone.localdate()
        procurement.save()
        record = transition_record(self.record, 'IN_PROGRESS', self.engineer)
        MaintenanceSparePart.objects.create(record=record, description='Gear', quantity=1, uom='pcs')
        transition_record(record, 'AWAITING_PARTS', self.engineer)
        new_request = record.spare_parts.get(description='Gear').item_request
        # first line still received, second not yet
        notify_parts_received_for(new_request)
        self.assertEqual(self.received_logs(), 1)
        second, _ = ItemProcurementTimeline.objects.get_or_create(request=new_request)
        second.received_date = timezone.localdate()
        second.save()
        self.assertEqual(self.received_logs(), 2)
