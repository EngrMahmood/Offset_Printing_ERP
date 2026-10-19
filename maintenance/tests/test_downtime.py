from maintenance.models import MachineDowntime
from maintenance.services import open_breakdown_downtime_if_needed, open_downtime, soft_delete_record
from maintenance.workflow import TransitionError, transition_record

from .base import MaintenanceTestCase


def finish_repair(record, user, test='PASS'):
    record.root_cause = 'cause'
    record.repair_action = 'fix'
    record.test_result = test
    record.save()


class DowntimeCorrectnessTests(MaintenanceTestCase):
    def test_stopped_complaint_opens_exactly_one_interval(self):
        self.make_record(stopped=True)
        self.assertEqual(self.open_downtimes().count(), 1)

    def test_breakdown_triage_does_not_duplicate_downtime(self):
        record = self.make_record(stopped=True)
        record.maintenance_type = 'BREAKDOWN'
        record.save()
        open_breakdown_downtime_if_needed(record)
        open_breakdown_downtime_if_needed(record)
        self.assertEqual(MachineDowntime.objects.filter(machine=self.machine).count(), 1)

    def test_two_tickets_on_one_machine_share_one_interval(self):
        self.make_record(stopped=True)
        second = self.make_record(stopped=True, fault_description='Second fault')
        self.assertEqual(MachineDowntime.objects.filter(machine=self.machine).count(), 1)
        self.assertIsNotNone(second.pk)

    def test_breakdown_triage_marks_ticket_as_stopping_the_machine(self):
        record = self.make_record(stopped=False, maintenance_type='BREAKDOWN')
        open_breakdown_downtime_if_needed(record)
        record.refresh_from_db()
        self.assertTrue(record.machine_stopped)
        self.assertEqual(self.open_downtimes().count(), 1)

    def test_completion_does_not_restore_the_machine(self):
        record = self.make_record(status='IN_PROGRESS', stopped=True)
        finish_repair(record, self.engineer)
        transition_record(record, 'COMPLETED', self.engineer)
        self.assertEqual(self.open_downtimes().count(), 1)

    def test_verification_closes_downtime_and_records_who_and_why(self):
        record = self.make_record(status='COMPLETED', stopped=True)
        finish_repair(record, self.engineer)
        transition_record(record, 'VERIFIED', self.engineer)
        self.assertEqual(self.open_downtimes().count(), 0)
        interval = MachineDowntime.objects.get(machine=self.machine)
        self.assertEqual(interval.closed_by, self.engineer)
        self.assertIn('verified', interval.closed_reason.lower())

    def test_cancel_stopped_ticket_requires_machine_state(self):
        record = self.make_record(status='DIAGNOSED', stopped=True)
        with self.assertRaises(TransitionError):
            transition_record(record, 'CANCELLED', self.engineer, note='duplicate')
        self.assertEqual(self.open_downtimes().count(), 1)

    def test_cancel_does_not_restore_a_machine_that_is_still_down(self):
        record = self.make_record(status='DIAGNOSED', stopped=True)
        # nobody else covers the stoppage, so "still down" is refused
        with self.assertRaises(TransitionError):
            transition_record(record, 'CANCELLED', self.engineer, note='dup', machine_state='still_down')
        self.assertEqual(self.open_downtimes().count(), 1)

    def test_cancel_with_still_down_hands_downtime_to_other_ticket(self):
        first = self.make_record(status='DIAGNOSED', stopped=True)
        second = self.make_record(status='DIAGNOSED', stopped=True, fault_description='Same fault again')
        transition_record(first, 'CANCELLED', self.engineer, note='duplicate', machine_state='still_down')
        interval = self.open_downtimes().get()
        self.assertEqual(interval.record_id, second.pk)

    def test_cancel_with_machine_restored_closes_downtime(self):
        record = self.make_record(status='DIAGNOSED', stopped=True)
        transition_record(record, 'CANCELLED', self.engineer, note='false alarm', machine_state='restored')
        self.assertEqual(self.open_downtimes().count(), 0)

    def test_closing_one_ticket_does_not_restore_machine_with_another_stopped_ticket(self):
        first = self.make_record(status='COMPLETED', stopped=True)
        second = self.make_record(status='IN_PROGRESS', stopped=True, fault_description='Gripper fault')
        finish_repair(first, self.engineer)
        transition_record(first, 'VERIFIED', self.engineer)
        interval = self.open_downtimes().get()          # machine is still down
        self.assertEqual(interval.record_id, second.pk)
        finish_repair(second, self.engineer)
        transition_record(second, 'COMPLETED', self.engineer)
        transition_record(second, 'VERIFIED', self.engineer)
        self.assertEqual(self.open_downtimes().count(), 0)

    def test_running_machine_ticket_does_not_keep_downtime_open(self):
        stopped = self.make_record(status='COMPLETED', stopped=True)
        self.make_record(status='IN_PROGRESS', stopped=False, fault_description='Cosmetic')
        finish_repair(stopped, self.engineer)
        transition_record(stopped, 'VERIFIED', self.engineer)
        self.assertEqual(self.open_downtimes().count(), 0)

    def test_other_machines_are_unaffected(self):
        first = self.make_record(status='COMPLETED', stopped=True)
        self.make_record(machine=self.other_machine, status='IN_PROGRESS', stopped=True)
        finish_repair(first, self.engineer)
        transition_record(first, 'VERIFIED', self.engineer)
        self.assertEqual(self.open_downtimes(self.machine).count(), 0)
        self.assertEqual(self.open_downtimes(self.other_machine).count(), 1)

    def test_soft_delete_keeps_the_ticket_and_ends_its_downtime(self):
        record = self.make_record(stopped=True)
        soft_delete_record(record, self.manager, 'entered by mistake')
        record.refresh_from_db()
        self.assertFalse(record.is_active)
        self.assertEqual(self.open_downtimes().count(), 0)
        self.assertTrue(record.approvals.filter(action='DELETE').exists())
        self.assertTrue(record.activity_log.filter(action='Record deleted').exists())

    def test_open_downtime_reuses_existing_interval(self):
        record = self.make_record(stopped=True)
        self.assertIsNone(open_downtime(record))
