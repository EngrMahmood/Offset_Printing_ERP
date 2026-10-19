from decimal import Decimal

from maintenance.models import MaintenanceActivityLog, MaintenanceRecord
from maintenance.workflow import TRANSITIONS, TransitionError, allowed_transitions, transition_record

from .base import MaintenanceTestCase


def add_wait_lines(record):
    """A ticket can only wait on parts / a vendor when there is something to wait for."""
    from maintenance.models import MaintenanceServiceJob, MaintenanceSparePart

    MaintenanceSparePart.objects.create(record=record, description='Belt', quantity=1, uom='pcs')
    MaintenanceServiceJob.objects.create(record=record, vendor='ACME', scope='Rebuild gearbox')


def ready_for_completion(record):
    record.root_cause = 'Worn feeder belt'
    record.repair_action = 'Replaced belt'
    record.save()
    return record


class TransitionRuleTests(MaintenanceTestCase):
    def test_every_valid_transition_works(self):
        for from_status, targets in TRANSITIONS.items():
            for target in targets - {'DIAGNOSED'}:
                with self.subTest(f'{from_status}->{target}'):
                    # a ticket that did not stop the machine keeps this test about status rules only
                    record = self.make_record(status=from_status)
                    add_wait_lines(record)
                    ready_for_completion(record)
                    record.test_result = 'PASS'
                    record.save()
                    note = 'because'
                    actor = self.manager if from_status == 'CLOSED' else self.engineer
                    updated = transition_record(record, target, actor, note=note, machine_state='restored')
                    self.assertEqual(updated.status, target)

    def test_every_invalid_transition_is_rejected_without_changes(self):
        for from_status in TRANSITIONS:
            valid = TRANSITIONS[from_status]
            for target, _label in MaintenanceRecord.STATUS_CHOICES:
                if target in valid or target == from_status and False:
                    continue
                with self.subTest(f'{from_status}->{target}'):
                    record = self.make_record(status=from_status)
                    before_logs = MaintenanceActivityLog.objects.filter(record=record).count()
                    with self.assertRaises(TransitionError):
                        transition_record(record, target, self.manager, note='x')
                    record.refresh_from_db()
                    self.assertEqual(record.status, from_status)
                    self.assertEqual(MaintenanceActivityLog.objects.filter(record=record).count(), before_logs)

    def test_diagnosed_only_through_triage(self):
        record = self.make_record(status='REPORTED')
        with self.assertRaises(TransitionError):
            transition_record(record, 'DIAGNOSED', self.engineer)

    def test_completion_needs_diagnosis_and_repair_action(self):
        record = self.make_record(status='IN_PROGRESS')
        with self.assertRaises(TransitionError):
            transition_record(record, 'COMPLETED', self.engineer)
        record.root_cause = 'x'
        record.save()
        with self.assertRaises(TransitionError):
            transition_record(record, 'COMPLETED', self.engineer)
        record.repair_action = 'y'
        record.save()
        self.assertEqual(transition_record(record, 'COMPLETED', self.engineer).status, 'COMPLETED')

    def test_negative_labour_rejected(self):
        record = ready_for_completion(self.make_record(status='IN_PROGRESS'))
        with self.assertRaises(TransitionError):
            transition_record(record, 'COMPLETED', self.engineer, labour_hours='-2')

    def test_verification_needs_a_passed_test(self):
        record = ready_for_completion(self.make_record(status='COMPLETED'))
        with self.assertRaises(TransitionError):
            transition_record(record, 'VERIFIED', self.engineer)
        record.test_result = 'FAIL'
        record.save()
        with self.assertRaises(TransitionError):
            transition_record(record, 'VERIFIED', self.engineer)
        record.test_result = 'PASS'
        record.save()
        verified = transition_record(record, 'VERIFIED', self.engineer)
        self.assertEqual(verified.verified_by, self.engineer)
        self.assertIsNotNone(verified.verified_at)

    def test_cancel_requires_reason(self):
        record = self.make_record(status='REPORTED')
        with self.assertRaises(TransitionError):
            transition_record(record, 'CANCELLED', self.engineer, note='')

    def test_reopen_restricted_and_needs_reason(self):
        record = self.make_record(status='CLOSED')
        with self.assertRaises(TransitionError):
            transition_record(record, 'IN_PROGRESS', self.engineer, note='again')   # engineer not allowed
        with self.assertRaises(TransitionError):
            transition_record(record, 'IN_PROGRESS', self.manager, note='')          # reason missing
        self.assertNotIn('IN_PROGRESS', allowed_transitions(record, self.engineer))
        self.assertIn('IN_PROGRESS', allowed_transitions(record, self.manager))
        reopened = transition_record(record, 'IN_PROGRESS', self.manager, note='Leak returned')
        self.assertEqual(reopened.status, 'IN_PROGRESS')

    def test_failed_test_sends_ticket_back_with_reason(self):
        record = ready_for_completion(self.make_record(status='COMPLETED'))
        with self.assertRaises(TransitionError):
            transition_record(record, 'IN_PROGRESS', self.engineer, note='')
        self.assertEqual(transition_record(record, 'IN_PROGRESS', self.engineer, note='Test run failed').status, 'IN_PROGRESS')

    def test_transition_writes_audit_entry(self):
        record = self.make_record(status='DIAGNOSED')
        transition_record(record, 'IN_PROGRESS', self.engineer)
        log = MaintenanceActivityLog.objects.filter(record=record, action='Status changed').first()
        self.assertEqual((log.from_status, log.to_status, log.actor), ('DIAGNOSED', 'IN_PROGRESS', self.engineer))

    def test_labour_hours_saved_as_decimal(self):
        record = ready_for_completion(self.make_record(status='IN_PROGRESS'))
        updated = transition_record(record, 'COMPLETED', self.engineer, labour_hours='2.5')
        self.assertEqual(Decimal(str(updated.labour_hours)), Decimal('2.5'))
