from django.core.files.uploadedfile import SimpleUploadedFile
from django.urls import reverse

from maintenance.models import MaintenanceAttachment, MaintenanceRecord

from .base import MaintenanceTestCase, User


class ComplaintReportingTests(MaintenanceTestCase):
    def post_complaint(self, **overrides):
        data = {'machine': self.machine.pk, 'fault_description': 'Delivery jam', 'machine_stopped': 'no'}
        data.update(overrides)
        return self.client.post(reverse('maintenance:complaint_create'), data, follow=True)

    def test_operator_can_report_a_running_fault(self):
        self.login(self.operator)
        self.post_complaint()
        record = MaintenanceRecord.objects.get()
        self.assertEqual((record.status, record.priority, record.machine_stopped), ('REPORTED', 'MEDIUM', False))
        self.assertTrue(record.record_no.startswith('MNT-'))
        self.assertEqual(self.open_downtimes().count(), 0)

    def test_stopped_machine_is_critical_and_opens_one_downtime(self):
        self.login(self.operator)
        self.post_complaint(machine_stopped='yes')
        record = MaintenanceRecord.objects.get()
        self.assertEqual((record.priority, record.machine_stopped), ('CRITICAL', True))
        self.assertEqual(self.open_downtimes().count(), 1)

    def test_double_submit_creates_one_ticket(self):
        self.login(self.operator)
        self.post_complaint(machine_stopped='yes')
        self.post_complaint(machine_stopped='yes')
        self.assertEqual(MaintenanceRecord.objects.count(), 1)
        self.assertEqual(self.open_downtimes().count(), 1)

    def test_existing_open_ticket_warns_before_a_second_one(self):
        self.make_record()
        self.login(self.operator2)
        response = self.post_complaint(fault_description='Different problem')
        self.assertEqual(MaintenanceRecord.objects.count(), 1)
        self.assertContains(response, self.machine.name)
        self.post_complaint(fault_description='Different problem', confirm_duplicate='1')
        self.assertEqual(MaintenanceRecord.objects.count(), 2)

    def test_photo_is_stored_and_disallowed_types_are_refused(self):
        self.login(self.operator)
        good = SimpleUploadedFile('fault.jpg', b'\xff\xd8\xff', content_type='image/jpeg')
        self.post_complaint(photo=good)
        self.assertEqual(MaintenanceAttachment.objects.count(), 1)
        bad = SimpleUploadedFile('run.exe', b'MZ', content_type='application/octet-stream')
        self.post_complaint(fault_description='Another one', photo=bad)
        self.assertEqual(MaintenanceRecord.objects.count(), 1)

    def test_attachment_size_limit(self):
        record = self.make_record()
        self.login(self.operator)
        big = SimpleUploadedFile('big.pdf', b'0' * (10 * 1024 * 1024 + 1), content_type='application/pdf')
        self.client.post(reverse('maintenance:attachment_add', args=[record.pk]), {'file': big})
        self.assertEqual(MaintenanceAttachment.objects.count(), 0)


class AccessControlTests(MaintenanceTestCase):
    def test_anonymous_is_redirected_to_login(self):
        record = self.make_record()
        for url in (
            reverse('maintenance:dashboard'), reverse('maintenance:record_list'),
            reverse('maintenance:record_detail', args=[record.pk]), reverse('maintenance:reports'),
        ):
            with self.subTest(url):
                response = self.client.get(url)
                self.assertEqual(response.status_code, 302)
                self.assertIn('login', response['Location'])

    def test_user_without_module_permission_is_denied(self):
        record = self.make_record()
        self.login(self.outsider)
        for url in (
            reverse('maintenance:dashboard'), reverse('maintenance:complaint_create'),
            reverse('maintenance:record_detail', args=[record.pk]),
        ):
            with self.subTest(url):
                self.assertEqual(self.client.get(url).status_code, 403)

    def test_operator_cannot_triage_reassign_transition_or_edit_repair(self):
        record = self.make_record(reporter=self.operator, status='DIAGNOSED')
        self.login(self.operator)
        targets = [
            ('maintenance:record_triage', {}), ('maintenance:record_reassign', {'assigned_to': self.engineer.pk}),
            ('maintenance:record_transition', {'to_status': 'IN_PROGRESS'}),
            ('maintenance:record_repair_update', {'root_cause': 'x', 'repair_action': 'y', 'test_result': 'PASS'}),
            ('maintenance:spare_part_add', {'description': 'belt', 'quantity': '1'}),
            ('maintenance:raise_demand', {}),
        ]
        for name, data in targets:
            with self.subTest(name):
                self.assertEqual(self.client.post(reverse(name, args=[record.pk]), data).status_code, 403)
        record.refresh_from_db()
        self.assertEqual(record.status, 'DIAGNOSED')
        self.assertEqual(record.root_cause, '')

    def test_operator_cannot_open_downtime_pm_or_report_screens(self):
        self.login(self.operator)
        for name in ('maintenance:downtime_list', 'maintenance:pm_plan_list', 'maintenance:reports'):
            with self.subTest(name):
                self.assertEqual(self.client.get(reverse(name)).status_code, 403)

    def test_operator_sees_only_their_own_tickets(self):
        mine = self.make_record(reporter=self.operator)
        theirs = self.make_record(reporter=self.operator2, fault_description='Not mine')
        self.login(self.operator)
        listing = self.client.get(reverse('maintenance:record_list'))
        self.assertContains(listing, mine.record_no)
        self.assertNotContains(listing, theirs.record_no)
        self.assertEqual(self.client.get(reverse('maintenance:record_detail', args=[mine.pk])).status_code, 200)
        self.assertEqual(self.client.get(reverse('maintenance:record_detail', args=[theirs.pk])).status_code, 403)

    def test_operator_cannot_comment_or_attach_on_someone_elses_ticket(self):
        theirs = self.make_record(reporter=self.operator2)
        self.login(self.operator)
        self.assertEqual(self.client.post(reverse('maintenance:record_comment', args=[theirs.pk]), {'note': 'hi'}).status_code, 403)
        upload = SimpleUploadedFile('a.pdf', b'x', content_type='application/pdf')
        self.assertEqual(self.client.post(reverse('maintenance:attachment_add', args=[theirs.pk]), {'file': upload}).status_code, 403)

    def test_operator_cannot_export(self):
        self.login(self.operator)
        self.assertEqual(self.client.get(reverse('maintenance:record_list') + '?export=xlsx').status_code, 403)

    def test_engineer_can_triage_and_assign(self):
        record = self.make_record(status='REPORTED')
        self.login(self.engineer)
        from maintenance.models import FaultCategory

        category = FaultCategory.objects.create(name='Electrical-test')
        response = self.client.post(reverse('maintenance:record_triage', args=[record.pk]), {
            'fault_types': [category.pk], 'maintenance_type': 'CORRECTIVE', 'priority': 'MAJOR',
            'execution_type': 'IN_HOUSE', 'assigned_to': self.engineer.pk,
        })
        self.assertEqual(response.status_code, 302)
        record.refresh_from_db()
        self.assertEqual((record.status, record.assigned_to), ('DIAGNOSED', self.engineer))

    def test_engineer_moves_ticket_through_the_view_and_invalid_move_changes_nothing(self):
        record = self.make_record(status='DIAGNOSED')
        self.login(self.engineer)
        url = reverse('maintenance:record_transition', args=[record.pk])
        self.client.post(url, {'to_status': 'CLOSED'})
        record.refresh_from_db()
        self.assertEqual(record.status, 'DIAGNOSED')
        self.client.post(url, {'to_status': 'IN_PROGRESS'})
        record.refresh_from_db()
        self.assertEqual(record.status, 'IN_PROGRESS')
        self.assertIsNotNone(record.work_start_at)

    def test_get_never_changes_state(self):
        record = self.make_record(status='DIAGNOSED')
        self.login(self.engineer)
        self.client.get(reverse('maintenance:record_transition', args=[record.pk]) + '?to_status=IN_PROGRESS')
        record.refresh_from_db()
        self.assertEqual(record.status, 'DIAGNOSED')

    def test_repair_update_validates_input(self):
        record = self.make_record(status='IN_PROGRESS')
        self.login(self.engineer)
        url = reverse('maintenance:record_repair_update', args=[record.pk])
        self.client.post(url, {'root_cause': 'a', 'repair_action': 'b', 'test_result': 'PASS', 'labour_hours': '-1'})
        record.refresh_from_db()
        self.assertEqual(record.root_cause, '')
        self.client.post(url, {'root_cause': 'a', 'repair_action': 'b', 'test_result': 'PASS', 'labour_hours': '1.5'})
        record.refresh_from_db()
        self.assertEqual((record.root_cause, str(record.labour_hours)), ('a', '1.50'))

    def test_only_superuser_can_bulk_delete(self):
        record = self.make_record()
        self.login(self.manager)
        self.assertEqual(self.client.post(reverse('maintenance:bulk_delete'), {'selected_ids': [record.pk]}).status_code, 403)

    def test_soft_deleted_ticket_is_hidden(self):
        record = self.make_record()
        record.is_active = False
        record.save()
        self.login(self.manager)
        self.assertEqual(self.client.get(reverse('maintenance:record_detail', args=[record.pk])).status_code, 404)


class SuperuserStaffConsistencyTests(MaintenanceTestCase):
    def test_superuser_with_default_profile_role_can_work_a_ticket(self):
        # The staff test used by templates and by the decorator must agree.
        admin = User.objects.create_superuser('root1', password='pw-for-tests-12345')
        record = self.make_record(status='DIAGNOSED')
        self.client.force_login(admin)
        self.client.post(reverse('maintenance:record_transition', args=[record.pk]), {'to_status': 'IN_PROGRESS'})
        record.refresh_from_db()
        self.assertEqual(record.status, 'IN_PROGRESS')
