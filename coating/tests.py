import datetime
from unittest import mock

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from core.models import Machine, Operator, Production, UserProfile
from coating import services
from coating.models import CoatingRun
from planning.tests_merge import make_card, make_job

DAY = datetime.date(2026, 10, 20)


class CoatingTests(TestCase):
    def setUp(self):
        patcher = mock.patch.object(UserProfile, '_has_permission', return_value=True)    # role matrix is not seeded in tests
        patcher.start()
        self.addCleanup(patcher.stop)
        self.user = get_user_model().objects.create_superuser('boss', 'b@example.com', 'pw12345678')
        self.client.force_login(self.user)
        self.machine, _ = Machine.objects.get_or_create(name='GTO 1A')
        self.operator = Operator.objects.create(name='Coater Chaudhry')
        self.job = make_job('CT1', 9000, sku='SKU-CT')
        self.job.application = 'UV'
        self.job.save(update_fields=['application'])
        self.card = make_card(self.job)
        self.printed(10000)

    def printed(self, sheets):
        Production.objects.bulk_create([Production(
            job_card=self.card, entry_type='printing', date=DAY, shift='A', machine=self.machine,
            output_sheets=sheets, impressions=sheets, run_time=120,
        )])

    def post(self, **extra):
        data = {'job_card': self.card.pk, 'date': DAY.isoformat(), 'shift': 'A', 'machine': self.machine.pk,
                'operator': self.operator.pk, 'sheets_coated': 4000, 'waste_sheets': 40, 'coats': 1,
                'start_time': '10:00', 'end_time': '11:30', 'status': 'in_progress'}
        data.update(extra)
        return self.client.post(reverse('coating_entry'), data)

    def test_kinds(self):
        self.assertEqual(services.coating_kind('UV'), 'uv')
        self.assertEqual(services.coating_kind('MATT VARNISH'), 'matt_varnish')
        self.assertEqual(services.coating_kind('Lamination Matt'), '')
        self.assertEqual(services.coating_kind('NO'), '')

    def test_entry_is_saved_with_run_time_from_clock_times(self):
        self.assertEqual(self.post().status_code, 302)
        run = CoatingRun.objects.get()
        self.assertEqual((run.sheets_coated, run.waste_sheets, run.run_minutes), (4000, 40, 90))
        self.assertEqual(run.impressions, 4040)
        self.assertEqual(run.coating_type, 'uv')

    def test_printing_numbers_are_untouched(self):
        before = (self.card.total_production, Production.objects.count())
        self.post()
        self.card.refresh_from_db()
        self.assertEqual((self.card.total_production, Production.objects.count()), before)

    def test_cannot_coat_more_than_was_printed(self):
        self.post(sheets_coated=9000)
        r = self.post(sheets_coated=2000)                       # 11,000 against 10,000 printed
        self.assertEqual(r.status_code, 200)
        self.assertContains(r, 'only 10,000 good sheets were printed')
        self.assertEqual(CoatingRun.objects.count(), 1)

    def test_no_coating_before_printing(self):
        other = make_card(make_job('CT2', 5000, sku='SKU-CT2'))
        r = self.post(job_card=other.pk)
        self.assertContains(r, 'no printing entry yet')

    def test_second_day_on_another_machine_adds_up(self):
        other_machine, _ = Machine.objects.get_or_create(name='SM 74')
        self.post(sheets_coated=4000)
        self.post(sheets_coated=3000, machine=other_machine.pk, date=(DAY + datetime.timedelta(days=3)).isoformat(), shift='B')
        info = services.summary(self.card)
        self.assertEqual((info['coated'], info['remaining'], info['runs']), (7000, 3000, 2))

    def test_pending_queue_and_done_flag(self):
        self.assertEqual([r['job_card'].pk for r in services.pending_rows()], [self.card.pk])
        self.post(sheets_coated=10000, status='completed')
        self.assertEqual(services.pending_rows(), [])

    def test_edit_needs_reason_and_remove_hides_the_record(self):
        self.post()
        run = CoatingRun.objects.get()
        edit = {'edit_id': run.pk, 'sheets_coated': 5000}
        self.assertContains(self.post(**edit), 'Change reason is required')
        self.assertEqual(self.post(change_reason='counted again', **edit).status_code, 302)
        run.refresh_from_db()
        self.assertEqual(run.sheets_coated, 5000)
        self.client.post(reverse('coating_records'), {'run_id': run.pk, 'reason': 'wrong job'})
        run.refresh_from_db()
        self.assertFalse(run.is_active)
        self.assertEqual(services.summary(self.card)['coated'], 0)

    def test_pages_render_and_export(self):
        self.post()
        self.assertEqual(self.client.get(reverse('coating_entry')).status_code, 200)
        page = self.client.get(reverse('coating_records'), {'from': '2026-10-01', 'to': '2026-10-31'})
        self.assertContains(page, 'Coater Chaudhry')
        self.assertEqual(self.client.get(reverse('coating_records'), {'tab': 'pending'}).status_code, 200)
        csv_response = self.client.get(reverse('coating_records'), {'from': '2026-10-01', 'to': '2026-10-31', 'export': 'csv'})
        self.assertIn('4000', csv_response.content.decode())

    def test_job_detail_shows_the_coating_panel(self):
        self.post()
        page = self.client.get(reverse('planning:job_detail', args=[self.job.pk]))
        self.assertContains(page, '4,000 of 10,000 printed sheets coated')
