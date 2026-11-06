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


class CoatingBase(TestCase):
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


class CoatingTests(CoatingBase):
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
        self.assertContains(r, 'limit is 10,100')
        self.assertEqual(CoatingRun.objects.count(), 1)

    def test_front_and_back_coating_counts_two_passes_per_sheet(self):
        self.job.application = 'Front Back Matt Varnish'
        self.job.save(update_fields=['application'])
        self.card.refresh_from_db()
        r = self.post(sheets_coated=10000, coats=2)                 # one run, both sides: fine
        self.assertEqual(r.status_code, 302)
        info = services.summary(self.card)
        self.assertEqual((info['coated'], info['remaining'], info['sides']), (10000, 0, 2))
        r = self.post(sheets_coated=500, coats=2)                   # nothing left beyond the 1% tolerance
        self.assertEqual(r.status_code, 200)
        self.assertEqual(CoatingRun.objects.count(), 1)

    def test_front_and_back_as_two_single_runs(self):
        self.job.application = 'Front Back Matt Varnish'
        self.job.save(update_fields=['application'])
        self.card.refresh_from_db()
        self.assertEqual(self.post(sheets_coated=10000, coats=1).status_code, 302)    # front
        self.assertEqual(self.post(sheets_coated=10000, coats=1).status_code, 302)    # back
        self.assertEqual(services.summary(self.card)['remaining'], 0)

    def test_front_and_back_chosen_on_the_form_for_a_plain_varnish_job(self):
        self.job.application = 'MATT VARNISH'
        self.job.save(update_fields=['application'])
        self.card.refresh_from_db()
        self.assertEqual(self.post(sheets_coated=10000, sides=2).status_code, 302)    # front
        self.assertEqual(self.post(sheets_coated=10000, sides=2).status_code, 302)    # back
        self.assertEqual(services.summary(self.card)['remaining'], 0)
        self.assertEqual(self.post(sheets_coated=500, sides=2).status_code, 200)       # nothing left

    def test_error_page_hands_back_what_was_typed(self):
        r = self.post(sheets_coated=99999, waste_sheets=7)
        self.assertEqual(r.status_code, 200)
        self.assertContains(r, 'id="repost-data"')
        self.assertContains(r, '99999')
        self.assertContains(r, 'form_repost.js')

    def test_job_search(self):
        r = self.client.get(reverse('coating_job_search'), {'q': self.card.job_card_no})
        self.assertEqual([x['id'] for x in r.json()['results']], [self.card.pk])
        self.assertEqual(self.client.get(reverse('coating_job_search'), {'q': 'x'}).json()['results'], [])

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


class CoatingInReportsTests(CoatingBase):
    """Coating shows beside printing in the reports and never inside the printing figures."""

    def setUp(self):
        super().setUp()
        Production.objects.filter(job_card=self.card).update(operator=self.operator)
        self.post(sheets_coated=4000, waste_sheets=100)             # 4,100 coating impressions

    def q(self):
        return {'period': 'custom', 'date_from': '2026-10-01', 'date_to': '2026-10-31'}

    def test_daily_production_lines_up_printing_coating_and_total(self):
        from reports.services import build_daily_production_context
        from django.test import RequestFactory

        request = RequestFactory().get('/', self.q())
        request.user = self.user
        data = build_daily_production_context(request)
        row = next(r for r in data['overview_rows'] if str(r['date']) == DAY.isoformat())
        self.assertEqual((row['impressions'], row['coating_impressions'], row['total_impressions']), (10000, 4100, 14100))
        self.assertEqual((data['totals']['impressions'], data['totals']['coating_impressions'], data['totals']['total_impressions']), (10000, 4100, 14100))
        self.assertEqual(data['coating_by_machine'][0]['impressions'], 4100)
        machine = next(m for m in data['printing_by_machine'] if m['machine__name'] == 'GTO 1A')
        self.assertEqual((machine['impressions'], machine['coating_impressions'], machine['total_impressions']), (10000, 4100, 14100))
        data_coating = build_daily_production_context(RequestFactory().get('/', {**self.q(), 'tab': 'coating'}))
        self.assertEqual(data_coating['export_rows'][0]['impressions'], 4100)

    def test_daily_production_page_and_exports(self):
        page = self.client.get(reverse('reports:detail', args=['daily-production']), self.q())
        self.assertContains(page, 'Total Impressions')
        self.assertContains(page, 'dp-tab-coating')
        for tab in ('overview', 'coating'):
            for kind in ('xlsx', 'pdf'):
                r = self.client.get(reverse('reports:reports_api:export_report', args=['daily-production']), {**self.q(), 'type': kind, 'tab': tab})
                self.assertEqual(r.status_code, 200, (tab, kind))

    def test_coating_never_reaches_operator_score_or_printing_impressions(self):
        from reports import operator_efficiency as oe

        rows = oe.compute(DAY, DAY)['rows']
        self.assertEqual(rows[0]['impressions'], 10000)
        page = self.client.get(reverse('reports:detail', args=['operator-efficiency']), self.q())
        self.assertContains(page, 'Coating work (UV / varnish)')
        self.assertContains(page, '4,100')

    def test_sku_history_shows_coated_sheets(self):
        page = self.client.get(reverse('reports:detail', args=['sku-history']), {'sku': 'SKU-CT'})
        self.assertContains(page, 'Coated (sheets)')

    def test_hourly_report_mentions_coating(self):
        from reports import hourly_report

        day = datetime.datetime(2026, 10, 20, 12, 0, tzinfo=datetime.timezone.utc)
        report = hourly_report.build_hourly_report(day)
        self.assertEqual(report['printing']['coating_day_impressions'], 4100)
        self.assertIn('Coating (UV/varnish)', report['text'])

    def test_production_dashboard_shows_coating_beside_printing(self):
        page = self.client.get(reverse('production_dashboard'), {'start_date': '2026-10-01', 'end_date': '2026-10-31'})
        self.assertEqual(page.status_code, 200)
        self.assertEqual(page.context['coating_impressions'], 4100)
        self.assertEqual(page.context['all_impressions'], page.context['total_impressions'] + 4100)
