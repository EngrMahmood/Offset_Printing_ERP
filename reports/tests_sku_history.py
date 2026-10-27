import datetime

from django.contrib.auth import get_user_model
from django.test import RequestFactory, TestCase
from django.urls import reverse

from core.models import Dispatch, JobCard, Production
from planning.models import PlanningJob
from planning.tests_merge import make_card, make_job
from reports.sku_history import build_sku_history_context, find_matching_skus


class SkuHistoryTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user('sku_hist', password='pw12345678')
        self.rf = RequestFactory()

    def build(self, sku):
        request = self.rf.get('/reports/sku-history/', {'sku': sku})
        request.user = self.user
        return build_sku_history_context(request)

    def history_fixture(self):
        job1 = make_job('SH1', 1000, sku='SKU-HIST-A')
        card1 = make_card(job1)
        JobCard.objects.filter(pk=card1.pk).update(total_sheet_quantity=260, status='completed', po_date=datetime.date(2026, 7, 1))
        card1.refresh_from_db()
        Production.objects.create(job_card=card1, entry_type='printing', date='2026-07-21', shift='A',
                                  output_sheets=250, waste_sheets=10)
        Dispatch.objects.bulk_create([
            Dispatch(job_card=card1, dc_no='DC-1', dispatch_date=datetime.date(2026, 7, 25), dispatch_qty=900),
        ])
        planned_only = make_job('SH2', 500, sku='sku-hist-a')                    # same SKU, other spelling, no job card yet
        cancelled = make_job('SH3', 300, sku='SKU-HIST-A', status='cancelled')
        make_job('SH4', 100, sku='SKU-HIST-B')
        return job1, card1, planned_only, cancelled

    def test_nothing_searched_returns_an_empty_report(self):
        data = self.build('')
        self.assertEqual((data['sku'], data['rows'], data['matches']), ('', [], []))

    def test_partial_search_lists_matches_and_exact_search_opens_the_sku(self):
        self.history_fixture()
        self.assertEqual(find_matching_skus('hist'), ['SKU-HIST-A', 'SKU-HIST-B'])
        data = self.build('hist')
        self.assertEqual(data['sku'], '')
        self.assertEqual(data['matches'], ['SKU-HIST-A', 'SKU-HIST-B'])
        self.assertEqual(self.build('sku-hist-b')['sku'], 'SKU-HIST-B')                 # exact, any case
        self.assertEqual(self.build('hist-a')['sku'], 'SKU-HIST-A')                     # one match: opens it

    def test_job_card_rows_totals_and_links(self):
        job1, card1, planned_only, cancelled = self.history_fixture()
        data = self.build('SKU-HIST-A')
        by_jc = {r['jc_no']: r for r in data['rows']}
        row = by_jc[card1.job_card_no]
        self.assertEqual(row['url'], reverse('planning:job_detail', kwargs={'job_id': job1.id}))
        self.assertEqual((row['order_qty_pcs'], row['printed_pcs'], row['dispatched_pcs']), (1000, 1000, 900))
        self.assertEqual(row['printing_waste_pcs'], 40)                                   # 10 sheets x 4 ups
        self.assertEqual(row['planned_pcs'], 1040)
        self.assertEqual(row['dispatch_gap_pcs'], 140)                                    # 1040 planned - 900 dispatched
        self.assertEqual(row['total_wastage_pcs'], 180)
        self.assertEqual(row['wastage_status'], 'Finalized')
        self.assertEqual(row['pending_pcs'], 0)                                           # finalized: nothing owed
        self.assertEqual(row['lead_time_days'], 24)                                       # 1 Jul order -> 25 Jul dispatch
        waiting = by_jc['SH2']
        self.assertEqual((waiting['kind'], waiting['pending_pcs']), ('planning', 500))

        summary = data['summary']
        self.assertEqual(summary['jobs'], 2)
        self.assertEqual(summary['cancelled'], 1)                                         # SH3, left out of the totals
        self.assertEqual(summary['ordered_pcs'], 1500)
        self.assertEqual(summary['dispatched_pcs'], 900)
        self.assertEqual(summary['pending_pcs'], 500)
        self.assertEqual(summary['finalized_wastage_pcs'], 180)
        self.assertEqual(summary['finalized_wastage_pct'], round(180 * 100 / 1040, 2))
        self.assertEqual([d['dc_no'] for d in data['dispatch_rows']], ['DC-1'])
        self.assertEqual(data['dispatch_rows'][0]['dispatch_pcs'], 900)
        self.assertEqual(sum(m['jobs'] for m in data['monthly_rows']), 2)

    def test_export_rows_carry_the_same_data_with_extra_sheets(self):
        self.history_fixture()
        data = self.build('SKU-HIST-A')
        self.assertEqual(data['export_rows'], data['rows'])
        self.assertIn('jc_no', data['headers'])
        self.assertEqual([s['title'] for s in data['extra_sheets']], ['Dispatches', 'By Month'])
        self.assertTrue(any('SKU-HIST-A' in note for note in data['export_notes']))

    def test_page_renders_with_links_and_exports_download(self):
        job1, card1, planned_only, cancelled = self.history_fixture()
        admin = get_user_model().objects.create_superuser('sku_admin', password='pw12345678')
        self.client.force_login(admin)
        page = self.client.get(reverse('reports:detail', args=['sku-history']), {'sku': 'SKU-HIST-A'})
        self.assertEqual(page.status_code, 200)
        html = page.content.decode()
        self.assertIn(reverse('planning:job_detail', kwargs={'job_id': job1.id}), html)    # job card link
        self.assertIn('Total order received', html)
        self.assertIn('DC-1', html)

        empty = self.client.get(reverse('reports:detail', args=['sku-history']))
        self.assertEqual(empty.status_code, 200)
        picker = self.client.get(reverse('reports:detail', args=['sku-history']), {'sku': 'hist'})
        self.assertIn('SKUs match', picker.content.decode())

        for kind, content_type in (('xlsx', 'spreadsheetml'), ('csv', 'text/csv'), ('pdf', 'application/pdf')):
            response = self.client.get(
                reverse('reports:reports_api:export_report', args=['sku-history']), {'type': kind, 'sku': 'SKU-HIST-A'},
            )
            self.assertEqual(response.status_code, 200, kind)
            self.assertIn(content_type, response['Content-Type'])
            self.assertIn('sku-hist-a', response['Content-Disposition'].lower())
