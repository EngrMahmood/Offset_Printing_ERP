import datetime
import io
from unittest import mock

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from core.models import Dispatch, Production, Sorter, UserProfile
from planning.models import SkuRecipe
from planning.tests_merge import make_card, make_job

from .daily_dispatch_report import build_daily_report, item_key, render_xlsx

DAY = datetime.date(2026, 11, 11)


class DailyDispatchReportTests(TestCase):
    def setUp(self):
        self.sorter = Sorter.objects.create(name='Uzair')
        SkuRecipe.objects.create(sku='SKU-LBL', product_type='Label')
        SkuRecipe.objects.create(sku='SKU-INS', product_type='Insert Card', default_unit_cost='2.25')
        label = make_job('JC-11-26-PP-7001', 10000, sku='SKU-LBL', po_number='WO-1', department='PRODUCTION PILLOW', destination='SITE-5')
        self.label = make_card(label)
        self.label.is_print_job = False
        self.label.status = 'in_production'
        self.label.save(update_fields=['is_print_job', 'status'])
        insert = make_job('JC-11-26-PP-7002', 12020, sku='SKU-INS', po_number='WO-1')
        self.insert = make_card(insert)
        self.insert.is_print_job = False
        self.insert.status = 'in_production'
        self.insert.save(update_fields=['is_print_job', 'status'])

    def _pack(self, card, qty, waste):
        Production.objects.create(job_card=card, entry_type='packing', date=DAY, shift='A', packing_qty=qty,
                                  sorting_waste_qty=waste, sorter=self.sorter)

    def test_item_key_maps_product_types(self):
        self.assertEqual(item_key('Label'), 'label')
        self.assertEqual(item_key('Insert Card'), 'insert_card')
        self.assertEqual(item_key('Letter Head'), 'letter_head')
        self.assertEqual(item_key('Fluffing Instruction'), 'fluffing')
        self.assertEqual(item_key(''), 'other')

    def test_packing_table_totals_and_percentage(self):
        self._pack(self.label, 9750, 250)
        self._pack(self.insert, 12000, 0)
        report = build_daily_report(DAY)
        rows = {r['item']: r for r in report['packing_table']}
        self.assertEqual((rows['Taffeta Label']['qty'], rows['Taffeta Label']['waste'], rows['Taffeta Label']['pct']), (9750, 250, 2.56))
        self.assertEqual(report['packing_total']['qty'], 21750)
        self.assertEqual(len(report['packing_data']), 2)

    def test_dispatch_waste_is_the_gap_on_the_closing_dispatch_and_cost_uses_piece_cost(self):
        Dispatch.objects.create(job_card=self.insert, dc_no='PP-0001-26', dispatch_date=DAY, dispatch_qty=12000)
        self._pack(self.insert, 12000, 7)
        report = build_daily_report(DAY)
        row = [r for r in report['dispatch_table'] if r['item'] == 'Insert cards'][0]
        self.assertEqual(row['qty'], 12000)
        self.assertEqual(row['waste'], 20)            # 12020 ordered - 12000 dispatched, within the delivery tolerance
        self.assertEqual(row['cost'], 27000)          # 12000 x 2.25 from the SKU master
        self.assertEqual(row['physical'], 7)

    def test_partial_delivery_shows_no_theoretical_waste(self):
        Dispatch.objects.create(job_card=self.label, dc_no='PP-0002-26', dispatch_date=DAY, dispatch_qty=4000)
        report = build_daily_report(DAY)
        self.assertEqual([r for r in report['dispatch_table'] if r['item'] == 'Label'][0]['waste'], 0)

    def test_manual_flexo_adds_to_the_grand_total_only(self):
        self._pack(self.label, 10000, 0)
        report = build_daily_report(DAY, {'pack': '160000', 'disp': '5'})
        self.assertEqual(report['packing_total']['qty'], 10000)
        self.assertEqual(report['packing_grand'], 170000)
        self.assertEqual(report['dispatch_grand'], 5)

    def test_excel_has_the_three_sheets(self):
        from openpyxl import load_workbook

        self._pack(self.label, 10000, 0)
        wb = load_workbook(io.BytesIO(render_xlsx(build_daily_report(DAY))))
        self.assertEqual(wb.sheetnames, ['Daily packing & Dispatc Summary', 'Packing Report', 'Dispatch Report'])
        self.assertEqual(wb['Packing Report']['H2'].value, self.label.job_card_no)

    def test_page_and_download(self):
        user = get_user_model().objects.create_superuser('boss', password='x')
        self.client.force_login(user)
        if True:
            page = self.client.get(reverse('reports:daily_dispatch_report'), {'date': '2026-11-11'})
            self.assertEqual(page.status_code, 200)
            self.assertContains(page, 'Daily Packing Report')
            xlsx = self.client.get(reverse('reports:daily_dispatch_report_excel'), {'date': '2026-11-11'})
            self.assertEqual(xlsx.status_code, 200)
