import io

from django.contrib.auth.models import User
from django.test import TestCase

from .excel_io import (
    RAW_ITEM_HEADERS, commit_import_raw_items, download_raw_items_template,
    download_sku_bom_template, preview_import_raw_items,
)
from .models import ItemCategory, RawItem, UnitOfMeasure


def _build_workbook(headers, rows):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


class TemplateDownloadTests(TestCase):
    def test_raw_items_template_downloads(self):
        response = download_raw_items_template()
        self.assertEqual(response.status_code, 200)
        self.assertIn('attachment', response['Content-Disposition'])

    def test_sku_bom_template_downloads(self):
        response = download_sku_bom_template()
        self.assertEqual(response.status_code, 200)


class RawItemImportTests(TestCase):
    def setUp(self):
        self.uom = UnitOfMeasure.objects.create(code='KG', name='Kilogram')
        self.category = ItemCategory.objects.create(name='Ink', code='INK', default_uom=self.uom)
        self.user = User.objects.create_user('importer1', password='x')

    def test_dry_run_reports_errors_without_writing(self):
        rows = [['', 'Missing code item', 'Ink', 'ink', 'KG', '', '', '', '', '0', '', '1', '0', '0', '0', '1']]
        upload = _build_workbook(RAW_ITEM_HEADERS, rows)
        preview, errors = preview_import_raw_items(upload)
        self.assertTrue(errors)
        self.assertEqual(RawItem.objects.count(), 0)

    def test_unknown_category_reported_as_error(self):
        rows = [['INK-X', 'Test Ink', 'Nonexistent Category', 'ink', 'KG', '', '', '', '', '0', '', '1', '0', '0', '0', '1']]
        upload = _build_workbook(RAW_ITEM_HEADERS, rows)
        preview, errors = preview_import_raw_items(upload)
        self.assertTrue(any('Nonexistent Category' in e for e in errors))

    def test_commit_creates_raw_item(self):
        rows = [['INK-200', 'Magenta Ink', 'Ink', 'ink', 'KG', 'Spec X', 'BrandY', '', '', '250.5', '', '5', '10', '20', '100', '3']]
        upload = _build_workbook(RAW_ITEM_HEADERS, rows)
        result, errors = commit_import_raw_items(upload, user=self.user)
        self.assertEqual(errors, [])
        self.assertEqual(result['created'], 1)
        item = RawItem.objects.get(item_code='INK-200')
        self.assertEqual(item.name, 'Magenta Ink')
        self.assertEqual(str(item.unit_cost), '250.500000')

    def test_commit_updates_existing_item(self):
        RawItem.objects.create(item_code='INK-300', name='Old Name', category=self.category, uom=self.uom, unit_cost=1)
        rows = [['INK-300', 'New Name', 'Ink', 'ink', 'KG', '', '', '', '', '99', '', '1', '0', '0', '0', '1']]
        upload = _build_workbook(RAW_ITEM_HEADERS, rows)
        result, errors = commit_import_raw_items(upload, user=self.user)
        self.assertEqual(errors, [])
        self.assertEqual(result['updated'], 1)
        item = RawItem.objects.get(item_code='INK-300')
        self.assertEqual(item.name, 'New Name')

    def test_export_import_round_trip_is_lossless(self):
        from .excel_io import export_raw_items

        original = RawItem.objects.create(
            item_code='INK-400', name='Round Trip Ink', category=self.category, uom=self.uom,
            specification='Test spec', brand_grade='BrandZ', unit_cost=123, pack_size=2, moq=4,
            safety_stock=6, max_stock_level=100, lead_time_days=7,
        )
        response = export_raw_items(RawItem.objects.filter(pk=original.pk))
        upload = io.BytesIO(response.content)
        preview, errors = preview_import_raw_items(upload)
        self.assertEqual(errors, [])
        self.assertEqual(preview[0]['action'], 'update')
