from django.test import SimpleTestCase

from .po_extractor import (
	_detect_expected_line_count,
	_extract_best_sku_token,
	_extract_items_from_table_rows,
	_looks_like_sku_token,
)


class PoExtractorSkuGuardTests(SimpleTestCase):
	def test_iso_date_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('2026-03-12'))

	def test_slash_date_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('12/03/2026'))

	def test_textual_date_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('Mar 12, 2026'))

	def test_regular_sku_is_valid(self):
		self.assertTrue(_looks_like_sku_token('SKU-AB12-9901'))

	def test_header_word_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('Dated'))

	def test_generated_word_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('Generated'))

	def test_alphabetic_long_sku_is_valid(self):
		self.assertTrue(_looks_like_sku_token('LABELCAREUBMICROBIBERBEDSKIRT'))

	def test_extract_best_sku_ignores_dimension_fragment(self):
		raw = 'LABELCAREUBMICROBIBERBEDSKIRT / MATERIAL: TAFFETA SIZE: 95x45 MM'
		self.assertEqual(_extract_best_sku_token(raw), 'LABELCAREUBMICROBIBERBEDSKIRT')


class PoExtractorLineCountTests(SimpleTestCase):
	def test_detect_expected_line_count_from_table_rows(self):
		table_rows = [
			['#', 'SKU', 'Delivery Date', 'Qty'],
			['1', 'SKU-1001', 'Mar 12, 2026', '100 PIECE'],
			['2', 'SKU-1002', 'Mar 12, 2026', '200 PIECE'],
			['3', 'SKU-1003', 'Mar 12, 2026', '300 PIECE'],
		]
		self.assertEqual(_detect_expected_line_count('', table_rows), 3)

	def test_extract_items_from_table_rows(self):
		table_rows = [
			['#', 'SKU', 'Delivery Date', 'Qty', 'Unit Cost', 'Subtotal', 'GST', 'Net Total'],
			['1', 'SKU-1001', 'Mar 12, 2026', '100 PIECE', '10', '1000', '180', '1180'],
			['2', 'SKU-1002', 'Mar 12, 2026', '200 PIECE', '12', '2400', '432', '2832'],
		]
		items = _extract_items_from_table_rows(table_rows)
		self.assertEqual(len(items), 2)
		self.assertEqual(items[0]['sku'], 'SKU-1001')

	def test_extract_rs_two_row_per_item_layout(self):
		"""Mirrors the Utopia Rs PO layout: row A = serial+jobname, row B = SKU+data."""
		table_rows = [
			['#', 'SKU', 'DELIVERY DATE', 'QUANTITY', 'UNIT COST', 'SUBTOTAL', 'GST AMOUNT', 'NET TOTAL'],
			# item 1
			['1', 'IMPORTERLABEL-CA-AND-US / IMPORTERLABEL-CA-AND-US Material : Tafetta W-50.8 H-50.8mm', None, None, None, None, None, None],
			[None, 'IMPORTERLABEL-CA-AND-US', 'May 01, 2026', '1000000.0 PIECE', 'Rs 0.20', 'Rs 200,000.00', 'Rs 0.00', 'Rs 200,000.00'],
			# item 2
			['2', 'WARNINGLABEL-USA-CAN-IMPORTERLABEL / WARNINGLABEL-USA-CAN-IMPORTERLABEL White Adhesive Sticker W-101.6 L-76.2mm', None, None, None, None, None, None],
			[None, 'WARNINGLABEL-USA-CAN-IMPORTERLABEL', 'May 01, 2026', '300000.0 PIECE', 'Rs 1.20', 'Rs 360,000.00', 'Rs 0.00', 'Rs 360,000.00'],
			# item 3
			['3', 'LABELCAREUBMICROFIBERFITTEDQUEENMIG1 / MATERIAL: TAFFETA SIZE: 95x45 MM', None, None, None, None, None, None],
			[None, 'LABELCAREUBMICROFIBERFITTEDQUEENMIG1', 'May 01, 2026', '200000.0 PIECE', 'Rs 0.95', 'Rs 190,000.00', 'Rs 0.00', 'Rs 190,000.00'],
		]
		items = _extract_items_from_table_rows(table_rows)
		self.assertEqual(len(items), 3, f"Expected 3 items, got {len(items)}: {items}")
		self.assertEqual(items[0]['sku'], 'IMPORTERLABEL-CA-AND-US')
		self.assertAlmostEqual(float(items[0]['unit_cost']), 0.20, places=2)
		self.assertAlmostEqual(float(items[0]['quantity']), 1000000.0)
		self.assertEqual(items[1]['sku'], 'WARNINGLABEL-USA-CAN-IMPORTERLABEL')
		self.assertAlmostEqual(float(items[1]['unit_cost']), 1.20, places=2)
		self.assertEqual(items[2]['sku'], 'LABELCAREUBMICROFIBERFITTEDQUEENMIG1')
		self.assertAlmostEqual(float(items[2]['unit_cost']), 0.95, places=2)
