from django.test import SimpleTestCase

from .po_extractor import _looks_like_sku_token


class PoExtractorSkuGuardTests(SimpleTestCase):
	def test_iso_date_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('2026-03-12'))

	def test_slash_date_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('12/03/2026'))

	def test_textual_date_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('Mar 12, 2026'))

	def test_regular_sku_is_valid(self):
		self.assertTrue(_looks_like_sku_token('SKU-AB12-9901'))
