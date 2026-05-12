from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import SimpleTestCase, TestCase
from django.urls import reverse

from core.models import JobCard, UserProfile
from workflow.services import _sync_new_jobs_for_approved_sku

from .models import PlanningJob, PoDocument, SkuRecipe
from .services import _sync_repeat_jobs_from_po
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


class PlanningWorkflowSyncTests(TestCase):
	def setUp(self):
		self.user = get_user_model().objects.create_user(username='planner', password='testpass123')
		profile, _created = UserProfile.objects.get_or_create(user=self.user)
		profile.role = 'planner'
		profile.save(update_fields=['role'])

	def _create_po_document(self, sku='SKU-001', po_number='PO-001', quantity=1000, unit_cost='1.25'):
		payload = {
			'po_number': po_number,
			'po_date': '2026-05-01',
			'delivery_location': 'Main Warehouse',
			'department': 'Printing',
			'items': [
				{
					'line_no': 1,
					'sku': sku,
					'job_name': f'{sku} Job',
					'quantity': quantity,
					'unit_cost': unit_cost,
					'delivery_date': '2026-05-10',
				}
			],
		}
		return PoDocument.objects.create(
			po_file=SimpleUploadedFile('po.pdf', b'pdf-content', content_type='application/pdf'),
			extracted_payload=payload,
			uploaded_by=self.user,
		)

	def _create_approved_recipe(self, sku):
		return SkuRecipe.objects.create(
			sku=sku,
			job_name=f'{sku} Approved',
			material='Paper',
			color_spec='4+0',
			application='UV',
			machine_name='Machine A',
			ups=2,
			print_sheet_size='25x36',
			purchase_sheet_size='25x36',
			purchase_sheet_ups=2,
			purchase_material='Art Paper',
			default_unit_cost='1.40',
			daily_demand='100',
			awc_no='AWC-1',
			plate_set_no='PLATE-1',
			die_cutting='NO',
			master_data_status='approved',
			approved_by=self.user,
			created_by=self.user,
		)

	def test_po_sync_creates_draft_job_for_missing_recipe(self):
		po_doc = self._create_po_document(sku='NEW-SKU-001', po_number='PO-NEW-1')

		result = _sync_repeat_jobs_from_po(po_doc, actor=self.user)

		job = PlanningJob.objects.get(po_number='PO-NEW-1', sku='NEW-SKU-001')
		self.assertEqual(result['created'], 1)
		self.assertEqual(result['missing_recipe'], 1)
		self.assertEqual(job.status, 'draft')
		self.assertEqual(job.repeat_flag, 'New')
		self.assertIn('NEW SKU:', job.requirement)

	def test_approved_sku_sync_updates_existing_job_without_creating_duplicate(self):
		po_doc = self._create_po_document(sku='NEW-SKU-002', po_number='PO-NEW-2')
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		original_job = PlanningJob.objects.get(po_number='PO-NEW-2', sku='NEW-SKU-002')
		self._create_approved_recipe('NEW-SKU-002')

		result = _sync_new_jobs_for_approved_sku('NEW-SKU-002', actor=self.user)

		refreshed_job = PlanningJob.objects.get(id=original_job.id)
		self.assertEqual(result['created'], 0)
		self.assertEqual(result['updated'], 1)
		self.assertEqual(result.get('missing_jobs', 0), 0)
		self.assertEqual(PlanningJob.objects.filter(po_number='PO-NEW-2', sku='NEW-SKU-002').count(), 1)
		self.assertEqual(refreshed_job.repeat_flag, 'New')
		self.assertEqual(refreshed_job.material, 'Paper')
		self.assertEqual(refreshed_job.plate_set_no, 'PLATE-1')

	def test_approved_sku_sync_does_not_create_missing_planning_job(self):
		self._create_po_document(sku='NEW-SKU-003', po_number='PO-NEW-3')
		self._create_approved_recipe('NEW-SKU-003')

		result = _sync_new_jobs_for_approved_sku('NEW-SKU-003', actor=self.user)

		self.assertEqual(result['created'], 0)
		self.assertEqual(result['updated'], 0)
		self.assertEqual(result.get('missing_jobs', 0), 1)
		self.assertFalse(PlanningJob.objects.filter(po_number='PO-NEW-3', sku='NEW-SKU-003').exists())

	def test_submit_to_qc_blocks_when_recipe_not_approved(self):
		po_doc = self._create_po_document(sku='NEW-SKU-004', po_number='PO-NEW-4')
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		job = PlanningJob.objects.get(po_number='PO-NEW-4', sku='NEW-SKU-004')
		self.client.force_login(self.user)

		response = self.client.post(
			reverse('planning:job_status_update', args=[job.id]),
			{'transition': 'submit_qc'},
		)

		job.refresh_from_db()
		self.assertEqual(response.status_code, 302)
		self.assertEqual(job.status, 'draft')
		self.assertFalse(JobCard.objects.filter(planning_job=job).exists())

	def test_repeat_flag_preserved_after_approved_sku_refresh(self):
		po_doc = self._create_po_document(sku='REP-SKU-001', po_number='PO-REP-1')
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		job = PlanningJob.objects.get(po_number='PO-REP-1', sku='REP-SKU-001')
		job.repeat_flag = 'Repeat'
		job.save(update_fields=['repeat_flag', 'updated_at'])
		self._create_approved_recipe('REP-SKU-001')

		result = _sync_new_jobs_for_approved_sku('REP-SKU-001', actor=self.user)

		job.refresh_from_db()
		self.assertEqual(result['created'], 0)
		self.assertEqual(result['updated'], 1)
		self.assertEqual(job.repeat_flag, 'Repeat')

	def test_po_review_counts_use_master_data_not_existing_planning_jobs(self):
		po_doc = self._create_po_document(sku='NEW-SKU-006', po_number='PO-NEW-6')
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		job = PlanningJob.objects.filter(po_number='PO-NEW-6', sku='NEW-SKU-006').first()
		self.assertIsNotNone(job)
		self.client.force_login(self.user)

		response = self.client.get(reverse('qc:po_review', args=[po_doc.id]))
		self.assertEqual(response.status_code, 200)
		self.assertEqual(response.context['repeat_count'], 0)
		self.assertEqual(response.context['new_count'], 1)
		self.assertContains(response, 'Repeat SKUs: 0')
		self.assertContains(response, 'New SKUs: 1')

	def test_po_inbox_counts_use_master_data_not_existing_planning_jobs(self):
		po_doc = self._create_po_document(sku='INBOX-SKU-001', po_number='PO-INBOX-1')
		self._create_approved_recipe('INBOX-SKU-001')
		self.client.force_login(self.user)

		response = self.client.get(reverse('planning:po_inbox'))
		self.assertEqual(response.status_code, 200)
		rows = response.context['rows']
		self.assertTrue(any(row['po_number'] == 'PO-INBOX-1' for row in rows))
		row = next(row for row in rows if row['po_number'] == 'PO-INBOX-1')
		self.assertEqual(row['repeat_count'], 1)
		self.assertEqual(row['new_count'], 0)
		self.assertEqual(row['missing_count'], 0)

	def test_job_card_created_only_on_submit_to_qc(self):
		po_doc = self._create_po_document(sku='NEW-SKU-005', po_number='PO-NEW-5')
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		self._create_approved_recipe('NEW-SKU-005')
		_sync_new_jobs_for_approved_sku('NEW-SKU-005', actor=self.user)

		job = PlanningJob.objects.get(po_number='PO-NEW-5', sku='NEW-SKU-005')
		job.machine_name = 'Machine A'
		job.wastage_sheets = 12
		job.save(update_fields=['machine_name', 'wastage_sheets', 'updated_at'])
		self.assertFalse(JobCard.objects.filter(planning_job=job).exists())

		self.client.force_login(self.user)
		response = self.client.post(
			reverse('planning:job_status_update', args=[job.id]),
			{'transition': 'submit_qc'},
		)

		job.refresh_from_db()
		self.assertEqual(response.status_code, 302)
		self.assertEqual(job.status, 'pending_qc')
		self.assertEqual(JobCard.objects.filter(planning_job=job).count(), 1)

	def test_approved_sku_refresh_skips_non_draft_jobs(self):
		po_doc = self._create_po_document(sku='NEW-SKU-006', po_number='PO-NEW-6')
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		job = PlanningJob.objects.get(po_number='PO-NEW-6', sku='NEW-SKU-006')
		job.status = 'pending_qc'
		job.save(update_fields=['status', 'updated_at'])
		self._create_approved_recipe('NEW-SKU-006')

		result = _sync_new_jobs_for_approved_sku('NEW-SKU-006', actor=self.user)

		job.refresh_from_db()
		self.assertEqual(result['updated'], 0)
		self.assertEqual(result['locked'], 1)
		self.assertEqual(job.status, 'pending_qc')
