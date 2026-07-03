from datetime import date
from decimal import Decimal

from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import SimpleTestCase, TestCase
from django.urls import reverse

from core.models import JobCard, UserProfile
from migration.models import MigrationImportJob, PlanningImportStaging
from migration.services.importer import (
    _import_planning_row,
    get_imported_planning_jobs,
    rollback_imported_planning_jobs,
)
from workflow.services import _po_payload_items, _sync_new_jobs_for_approved_sku

from .models import PlanningJob, PoDocument, SkuRecipe, JobCardChangeRequest
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

	def test_raw_item_description_can_be_used_as_sku(self):
		self.assertEqual(_extract_best_sku_token('A3 PAPER RIM'), 'A3 PAPER RIM')

	def test_header_word_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('Dated'))

	def test_generated_word_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('Generated'))

	def test_alphabetic_long_sku_is_valid(self):
		self.assertTrue(_looks_like_sku_token('LABELCAREUBMICROBIBERBEDSKIRT'))

	def test_year_token_is_not_sku(self):
		self.assertFalse(_looks_like_sku_token('2026'))

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

	def test_extract_two_row_item_with_blank_sku_cell_uses_job_name(self):
		table_rows = [
			['#', 'SKU', 'DELIVERY DATE', 'QUANTITY', 'UNIT COST', 'SUBTOTAL', 'GST AMOUNT', 'NET TOTAL'],
			['1', 'A3 PAPER RIM', None, None, None, None, None, None],
			[None, '', 'Jun 30, 2026', '4.0 PIECE', 'Rs 1,780.00', 'Rs 7,120.00', 'Rs 0.00', 'Rs 7,120.00'],
		]
		items = _extract_items_from_table_rows(table_rows)
		self.assertEqual(len(items), 1)
		self.assertEqual(items[0]['sku'], 'A3 PAPER RIM')
		self.assertEqual(items[0]['job_name'], 'A3 PAPER RIM')


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
			product_type='Label',
			ups=2,
			print_sheet_size='25x36',
			purchase_sheet_size='25x36',
			purchase_sheet_ups=2,
			default_unit_cost='1.40',
			daily_demand='100',
			awc_no='AWC-1',
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
		self.assertEqual(refreshed_job.plate_set_no, '')

	def test_integer_ups_in_recipe_persists_and_calculates_sheets(self):
		sku = 'DEC-SKU-001'
		SkuRecipe.objects.create(
			sku=sku,
			job_name=f'{sku} Approved',
			material='Paper',
			color_spec='4+0',
			application='UV',
			product_type='Label',
			ups=2,
			print_sheet_size='25x36',
			purchase_sheet_size='25x36',
			purchase_sheet_ups=2,
			default_unit_cost='1.40',
			daily_demand='100',
			awc_no='AWC-1',
			die_cutting='NO',
			master_data_status='approved',
			approved_by=self.user,
			created_by=self.user,
		)
		po_doc = self._create_po_document(sku=sku, po_number='PO-DEC-1', quantity=15)
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		result = _sync_new_jobs_for_approved_sku(sku, actor=self.user)

		job = PlanningJob.objects.get(po_number='PO-DEC-1', sku=sku)
		self.assertEqual(job.ups, 2)
		self.assertEqual(job.purchase_sheet_ups, 2)
		self.assertEqual(job.calculated_sheets_required, 8)
		self.assertEqual(job.calculated_purchase_sheet_required, 4)
		self.assertEqual(result['updated'], 1)

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

	def test_submit_to_qc_syncs_machine_plate_from_locked_sku_master(self):
		po_doc = self._create_po_document(sku='NEW-SKU-SYNC', po_number='PO-NEW-SYNC')
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		recipe = self._create_approved_recipe('NEW-SKU-SYNC')
		recipe.machine_name = 'Komori 6'
		recipe.plate_set_no = 'SET-99'
		recipe.save(update_fields=['machine_name', 'plate_set_no', 'updated_at'])
		_sync_new_jobs_for_approved_sku('NEW-SKU-SYNC', actor=self.user)

		job = PlanningJob.objects.get(po_number='PO-NEW-SYNC', sku='NEW-SKU-SYNC')
		job.wastage_sheets = 10
		job.purchase_material_origin = 'local'
		job.print_passes = 1
		job.machine_name = ''
		job.plate_set_no = ''
		job.save(update_fields=[
			'wastage_sheets', 'purchase_material_origin', 'print_passes',
			'machine_name', 'plate_set_no', 'updated_at',
		])
		self.client.force_login(self.user)

		response = self.client.post(
			reverse('planning:job_status_update', args=[job.id]),
			{'transition': 'submit_qc'},
		)

		job.refresh_from_db()
		self.assertEqual(response.status_code, 302)
		self.assertEqual(job.status, 'pending_qc')
		self.assertEqual(job.machine_name, 'Komori 6')
		self.assertEqual(job.plate_set_no, 'SET-99')

	def test_submit_to_qc_blocks_when_product_type_missing_on_approved_sku(self):
		po_doc = self._create_po_document(sku='NEW-SKU-PT', po_number='PO-NEW-PT')
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		recipe = self._create_approved_recipe('NEW-SKU-PT')
		recipe.product_type = ''
		recipe.save(update_fields=['product_type', 'updated_at'])
		_sync_new_jobs_for_approved_sku('NEW-SKU-PT', actor=self.user)

		job = PlanningJob.objects.get(po_number='PO-NEW-PT', sku='NEW-SKU-PT')
		job.machine_name = 'Machine A'
		job.wastage_sheets = 12
		job.plate_set_no = 'PLATE-1'
		job.purchase_material_origin = 'local'
		job.print_passes = 2
		job.save(update_fields=[
			'machine_name', 'wastage_sheets', 'plate_set_no',
			'purchase_material_origin', 'print_passes', 'updated_at',
		])
		self.client.force_login(self.user)

		response = self.client.post(
			reverse('planning:job_status_update', args=[job.id]),
			{'transition': 'submit_qc'},
		)

		job.refresh_from_db()
		self.assertEqual(response.status_code, 302)
		self.assertEqual(job.status, 'draft')
		follow_up = self.client.get(response.url)
		self.assertContains(follow_up, 'Product Type')

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

	def test_po_inbox_keeps_alphanumeric_skus_that_only_differ_by_trailing_digit(self):
		payload = {
			'po_number': 'PO-ALNUM-1',
			'expected_line_count': 4,
			'items': [
				{'line_no': 1, 'sku': 'LABELCAREUB7DPSFPOLYFILLFILLING1000GEU', 'quantity': 1500, 'delivery_date': '2026-06-24'},
				{'line_no': 2, 'sku': 'LABELCAREUB7DPSFPOLYFILLFILLING500GEU', 'quantity': 1000, 'delivery_date': '2026-06-24'},
				{'line_no': 3, 'sku': 'LABELCAREUB7DPSFPOLYFILLFILLING500GEU1', 'quantity': 1000, 'delivery_date': '2026-06-24'},
				{'line_no': 4, 'sku': 'LABELCAREUB7DPSFPOLYFILLFILLING1000GEU1', 'quantity': 1500, 'delivery_date': '2026-06-24'},
			],
		}

		items = _po_payload_items(payload)

		self.assertEqual(len(items), 4)
		self.assertEqual([item['sku'] for item in items], [line['sku'] for line in payload['items']])

	def test_po_inbox_supports_search_and_pagination(self):
		for index in range(25):
			self._create_po_document(sku=f'PAGE-SKU-{index}', po_number=f'PO-PAGE-{index}')
		self.client.force_login(self.user)

		response = self.client.get(reverse('planning:po_inbox') + '?q=PO-PAGE-1&per_page=10&page=1')
		self.assertEqual(response.status_code, 200)
		self.assertEqual(response.context['page_obj'].paginator.per_page, 10)
		self.assertGreaterEqual(response.context['page_obj'].paginator.count, 2)
		self.assertTrue(any('PO-PAGE-1' in row['po_number'] for row in response.context['rows']))

	def test_qc_user_can_view_pending_sku_master_entry_readonly(self):
		qc_user = get_user_model().objects.create_user(username='qc_user', password='testpass123')
		qc_profile, _ = UserProfile.objects.get_or_create(user=qc_user)
		qc_profile.role = 'qc'
		qc_profile.save(update_fields=['role'])

		po_doc = self._create_po_document(sku='QC-SKU-001', po_number='PO-QC-1')
		self.client.force_login(qc_user)

		response = self.client.get(
			reverse('planning:pending_sku_master_entry') + f'?po_doc_id={po_doc.id}&sku=QC-SKU-001&readonly=1'
		)
		self.assertEqual(response.status_code, 200)
		self.assertTrue(response.context['is_readonly'])
		self.assertContains(response, 'QC-SKU-001')

	def test_planner_can_access_pending_sku_master_entry(self):
		planner_user = get_user_model().objects.create_user(username='planner_user', password='testpass123')
		planner_profile, _ = UserProfile.objects.get_or_create(user=planner_user)
		planner_profile.role = 'planner'
		planner_profile.save(update_fields=['role'])

		po_doc = self._create_po_document(sku='PLAN-SKU-001', po_number='PO-PLAN-1')
		self.client.force_login(planner_user)

		response = self.client.get(
			reverse('planning:pending_sku_master_entry') + f'?po_doc_id={po_doc.id}&sku=PLAN-SKU-001'
		)
		self.assertEqual(response.status_code, 200)
		self.assertFalse(response.context['is_readonly'])
		self.assertContains(response, 'PLAN-SKU-001')

	def test_po_inbox_uploaded_timestamp_uses_document_created_time(self):
		po_doc = self._create_po_document(sku='INBOX-TIME-001', po_number='PO-TIME-1')
		self.client.force_login(self.user)

		response = self.client.get(reverse('planning:po_inbox'))
		self.assertEqual(response.status_code, 200)
		row = next(r for r in response.context['rows'] if r['po_number'] == 'PO-TIME-1')
		uploaded_local = row['uploaded']
		created_local = po_doc.created_at.astimezone().replace(tzinfo=None)
		self.assertAlmostEqual((uploaded_local - created_local).total_seconds(), 0, delta=5)

	def test_job_card_created_only_on_submit_to_qc(self):
		po_doc = self._create_po_document(sku='NEW-SKU-005', po_number='PO-NEW-5')
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		self._create_approved_recipe('NEW-SKU-005')
		_sync_new_jobs_for_approved_sku('NEW-SKU-005', actor=self.user)

		job = PlanningJob.objects.get(po_number='PO-NEW-5', sku='NEW-SKU-005')
		job.machine_name = 'Machine A'
		job.wastage_sheets = 12
		job.plate_set_no = 'PLATE-1'
		job.purchase_material_origin = 'local'
		job.print_passes = 2
		job.save(update_fields=['machine_name', 'wastage_sheets', 'plate_set_no', 'purchase_material_origin', 'print_passes', 'updated_at'])
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

	def test_migration_import_uses_po_received_date_and_order_qty(self):
		from unittest.mock import patch
		import datetime
		from django.utils import timezone

		import_job = MigrationImportJob.objects.create(
			module='PLANNING',
			sheet_url='http://example.com',
			status='STAGED',
			total_rows=1,
			created_by=self.user,
		)
		row = PlanningImportStaging.objects.create(
			import_job=import_job,
			row_number=1,
			po_number='PO-IMPORT-1',
			customer='Main Warehouse',
			sku='SKU-IMPORT-1',
			quantity=500,
			delivery_date=date(2026, 5, 10),
			raw_data={
				'po_number': 'PO-IMPORT-1',
				'po_received_date': '2026-05-10',
				'month': 'May',
				'date': '2026-05-10',
				'sku': 'SKU-IMPORT-1',
				'job_name': 'Imported Job',
				'quantity': '500',
				'ups': '2',
				'print_sheet_size': '25x36',
				'department': 'Printing',
				'delivery_location': 'Main Warehouse',
			},
		)

		with patch('django.utils.timezone.now') as mock_now:
			mock_now.return_value = timezone.make_aware(datetime.datetime(2026, 5, 10, 12, 0, 0))
			imported = _import_planning_row(row, actor=self.user)

		self.assertTrue(imported)
		job = PlanningJob.objects.get(po_number='PO-IMPORT-1', sku='SKU-IMPORT-1')
		self.assertEqual(job.order_qty, 500)
		self.assertEqual(job.plan_date, date(2026, 5, 10))
		self.assertEqual(job.plan_month, 'May')
		self.assertEqual(job.po_received_date, date(2026, 5, 10))

	def test_migration_import_updates_matching_existing_planning_job(self):
		from unittest.mock import patch
		import datetime
		from django.utils import timezone

		existing_job = PlanningJob.objects.create(
			jc_number='JC-EXIST',
			po_number='PO-IMPORT-2',
			sku='SKU-IMPORT-2',
			order_qty=100,
			status='draft',
		)
		import_job = MigrationImportJob.objects.create(
			module='PLANNING',
			sheet_url='http://example.com',
			status='STAGED',
			total_rows=1,
			created_by=self.user,
		)
		row = PlanningImportStaging.objects.create(
			import_job=import_job,
			row_number=1,
			po_number='PO-IMPORT-2',
			customer='Main Warehouse',
			sku='SKU-IMPORT-2',
			quantity=500,
			delivery_date=date(2026, 5, 10),
			raw_data={
				'po_number': 'PO-IMPORT-2',
				'po_received_date': '2026-05-10',
				'month': 'May',
				'date': '2026-05-10',
				'sku': 'SKU-IMPORT-2',
				'job_name': 'Imported Job',
				'quantity': '500',
				'ups': '2',
				'print_sheet_size': '25x36',
				'department': 'Printing',
				'delivery_location': 'Main Warehouse',
				'jc_number': 'JC-OTHER',
			},
		)

		with patch('django.utils.timezone.now') as mock_now:
			mock_now.return_value = timezone.make_aware(datetime.datetime(2026, 5, 10, 12, 0, 0))
			imported = _import_planning_row(row, actor=self.user)

		self.assertTrue(imported)
		refreshed = PlanningJob.objects.get(id=existing_job.id)
		self.assertEqual(refreshed.order_qty, 500)
		self.assertEqual(refreshed.plan_date, date(2026, 5, 10))
		self.assertEqual(refreshed.plan_month, 'May')
		self.assertEqual(refreshed.jc_number, 'JC-EXIST')
		self.assertEqual(PlanningJob.objects.filter(po_number='PO-IMPORT-2', sku='SKU-IMPORT-2').count(), 1)
		row.refresh_from_db()
		self.assertEqual(row.imported_reference, 'JC-EXIST')

	def test_rollback_imported_planning_jobs_deletes_imported_records_only(self):
		import_job = MigrationImportJob.objects.create(
			module='PLANNING',
			sheet_url='http://example.com',
			status='STAGED',
			total_rows=1,
			created_by=self.user,
		)
		row = PlanningImportStaging.objects.create(
			import_job=import_job,
			row_number=1,
			po_number='PO-ROLLBACK-1',
			customer='Main Warehouse',
			sku='SKU-ROLLBACK-1',
			quantity=100,
			delivery_date=date(2026, 5, 11),
			raw_data={
				'po_number': 'PO-ROLLBACK-1',
				'po_date': '2026-05-11',
				'month': 'May',
				'sku': 'SKU-ROLLBACK-1',
				'job_name': 'Rollback Job',
				'quantity': '100',
				'ups': '2',
				'print_sheet_size': '25x36',
				'department': 'Printing',
				'delivery_location': 'Main Warehouse',
			},
		)
		imported = _import_planning_row(row, actor=self.user)
		self.assertTrue(imported)
		job = PlanningJob.objects.get(po_number='PO-ROLLBACK-1', sku='SKU-ROLLBACK-1')
		self.assertEqual(job.order_qty, 100)
		self.assertEqual(rollback_imported_planning_jobs(import_job, dry_run=True), 1)
		self.assertEqual(get_imported_planning_jobs(import_job)[0].id, job.id)
		deleted = rollback_imported_planning_jobs(import_job)
		self.assertEqual(deleted, 1)
		self.assertFalse(PlanningJob.objects.filter(id=job.id).exists())


class MasterDataSyncTests(TestCase):
	def setUp(self):
		self.user = get_user_model().objects.create_user(username='planner', password='testpass123')
		profile, _created = UserProfile.objects.get_or_create(user=self.user)
		profile.role = 'planner'
		profile.save(update_fields=['role'])

	def _create_job(self, sku='SKU-SYNC-1', ups=2, status='pending_qc', jc_number='JC-SYNC-1'):
		return PlanningJob.objects.create(
			jc_number=jc_number,
			po_number='PO-SYNC-1',
			sku=sku,
			job_name='Sync Job',
			order_qty=1000,
			ups=ups,
			print_sheet_size='20x30',
			purchase_sheet_size='20x30',
			purchase_sheet_ups=2,
			wastage_sheets=10,
			status=status,
		)

	def _create_approved_recipe(self, sku='SKU-SYNC-1', ups=4):
		return SkuRecipe.objects.create(
			sku=sku,
			job_name='Sync Job Approved',
			material='Paper',
			color_spec='4+0',
			application='UV',
			product_type='Label',
			ups=ups,
			print_sheet_size='25x36',
			purchase_sheet_size='25x36',
			purchase_sheet_ups=2,
			default_unit_cost='1.40',
			daily_demand='100',
			awc_no='AWC-1',
			die_cutting='NO',
			master_data_status='approved',
			approved_by=self.user,
			created_by=self.user,
		)

	def test_request_and_apply_master_sync_updates_only_requested_job(self):
		self._create_approved_recipe()
		job = self._create_job()
		completed_job = self._create_job(sku='SKU-SYNC-OLD', ups=2, status='completed', jc_number='JC-SYNC-OLD')
		self._create_approved_recipe(sku='SKU-SYNC-OLD', ups=8)

		from planning.services import (
			apply_master_data_sync,
			can_request_master_data_sync,
			get_master_data_field_diffs,
			request_master_data_sync,
		)

		self.assertTrue(can_request_master_data_sync(job))
		self.assertTrue(get_master_data_field_diffs(job))
		request_master_data_sync(job, actor=self.user, reason='UPS revised in master')

		job.refresh_from_db()
		self.assertTrue(job.master_sync_requested)
		self.assertEqual(job.ups, 2)

		job, result = apply_master_data_sync(job, actor=self.user)
		self.assertIn('ups', result['updated_fields'])
		self.assertEqual(job.ups, 4)
		self.assertEqual(job.print_sheet_size, '25x36')
		self.assertFalse(job.master_sync_requested)
		self.assertEqual(job.calculated_sheets_required, 260)

		completed_job.refresh_from_db()
		self.assertEqual(completed_job.ups, 2)

	def test_completed_job_cannot_request_master_sync(self):
		self._create_approved_recipe()
		job = self._create_job(status='completed')
		from planning.services import can_request_master_data_sync, request_master_data_sync

		self.assertFalse(can_request_master_data_sync(job))
		with self.assertRaises(ValueError):
			request_master_data_sync(job, actor=self.user, reason='Should fail')

	def test_reopen_and_apply_master_sync_for_released_job(self):
		from core.models import JobCard, Machine
		from planning.services import get_master_data_field_diffs, reopen_and_apply_master_data_sync

		machine = Machine.objects.create(name='Test Machine')
		recipe = self._create_approved_recipe(sku='SKU-REL-1', ups=4)
		job = self._create_job(sku='SKU-REL-1', ups=2, status='released', jc_number='JC-REL-1')
		job.machine_name = machine.name
		job.plate_set_no = 'PLATE-1'
		job.save(update_fields=['machine_name', 'plate_set_no', 'updated_at'])
		JobCard.objects.create(
			planning_job=job,
			job_card_no='JC-REL-1',
			SKU='SKU-REL-1',
			order_qty=1000,
			ups=2,
			print_sheet_size='20x30',
			po_date=date.today(),
			plate_set_no='PLATE-1',
			machine_name=machine,
			status='released',
		)
		recipe.ups = Decimal('4')
		recipe.print_sheet_size = '25x36'
		recipe.purchase_sheet_ups = Decimal('2')
		recipe.save(update_fields=['ups', 'print_sheet_size', 'purchase_sheet_ups', 'updated_at'])

		self.assertTrue(get_master_data_field_diffs(job))
		job, result = reopen_and_apply_master_data_sync(job, actor=self.user, reason='UPS revised')

		job.refresh_from_db()
		self.assertEqual(job.ups, Decimal('4'))
		self.assertEqual(job.status, 'draft')
		self.assertIn('ups', result['updated_fields'])
		self.assertTrue(result['reopened_job_card'])
		job.job_card.refresh_from_db()
		self.assertEqual(job.job_card.ups, Decimal('4'))
		self.assertIn(job.job_card.workflow_status, {'draft', 'pending_data'})


class PendingSkuMasterEntryPlannerTests(TestCase):
	def setUp(self):
		self.user = get_user_model().objects.create_user(username='planner', password='testpass123')
		profile, _created = UserProfile.objects.get_or_create(user=self.user)
		profile.role = 'planner'
		profile.save(update_fields=['role'])
		self.client.force_login(self.user)

	def _create_po_document(self, sku='TEST-SKU-100', po_number='PO-100'):
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
					'quantity': 1000,
					'unit_cost': '1.25',
					'delivery_date': '2026-05-10',
					'remarks': 'Mocked PO Remarks',
				}
			],
		}
		from django.core.files.uploadedfile import SimpleUploadedFile
		from planning.models import PoDocument
		return PoDocument.objects.create(
			po_file=SimpleUploadedFile('po.pdf', b'pdf-content', content_type='application/pdf'),
			extracted_payload=payload,
			uploaded_by=self.user,
		)

	def test_pending_sku_master_entry_designer_fields_disabled(self):
		po_doc = self._create_po_document()
		from django.urls import reverse
		url = reverse('planning:pending_sku_master_entry') + f"?po_doc_id={po_doc.id}&sku=TEST-SKU-100"
		response = self.client.get(url)
		self.assertEqual(response.status_code, 200)
		form = response.context['form']
		
		# Verify that remarks display in form remarks initial value
		self.assertEqual(form.initial.get('remarks'), 'Mocked PO Remarks')
		
		# Designer fields (including Print Color) are disabled for planner on first draft
		self.assertTrue(form.fields['color_spec'].disabled)
		self.assertTrue(form.fields['size_w_mm'].disabled)
		self.assertTrue(form.fields['ups'].disabled)
		self.assertTrue(form.fields['die_cutting'].disabled)
		self.assertEqual(form.fields['color_spec'].sku_role, 'designer')
		
		# Planner fields (including Job Process) are editable
		self.assertFalse(form.fields['material'].disabled)
		self.assertFalse(form.fields['application'].disabled)
		self.assertFalse(form.fields['product_type'].disabled)
		self.assertFalse(form.fields['job_process_type'].disabled)
		self.assertEqual(form.fields['job_process_type'].sku_role, 'planner')

	def test_pending_sku_master_entry_send_to_plate_making(self):
		from core.models import Machine

		Machine.objects.create(name='Machine A')
		po_doc = self._create_po_document()
		from planning.services import _sync_repeat_jobs_from_po
		_sync_repeat_jobs_from_po(po_doc, actor=self.user)
		
		from planning.models import PlanningJob, SkuRecipe
		job = PlanningJob.objects.get(po_number='PO-100', sku='TEST-SKU-100')
		job.machine_name = 'Machine A'
		job.save(update_fields=['machine_name'])

		from django.urls import reverse
		url = reverse('planning:pending_sku_master_entry')
		
		# Submit form with action "send_to_plate_making"
		response = self.client.post(url, {
			'po_doc_id': po_doc.id,
			'sku': 'TEST-SKU-100',
			'material': 'Paper',
			'application': 'UV',
			'machine_name': 'Machine A',
			'action': 'send_to_plate_making',
		})
		
		self.assertEqual(response.status_code, 302)
		
		# Verify SkuRecipe is saved as draft
		recipe = SkuRecipe.objects.get(sku='TEST-SKU-100')
		self.assertEqual(recipe.master_data_status, 'draft')
		self.assertEqual(recipe.material, 'Paper')
		self.assertEqual(recipe.application, 'UV')

		# Verify PlanningJob copied planner fields and transitioned to new_plate_making
		job.refresh_from_db()
		self.assertEqual(job.material, 'Paper')
		self.assertEqual(job.application, 'UV')
		self.assertEqual(job.planning_stage, 'new_plate_making')

	def test_pending_sku_master_entry_admin_fields_enabled(self):
		# Create an admin user
		admin_user = get_user_model().objects.create_user(username='admin_user', password='testpass123')
		profile, _created = UserProfile.objects.get_or_create(user=admin_user)
		profile.role = 'admin'
		profile.save(update_fields=['role'])
		self.client.force_login(admin_user)

		po_doc = self._create_po_document()
		from django.urls import reverse
		url = reverse('planning:pending_sku_master_entry') + f"?po_doc_id={po_doc.id}&sku=TEST-SKU-100"
		response = self.client.get(url)
		self.assertEqual(response.status_code, 200)
		form = response.context['form']
		
		# Verify that designer fields are NOT disabled for the admin
		self.assertFalse(form.fields['color_spec'].disabled)
		self.assertFalse(form.fields['size_w_mm'].disabled)
		self.assertFalse(form.fields['ups'].disabled)
		self.assertFalse(form.fields['die_cutting'].disabled)
		self.assertFalse(form.fields['plate_set_no'].disabled)

	def test_pending_sku_master_entry_shows_role_legend(self):
		po_doc = self._create_po_document()
		from django.urls import reverse
		url = reverse('planning:pending_sku_master_entry') + f"?po_doc_id={po_doc.id}&sku=TEST-SKU-100"
		response = self.client.get(url)
		self.assertEqual(response.status_code, 200)
		self.assertContains(response, 'sku-recipe-legend')
		self.assertContains(response, 'Planner fields')
		self.assertContains(response, 'Designer fields')
		self.assertEqual(response.context['sku_recipe_viewer_role'], 'planner')


class SkuRecipeFormRolePermissionTests(TestCase):
	def test_product_type_required_for_qc_submission(self):
		from planning.models import SkuRecipe
		from workflow.services import _missing_required_master_fields

		recipe = SkuRecipe(
			sku='SKU-PT-1',
			job_name='Test Job',
			material='Paper',
			color_spec='4 color',
			application='UV',
			product_type='',
			print_sheet_size='720x1020',
			purchase_sheet_size='720x1020',
			ups=4,
			die_cutting='YES',
		)
		missing = _missing_required_master_fields(recipe)
		self.assertIn('Product Type', missing)

	def test_graphics_designer_cannot_edit_product_type(self):
		from planning.forms import SkuRecipeForm
		from planning.services import apply_sku_recipe_form_role_permissions, get_sku_recipe_form_ui_context

		designer_user = get_user_model().objects.create_user(username='designer_role_user', password='testpass123')
		designer_profile, _ = UserProfile.objects.get_or_create(user=designer_user)
		designer_profile.role = 'graphics_designer'
		designer_profile.save(update_fields=['role'])
		designer_user.refresh_from_db()

		form = SkuRecipeForm()
		apply_sku_recipe_form_role_permissions(form, designer_user)
		self.assertTrue(form.fields['product_type'].disabled)
		self.assertFalse(form.fields['color_spec'].disabled)
		self.assertEqual(form.fields['color_spec'].sku_role, 'designer')
		self.assertTrue(form.fields['color_spec'].sku_is_mine)
		self.assertEqual(form.fields['material'].sku_role, 'planner')
		self.assertFalse(form.fields['material'].sku_is_mine)
		self.assertEqual(form.fields['job_process_type'].sku_role, 'planner')
		self.assertTrue(form.fields['job_process_type'].disabled)
		self.assertFalse(form.fields['size_w_mm'].disabled)
		self.assertEqual(form.fields['size_w_mm'].sku_role, 'designer')

		ui = get_sku_recipe_form_ui_context(designer_user)
		self.assertEqual(ui['sku_recipe_viewer_role'], 'designer')

	def test_planner_cannot_edit_designer_fields_on_first_draft(self):
		from planning.forms import SkuRecipeForm
		from planning.models import SkuRecipe
		from planning.services import apply_sku_recipe_form_role_permissions

		planner_user = get_user_model().objects.create_user(username='planner_role_user', password='testpass123')
		planner_profile, _ = UserProfile.objects.get_or_create(user=planner_user)
		planner_profile.role = 'planner'
		planner_profile.save(update_fields=['role'])
		planner_user.refresh_from_db()

		recipe = SkuRecipe(sku='SKU-FIRST', job_name='First Job', master_data_status='draft')
		form = SkuRecipeForm(instance=recipe)
		apply_sku_recipe_form_role_permissions(form, planner_user, recipe=recipe)
		self.assertTrue(form.fields['print_sheet_size'].disabled)
		self.assertFalse(form.fields['machine_name'].disabled)
		self.assertTrue(form.fields['sku'].disabled)
		self.assertTrue(form.fields['job_name'].disabled)

	def test_sku_and_job_name_locked_for_non_admin(self):
		from planning.forms import SkuRecipeForm
		from planning.models import SkuRecipe
		from planning.services import apply_sku_recipe_form_role_permissions

		for role, username in (
			('planner', 'planner_lock_user'),
			('graphics_designer', 'designer_lock_user'),
		):
			user = get_user_model().objects.create_user(username=username, password='testpass123')
			profile, _ = UserProfile.objects.get_or_create(user=user)
			profile.role = role
			profile.save(update_fields=['role'])
			user.refresh_from_db()
			recipe = SkuRecipe(sku='SKU-LOCK', job_name='Locked Job', master_data_status='draft')
			form = SkuRecipeForm(instance=recipe)
			apply_sku_recipe_form_role_permissions(form, user, recipe=recipe)
			self.assertTrue(form.fields['sku'].disabled, role)
			self.assertTrue(form.fields['job_name'].disabled, role)

	def test_planner_can_edit_designer_fields_after_reopen(self):
		from django.utils import timezone
		from planning.forms import SkuRecipeForm
		from planning.models import SkuRecipe
		from planning.services import (
			apply_sku_recipe_form_role_permissions,
			build_plate_remake_warning,
			get_plate_remake_impact_changes,
		)

		planner_user = get_user_model().objects.create_user(username='planner_reopen_user', password='testpass123')
		planner_profile, _ = UserProfile.objects.get_or_create(user=planner_user)
		planner_profile.role = 'planner'
		planner_profile.save(update_fields=['role'])
		planner_user.refresh_from_db()

		recipe = SkuRecipe(
			sku='SKU-REOPEN',
			master_data_status='draft',
			rejection_comment='Reopened for machine change',
			last_rejected_at=timezone.now(),
		)
		form = SkuRecipeForm(instance=recipe)
		apply_sku_recipe_form_role_permissions(form, planner_user, recipe=recipe)
		self.assertFalse(form.fields['print_sheet_size'].disabled)
		self.assertFalse(form.fields['machine_name'].disabled)
		self.assertTrue(form.fields['print_sheet_size'].sku_is_mine)

		changed = get_plate_remake_impact_changes(
			{'machine_name': 'SM74', 'print_sheet_size': '9x15'},
			{'machine_name': 'GTO', 'print_sheet_size': '9x15'},
		)
		self.assertEqual(changed, ['Machine'])
		warning = build_plate_remake_warning(changed, context='sync')
		self.assertIn('Machine', warning)
		self.assertIn('Request plates', warning)


class PoRemarksExtractorTests(TestCase):
	def test_remarks_extraction_from_po_text(self):
		from planning.po_extractor import _parse_po_text
		
		# Mock Utopia PDF two-row layout format text
		text = (
			"PURCHASE ORDER PO-04-2026-164283\n"
			"Dated Apr 10, 2026\n"
			"Approval Date Apr 10, 2026\n"
			"Department/Broker Offset Printing\n"
			"Delivery Location SITE-2\n"
			"SUPPLIER DETAILS Name Supplier A NTN 123\n"
			"BUYER DETAILS Name Buyer B STRN 456\n"
			"# SKU\n"
			"1 IMPORTERLABEL-CA-AND-US / IMPORTERLABEL-CA-AND-US Material : Tafetta W-50.8 H-50.8mm\n"
			"None IMPORTERLABEL-CA-AND-US May 01, 2026 1000000.0 PIECE Rs 0.20 Rs 200,000.00 Rs 0.00 Rs 200,000.00\n"
			"2 WARNINGLABEL-USA-CAN-IMPORTERLABEL / White Adhesive Sticker W-101.6 L-76.2mm\n"
			"None WARNINGLABEL-USA-CAN-IMPORTERLABEL May 01, 2026 300000.0 PIECE Rs 1.20 Rs 360,000.00 Rs 0.00 Rs 360,000.00\n"
			"GRAND TOTAL Rs 560,000.00\n"
		)
		
		table_rows = [
			["1", "IMPORTERLABEL-CA-AND-US / IMPORTERLABEL-CA-AND-US Material : Tafetta W-50.8 H-50.8mm"],
			["None", "IMPORTERLABEL-CA-AND-US", "May 01, 2026", "1000000.0 PIECE", "Rs 0.20", "Rs 200,000.00", "Rs 0.00", "Rs 200,000.00"],
			["2", "WARNINGLABEL-USA-CAN-IMPORTERLABEL / White Adhesive Sticker W-101.6 L-76.2mm"],
			["None", "WARNINGLABEL-USA-CAN-IMPORTERLABEL", "May 01, 2026", "300000.0 PIECE", "Rs 1.20", "Rs 360,000.00", "Rs 0.00", "Rs 360,000.00"],
		]
		
		result = _parse_po_text(text, table_blobs=[], table_rows=table_rows)
		self.assertEqual(len(result['items']), 2)
		
		item1 = result['items'][0]
		self.assertEqual(item1['job_name'], 'IMPORTERLABEL-CA-AND-US')
		self.assertEqual(item1['remarks'], 'IMPORTERLABEL-CA-AND-US Material : Tafetta W-50.8 H-50.8mm')
		
		item2 = result['items'][1]
		self.assertEqual(item2['job_name'], 'WARNINGLABEL-USA-CAN-IMPORTERLABEL')
		self.assertEqual(item2['remarks'], 'White Adhesive Sticker W-101.6 L-76.2mm')


class JobCardChangeRequestTests(TestCase):
	def setUp(self):
		self.planner = get_user_model().objects.create_user(username='planner', password='testpass123')
		self.pm = get_user_model().objects.create_user(username='pm', password='testpass123')
		self.operator = get_user_model().objects.create_user(username='operator', password='testpass123')
		
		planner_profile, _ = UserProfile.objects.get_or_create(user=self.planner)
		planner_profile.role = 'planner'
		planner_profile.save()
		
		pm_profile, _ = UserProfile.objects.get_or_create(user=self.pm)
		pm_profile.role = 'production_manager'
		pm_profile.save()
		
		operator_profile, _ = UserProfile.objects.get_or_create(user=self.operator)
		operator_profile.role = 'operator'
		operator_profile.save()
		
		from core.models import Machine
		self.machine1 = Machine.objects.create(name='Machine-1', standard_impressions_per_hour=4000)
		self.machine2 = Machine.objects.create(name='Machine-2', standard_impressions_per_hour=5000)
		
		self.job = PlanningJob.objects.create(
			jc_number='JC-TEST-CRM',
			po_number='PO-TEST-CRM',
			sku='SKU-CRM-1',
			order_qty=1000,
			ups=2,
			wastage_sheets=20,
			machine_name=self.machine1.name,
			status='released',
		)
		
		self.job_card = JobCard.objects.create(
			planning_job=self.job,
			job_card_no='JC-TEST-CRM',
			SKU='SKU-CRM-1',
			order_qty=1000,
			ups=2,
			wastage=20,
			machine_name=self.machine1,
			status='released',
			po_date=date.today(),
			total_colors=4,
			plate_set_no='PLATE-SET-1',
		)

	def test_request_wastage_machine_change_requires_planner(self):
		self.client.force_login(self.operator)
		response = self.client.post(
			reverse('planning:request_wastage_machine_change', args=[self.job.id]),
			{
				'reason': 'Planner requested reopening',
			}
		)
		self.assertEqual(response.status_code, 302)
		self.assertFalse(JobCardChangeRequest.objects.filter(planning_job=self.job).exists())

	def test_request_wastage_machine_change_success(self):
		self.client.force_login(self.planner)
		response = self.client.post(
			reverse('planning:request_wastage_machine_change', args=[self.job.id]),
			{
				'reason': 'Test reopen request',
			}
		)
		self.assertEqual(response.status_code, 302)
		change_req = JobCardChangeRequest.objects.get(planning_job=self.job)
		self.assertEqual(change_req.status, 'pending')
		self.assertEqual(change_req.request_type, 'reopen_to_draft')
		self.assertEqual(change_req.reason, 'Test reopen request')

	def test_pm_can_approve_change_request(self):
		change_req = JobCardChangeRequest.objects.create(
			planning_job=self.job,
			request_type='reopen_to_draft',
			reason='Need to edit wastage & machine',
			requested_by=self.planner
		)
		
		self.client.force_login(self.pm)
		response = self.client.post(reverse('planning:approve_change_request', args=[change_req.id]))
		self.assertEqual(response.status_code, 302)
		
		change_req.refresh_from_db()
		self.assertEqual(change_req.status, 'approved')
		
		self.job.refresh_from_db()
		self.assertEqual(self.job.status, 'draft')
		
		self.job_card.refresh_from_db()
		self.assertEqual(self.job_card.workflow_status, 'draft')
		
		from core.models import ChangeLog
		self.assertTrue(ChangeLog.objects.filter(entity_type='job_card', record_id=self.job_card.id).exists())

	def test_pm_can_reject_change_request(self):
		change_req = JobCardChangeRequest.objects.create(
			planning_job=self.job,
			request_type='reopen_to_draft',
			reason='Need to edit wastage & machine',
			requested_by=self.planner
		)
		
		self.client.force_login(self.pm)
		response = self.client.post(
			reverse('planning:reject_change_request', args=[change_req.id]),
			{'rejection_reason': 'Invalid reasons'}
		)
		self.assertEqual(response.status_code, 302)
		
		change_req.refresh_from_db()
		self.assertEqual(change_req.status, 'rejected')
		self.assertEqual(change_req.rejection_reason, 'Invalid reasons')
		
		self.job.refresh_from_db()
		self.assertEqual(self.job.status, 'released')
		self.job_card.refresh_from_db()
		self.assertEqual(self.job_card.workflow_status, 'released')


class PlanningJobImpressionCalculationTests(TestCase):
	def setUp(self):
		self.user = get_user_model().objects.create_user(username='impression_planner', password='testpass123')

	def _create_job(self, order_qty=1000, ups=2, wastage_sheets=12):
		return PlanningJob.objects.create(
			jc_number='JC-IMP-001',
			sku='IMP-SKU-1',
			job_name='Impression Job',
			order_qty=order_qty,
			ups=ups,
			wastage_sheets=wastage_sheets,
			status='draft',
			created_by=self.user,
		)

	def test_calculated_planned_total_impressions_uses_sheets_times_passes(self):
		job = self._create_job()
		job.print_passes = 3
		self.assertEqual(job.calculated_sheets_required, 512)
		self.assertEqual(job.calculated_planned_total_impressions, 1536)

	def test_save_syncs_planned_total_impressions(self):
		job = self._create_job()
		job.print_passes = 2
		job.save()
		job.refresh_from_db()
		self.assertEqual(job.planned_total_impressions, 1024)

	def test_finalization_form_rejects_missing_passes(self):
		from planning.forms import PlanningJobFinalizationForm

		job = self._create_job()
		form = PlanningJobFinalizationForm(
			data={
				'delivery_date': '2026-05-10',
				'wastage_sheets': 12,
				'purchase_material_origin': 'local',
				'destination': 'Main Warehouse',
				'requirement': 'Standard run',
			},
			instance=job,
		)
		self.assertFalse(form.is_valid())
		self.assertIn('print_passes', form.errors)

	def test_job_card_uses_planned_impressions_not_raw_sheets(self):
		from core.jobcard_service import ensure_job_card_from_planning_job

		job = self._create_job()
		job.print_passes = 2
		job.save()
		job_card, _ = ensure_job_card_from_planning_job(job, actor=self.user)
		self.assertEqual(job_card.total_impressions_required, 1024)
		self.assertEqual(job_card.impression_pass_multiplier, 1)


class SkuRecipePlanningSyncTests(TestCase):
	def setUp(self):
		self.user = get_user_model().objects.create_user(username='sku_sync_user', password='pass')
		self.planning_job = PlanningJob.objects.create(
			jc_number='JC-SYNC-001',
			sku='SKU-SYNC-001',
			job_name='Sync Test SKU',
			color_spec='4',
			size_w_mm=100,
			size_h_mm=150,
			ups=12,
			print_sheet_size='20x30',
			purchase_sheet_size='20x30',
			plate_set_no='SET-99',
			repeat_flag='New',
			status='draft',
			created_by=self.user,
		)

	def test_ensure_sku_recipe_copies_planning_designer_fields(self):
		from planning.services import ensure_sku_recipe_for_planning_job, sync_planning_job_fields_to_sku_recipe

		recipe = ensure_sku_recipe_for_planning_job(self.planning_job, actor=self.user)
		sync_planning_job_fields_to_sku_recipe(self.planning_job, recipe)

		self.assertEqual(recipe.sku, 'SKU-SYNC-001')
		self.assertEqual(recipe.color_spec, '4')
		self.assertEqual(recipe.size_w_mm, 100)
		self.assertEqual(recipe.ups, 12)
		self.assertEqual(recipe.plate_set_no, 'SET-99')
