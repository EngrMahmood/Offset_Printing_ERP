from datetime import date

from django.test import TestCase

from core.models import Dispatch, JobCard, Machine, Material, Production, Sorter
from planning.models import PlanningDispatchRun, PlanningJob, SkuRecipe
from supply_chain.demand_gap import (
    build_demand_gap_report,
    compute_job_demand_sheets,
    is_completed_planning_job,
)
from supply_chain.models import RawMaterialSku, StockTransaction


class DemandGapLogicTests(TestCase):
    def _create_print_job(self, jc_number='JC-GAP-1', status='draft', order_qty=10000, ups=10, purchase_sheet_ups=2):
        material = Material.objects.get_or_create(name='Art Card 210')[0]
        SkuRecipe.objects.create(
            sku=f'SKU-{jc_number}',
            job_name='Gap Test Job',
            material='Art Card 210',
            ups=ups,
            purchase_sheet_ups=purchase_sheet_ups,
            print_sheet_size='25x36',
            purchase_sheet_size='25x36',
            master_data_status='approved',
        )
        return PlanningJob.objects.create(
            jc_number=jc_number,
            sku=f'SKU-{jc_number}',
            material='Art Card 210',
            order_qty=order_qty,
            ups=ups,
            purchase_sheet_ups=purchase_sheet_ups,
            wastage_sheets=0,
            status=status,
        )

    def test_full_demand_for_draft_job_without_production(self):
        job = self._create_print_job()
        demand = compute_job_demand_sheets(job)
        self.assertEqual(job.purchase_sheet_required_display, 500)
        self.assertEqual(demand['job_demand_sheets'], 500)

    def test_excludes_completed_jobs_from_report(self):
        self._create_print_job(status='draft')
        self._create_print_job(jc_number='JC-GAP-DONE', status='completed')
        report = build_demand_gap_report()
        jc_numbers = [row['jc_number'] for row in report['job_rows']]
        self.assertIn('JC-GAP-1', jc_numbers)
        self.assertNotIn('JC-GAP-DONE', jc_numbers)
        self.assertFalse(is_completed_planning_job(self._create_print_job(jc_number='JC-X', status='draft')))

    def test_print_job_uses_print_pack_and_dispatch_signals(self):
        job = self._create_print_job(order_qty=10000)
        machine = Machine.objects.create(name='Press 1')
        job_card = JobCard.objects.create(
            job_card_no=job.jc_number,
            planning_job=job,
            SKU=job.sku,
            material=Material.objects.get(name='Art Card 210'),
            order_qty=10000,
            ups=10,
            purchase_sheet_ups=2,
            total_sheet_quantity=5000,
            total_impressions_required=5000,
            total_colors=4,
            plate_set_no='PS-1',
            po_date=date(2026, 6, 1),
            status='in_production',
            machine_name=machine,
            is_print_job=True,
        )
        Production.objects.create(
            job_card=job_card,
            entry_type='printing',
            date=date(2026, 6, 10),
            shift='A',
            machine=machine,
            output_sheets=300,
            waste_sheets=0,
            impressions=300,
            planned_time=60,
            run_time=55,
        )
        sorter = Sorter.objects.create(name='Sorter 1')
        Production.objects.create(
            job_card=job_card,
            entry_type='packing',
            date=date(2026, 6, 11),
            shift='A',
            packing_qty=3000,
            sorting_waste_qty=0,
            sorter=sorter,
        )
        Dispatch.objects.create(
            job_card=job_card,
            dc_no='DC-001',
            dispatch_date=date(2026, 6, 12),
            dispatch_qty=3000,
        )

        demand = compute_job_demand_sheets(job, job_card)
        self.assertEqual(demand['remaining_from_print'], 350)
        self.assertEqual(demand['remaining_from_pack'], 350)
        self.assertEqual(demand['remaining_from_dispatch'], 350)
        self.assertEqual(demand['job_demand_sheets'], 350)

    def test_cut_and_pack_uses_pack_and_dispatch_only(self):
        material = Material.objects.create(name='A4 Rim Paper')
        SkuRecipe.objects.create(
            sku='SKU-CUT-1',
            job_name='Cut Pack Job',
            material='A4 Rim Paper',
            job_process_type='cut_and_pack',
            ups=1,
            purchase_sheet_ups=1,
            purchase_sheet_size='A4',
            master_data_status='approved',
        )
        job = PlanningJob.objects.create(
            jc_number='JC-CUT-1',
            sku='SKU-CUT-1',
            material='A4 Rim Paper',
            order_qty=5000,
            ups=1,
            purchase_sheet_ups=100,
            purchase_sheet_required=50,
            wastage_sheets=0,
            status='released',
        )
        machine = Machine.objects.create(name='Cutter 1')
        job_card = JobCard.objects.create(
            job_card_no=job.jc_number,
            planning_job=job,
            SKU=job.sku,
            material=material,
            order_qty=5000,
            total_sheet_quantity=50,
            total_impressions_required=0,
            total_colors=0,
            plate_set_no='N/A',
            po_date=date(2026, 6, 1),
            status='in_production',
            machine_name=machine,
            is_print_job=False,
        )
        sorter = Sorter.objects.create(name='Sorter Cut')
        Production.objects.create(
            job_card=job_card,
            entry_type='packing',
            date=date(2026, 6, 10),
            shift='A',
            packing_qty=2000,
            sorting_waste_qty=0,
            sorter=sorter,
        )

        demand = compute_job_demand_sheets(job, job_card)
        self.assertTrue(demand['is_cut_and_pack'])
        self.assertEqual(demand['consumed_print_sheets'], 0)
        self.assertEqual(demand['job_demand_sheets'], 30)

    def test_material_gap_rollup(self):
        job = self._create_print_job()
        material = Material.objects.get(name='Art Card 210')
        sku = RawMaterialSku.objects.create(
            material=material,
            sku='ITM-GAP-1',
            purchase_sheet_size='25x36',
        )
        StockTransaction.objects.create(
            raw_material_sku=sku,
            transaction_type='OPENING',
            sheet_qty_pcs=100,
        )

        report = build_demand_gap_report()
        material_row = next(
            row for row in report['material_rows']
            if row.get('raw_material_sku') and row['raw_material_sku'].pk == sku.pk
        )
        self.assertEqual(material_row['total_demand'], 500)
        self.assertEqual(material_row['on_hand'], 100)
        self.assertEqual(material_row['gap'], 400)
        self.assertEqual(material_row['gap_status'], 'shortage')
        self.assertEqual(material_row['purchase_sheet_size'], '25*36')

    def test_purchase_size_variants_merge(self):
        from supply_chain.models import normalize_purchase_sheet_size

        self.assertEqual(normalize_purchase_sheet_size('10.5 x 15'), '10.5*15')
        self.assertEqual(normalize_purchase_sheet_size('10.5*15'), '10.5*15')
        self.assertEqual(normalize_purchase_sheet_size('10.5X15'), '10.5*15')
        self.assertEqual(normalize_purchase_sheet_size('A4'), 'A4')

        material = Material.objects.create(name='Taffeta')
        SkuRecipe.objects.create(
            sku='SKU-TAF-A',
            job_name='Taf A',
            material='Taffeta',
            ups=10,
            purchase_sheet_ups=2,
            purchase_sheet_size='10.5 x 15',
            master_data_status='approved',
        )
        SkuRecipe.objects.create(
            sku='SKU-TAF-B',
            job_name='Taf B',
            material='Taffeta',
            ups=10,
            purchase_sheet_ups=2,
            purchase_sheet_size='10.5*15',
            master_data_status='approved',
        )
        PlanningJob.objects.create(
            jc_number='JC-TAF-A',
            sku='SKU-TAF-A',
            material='Taffeta',
            order_qty=10000,
            ups=10,
            purchase_sheet_ups=2,
            purchase_sheet_size='10.5 x 15',
            wastage_sheets=0,
            status='draft',
        )
        PlanningJob.objects.create(
            jc_number='JC-TAF-B',
            sku='SKU-TAF-B',
            material='Taffeta',
            order_qty=20000,
            ups=10,
            purchase_sheet_ups=2,
            purchase_sheet_size='10.5*15',
            wastage_sheets=0,
            status='draft',
        )

        report = build_demand_gap_report()
        taffeta_rows = [
            row for row in report['material_rows']
            if row.get('material_name') == 'Taffeta' and row.get('purchase_sheet_size') == '10.5*15'
        ]
        self.assertEqual(len(taffeta_rows), 1)
        self.assertEqual(taffeta_rows[0]['job_count'], 2)
        self.assertEqual(taffeta_rows[0]['total_demand'], 1500)

    def test_material_name_case_variants_merge(self):
        from supply_chain.models import normalize_material_name

        self.assertEqual(normalize_material_name('Art Paper 128'), 'art paper 128')
        self.assertEqual(normalize_material_name('Art paper 128'), 'art paper 128')
        self.assertEqual(normalize_material_name('  BROWN   STICKER '), 'brown sticker')

        material = Material.objects.create(name='Art Paper 128')
        sku = RawMaterialSku.objects.create(
            material=material,
            sku='ARTPAPER128-2536',
            purchase_sheet_size='25*36',
        )
        PlanningJob.objects.create(
            jc_number='JC-ART-A',
            sku='SKU-ART-A',
            material='Art Paper 128',
            order_qty=10000,
            ups=10,
            purchase_sheet_ups=2,
            purchase_sheet_size='25x36',
            wastage_sheets=0,
            status='draft',
        )
        PlanningJob.objects.create(
            jc_number='JC-ART-B',
            sku='SKU-ART-B',
            material='Art paper 128',
            order_qty=4000,
            ups=10,
            purchase_sheet_ups=2,
            purchase_sheet_size='25*36',
            wastage_sheets=0,
            status='draft',
        )

        report = build_demand_gap_report()
        art_rows = [
            row for row in report['material_rows']
            if normalize_material_name(row.get('material_name')) == 'art paper 128'
            and row.get('purchase_sheet_size') == '25*36'
        ]
        self.assertEqual(len(art_rows), 1)
        self.assertEqual(art_rows[0]['job_count'], 2)
        self.assertEqual(art_rows[0]['total_demand'], 700)
        self.assertEqual(art_rows[0]['raw_material_sku'].pk, sku.pk)
        self.assertEqual(art_rows[0]['gap_status'], 'shortage')

    def test_dispatch_from_planning_run_when_no_job_card(self):
        job = self._create_print_job(jc_number='JC-NO-JC')
        PlanningDispatchRun.objects.create(
            planning_job=job,
            dispatch_index=1,
            delivered_qty=5000,
        )
        demand = compute_job_demand_sheets(job)
        self.assertEqual(demand['dispatched_pcs'], 5000)
        self.assertEqual(demand['job_demand_sheets'], 250)
