from datetime import date, timedelta
from decimal import Decimal

from django.test import TestCase
from django.utils import timezone

from core.models import JobCard, Machine, Material, Production
from supply_chain.jc_sync import sync_issuance_for_job_card, sync_issuance_from_production
from supply_chain.kpis import (
    assign_abc_classifications,
    build_kpi_dashboard_data,
    classify_fsn,
    compute_reorder_point,
)
from supply_chain.models import RawMaterialSku, StockDemand, StockTransaction
from supply_chain.services import build_dashboard_data
from supply_chain.reports import build_item_wise_monthly_consumption, build_month_wise_item_consumption


def _create_raw_sku(material, sku, purchase_sheet_size='25x36', **kwargs):
    defaults = {
        'uom': 'Packet',
        'safety_stock': 50,
        'max_stock_level': 500,
        'unit_cost': 10,
        'lead_time_days': 1,
    }
    defaults.update(kwargs)
    return RawMaterialSku.objects.create(
        material=material,
        sku=sku,
        purchase_sheet_size=purchase_sheet_size,
        **defaults,
    )


class SupplyChainDashboardTests(TestCase):
    def setUp(self):
        material = Material.objects.create(name='Art Card 210')
        self.item = _create_raw_sku(material, 'ITM-0001')

    def test_closing_stock_formula(self):
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='OPENING',
            sheet_qty_pcs=100,
        )
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='RECEIVING',
            sheet_qty_pcs=50,
        )
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='ISSUANCE',
            sheet_qty_pcs=30,
        )
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='ADJUSTMENT',
            sheet_qty_pcs=10,
        )
        StockDemand.objects.create(raw_material_sku=self.item, month_str='June 2026', sheet_qty_pcs=300)

        row = build_dashboard_data([self.item])[0]

        self.assertEqual(row['opening'], 100)
        self.assertEqual(row['receiving'], 50)
        self.assertEqual(row['issuance'], 30)
        self.assertEqual(row['adjustment'], 10)
        self.assertEqual(row['closing'], 130)
        self.assertEqual(row['monthly_demand'], 300)
        self.assertFalse(row['stockout'])
        self.assertFalse(row['overstock'])

    def test_stockout_alert(self):
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='ISSUANCE',
            sheet_qty_pcs=10,
        )
        row = build_dashboard_data([self.item])[0]
        self.assertTrue(row['stockout'])


class ConsumptionReportTests(TestCase):
    def setUp(self):
        material = Material.objects.create(name='Offset Paper 75')
        self.item = _create_raw_sku(material, 'ITM-0012', uom='Rim', unit_cost=10)

    def test_item_wise_monthly_consumption(self):
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='ISSUANCE',
            month_str='June 2026',
            sheet_qty_pcs=100,
            pkt_rim_qty=2,
        )
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='ISSUANCE',
            month_str='June 2026',
            sheet_qty_pcs=50,
            pkt_rim_qty=1,
        )
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='ISSUANCE',
            month_str='July 2026',
            sheet_qty_pcs=30,
        )

        rows = build_item_wise_monthly_consumption()
        self.assertEqual(len(rows), 2)
        june = next(row for row in rows if row['month'] == 'June 2026')
        self.assertEqual(june['sheet_qty_pcs'], 150)
        self.assertEqual(june['pkt_rim_qty'], 3)
        self.assertEqual(june['consumption_value'], 1500)

    def test_month_wise_sort_order(self):
        material_b = Material.objects.create(name='Art Card 210')
        item_b = _create_raw_sku(material_b, 'ITM-0002', unit_cost=5)
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='ISSUANCE',
            month_str='June 2026',
            sheet_qty_pcs=10,
        )
        StockTransaction.objects.create(
            raw_material_sku=item_b,
            transaction_type='ISSUANCE',
            month_str='June 2026',
            sheet_qty_pcs=20,
        )

        rows = build_month_wise_item_consumption()
        self.assertEqual(rows[0]['month'], 'June 2026')
        self.assertEqual(rows[0]['item_id'], 'ITM-0002')
        self.assertEqual(rows[1]['item_id'], 'ITM-0012')


class KpiEngineTests(TestCase):
    def setUp(self):
        material = Material.objects.create(name='Art Card 210')
        self.item = _create_raw_sku(
            material,
            'ITM-0001',
            unit_cost=10,
            safety_stock=50,
            max_stock_level=500,
            lead_time_days=5,
        )

    def test_abc_classification(self):
        result = assign_abc_classifications({1: 700, 2: 200, 3: 100})
        self.assertEqual(result[1], 'A')
        self.assertEqual(result[2], 'B')
        self.assertEqual(result[3], 'C')

    def test_fsn_classification(self):
        self.assertEqual(classify_fsn(10), 'Fast')
        self.assertEqual(classify_fsn(45), 'Slow')
        self.assertEqual(classify_fsn(120), 'Non-Moving')
        self.assertEqual(classify_fsn(None), 'Non-Moving')

    def test_reorder_point_formula(self):
        self.assertEqual(compute_reorder_point(10, 5, 50), 100)

    def test_kpi_alerts_and_metrics(self):
        today = timezone.now().date()
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='OPENING',
            sheet_qty_pcs=200,
        )
        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='ISSUANCE',
            sheet_qty_pcs=40,
            date=today - timedelta(days=45),
        )
        StockDemand.objects.create(raw_material_sku=self.item, month_str='June 2026', sheet_qty_pcs=300)

        rows, summary = build_kpi_dashboard_data([self.item])
        row = rows[0]

        self.assertEqual(row['closing'], 160)
        self.assertEqual(row['abc_class'], 'A')
        self.assertEqual(row['fsn_class'], 'Slow')
        self.assertEqual(row['reorder_point'], 100)
        self.assertEqual(row['inventory_value'], 1600)
        self.assertFalse(row['stockout'])
        self.assertFalse(row['reorder'])
        self.assertFalse(row['safety_stock_alert'])
        self.assertTrue(row['slow_moving'])
        self.assertEqual(summary['slow_moving'], 1)


class JobCardIssuanceSyncTests(TestCase):
    def setUp(self):
        material = Material.objects.create(name='Offset Paper 75')
        self.item = _create_raw_sku(material, 'ITM-0012', uom='Rim')
        machine = Machine.objects.create(name='KBA 1')
        self.job_card = JobCard.objects.create(
            job_card_no='JC-1001',
            SKU='SKU-1',
            material=material,
            purchase_sheet_size='25x36',
            order_qty=1000,
            total_impressions_required=1000,
            total_sheet_quantity=500,
            total_colors=4,
            plate_set_no='PS-1',
            po_date=date(2026, 6, 1),
            status='released',
            machine_name=machine,
        )

    def test_production_auto_syncs_issuance(self):
        production = Production.objects.create(
            job_card=self.job_card,
            date=date(2026, 6, 15),
            shift='A',
            machine=self.job_card.machine_name,
            output_sheets=120,
            waste_sheets=10,
            impressions=120,
            planned_time=60,
            run_time=55,
        )

        txn = sync_issuance_from_production(production)
        self.assertIsNotNone(txn)
        self.assertEqual(txn.source, 'JOB_CARD')
        self.assertEqual(txn.gin_jc, 'JC-1001')
        self.assertEqual(txn.sheet_qty_pcs, 130)
        self.assertEqual(txn.job_card_id, self.job_card.id)
        self.assertEqual(txn.raw_material_sku_id, self.item.id)

        production.output_sheets = 150
        production.save()
        txn.refresh_from_db()
        self.assertEqual(txn.sheet_qty_pcs, 160)

    def test_sync_job_card_bulk(self):
        Production.objects.create(
            job_card=self.job_card,
            date=date(2026, 6, 10),
            shift='A',
            machine=self.job_card.machine_name,
            output_sheets=50,
            waste_sheets=0,
            impressions=50,
            planned_time=30,
            run_time=28,
        )
        Production.objects.create(
            job_card=self.job_card,
            date=date(2026, 6, 11),
            shift='B',
            machine=self.job_card.machine_name,
            output_sheets=40,
            waste_sheets=5,
            impressions=40,
            planned_time=30,
            run_time=27,
        )

        synced, skipped = sync_issuance_for_job_card(self.job_card)
        self.assertEqual(synced, 2)
        self.assertEqual(skipped, 0)
        self.assertEqual(self.job_card.stock_transactions.filter(source='JOB_CARD').count(), 2)


class PhysicalStockCountTests(TestCase):
    def setUp(self):
        material = Material.objects.create(name='Art Card 210')
        self.item = _create_raw_sku(material, 'ITM-0002')

    def test_inventory_accuracy_formula(self):
        from supply_chain.physical_count import compute_inventory_accuracy, save_physical_count

        self.assertEqual(compute_inventory_accuracy(95, 100), Decimal('95.00'))
        self.assertEqual(compute_inventory_accuracy(0, 0), Decimal('100.00'))
        self.assertEqual(compute_inventory_accuracy(10, 0), Decimal('0.00'))

        StockTransaction.objects.create(
            raw_material_sku=self.item,
            transaction_type='OPENING',
            sheet_qty_pcs=200,
        )
        count = save_physical_count(self.item, date(2026, 6, 20), physical_sheet_qty=180)
        self.assertEqual(count.system_sheet_qty, 200)
        self.assertEqual(count.accuracy_percent, Decimal('90.00'))
        self.assertEqual(count.variance, -20)


class TransactionImportDedupeTests(TestCase):
    def setUp(self):
        material = Material.objects.create(name='Dup Paper')
        self.item = _create_raw_sku(material, 'DUP-001', purchase_sheet_size='20*30')

    def _workbook(self, rows):
        from io import BytesIO
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([
            'Month', 'Date', 'GIN / JC', 'Raw Material SKU', 'Material Name',
            'Purchase Sheet Size', 'UOM', 'Sheet Packing/Pcs', 'Sheet Qty/Pcs', 'Pkt/Rim Qty',
        ])
        for row in rows:
            ws.append(row)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        buf.name = 'receiving.xlsx'
        return buf

    def test_blocks_duplicate_receiving_rows(self):
        from supply_chain.excel_io import import_transactions

        row = [
            'July 2026', '2026-07-11', 'GIN-9', self.item.sku, 'Dup Paper',
            '20*30', 'Sheets', 1, 50, 0,
        ]
        first = import_transactions('RECEIVING', self._workbook([row, row]))
        second = import_transactions('RECEIVING', self._workbook([row]))

        self.assertEqual(first['created'], 1)
        self.assertEqual(first['skipped_duplicates'], 1)
        self.assertEqual(second['created'], 0)
        self.assertEqual(second['skipped_duplicates'], 1)
        self.assertEqual(
            StockTransaction.objects.filter(
                raw_material_sku=self.item,
                transaction_type='RECEIVING',
                source='MANUAL',
            ).count(),
            1,
        )


class ChangeManagementTests(TestCase):
    def setUp(self):
        from django.contrib.auth.models import User
        self.admin = User.objects.create_superuser(username='admin', password='password')
        self.user = User.objects.create_user(username='user', password='password')
        
        # Setup roles using UserProfile
        from core.models import UserProfile
        admin_profile = self.admin.profile
        admin_profile.role = 'admin'
        admin_profile.save()

        user_profile = self.user.profile
        user_profile.role = 'supply_chain'
        user_profile.save()
        
        self.material = Material.objects.create(name='Taffeta')
        self.item = _create_raw_sku(self.material, 'RM-TAF-2536', purchase_sheet_size='25x36')
        
    def test_admin_direct_create(self):
        # Admin can create a demand directly, and it logs an approved ChangeRequest
        from supply_chain.models import ChangeRequest
        self.client.force_login(self.admin)
        response = self.client.post('/supply-chain/monthly-demand/', {
            'raw_material_sku': self.item.pk,
            'month_str': 'July 2026',
            'sheet_qty_pcs': 200,
            'pkt_rim_qty': 0
        })
        self.assertEqual(response.status_code, 302)
        self.assertEqual(StockDemand.objects.count(), 1)
        self.assertEqual(ChangeRequest.objects.count(), 1)
        req = ChangeRequest.objects.first()
        self.assertEqual(req.status, 'APPROVED')
        self.assertEqual(req.action, 'CREATE')
        
    def test_user_request_create(self):
        # User cannot create directly, creates a pending ChangeRequest instead
        from supply_chain.models import ChangeRequest
        self.client.force_login(self.user)
        response = self.client.post('/supply-chain/monthly-demand/', {
            'raw_material_sku': self.item.pk,
            'month_str': 'July 2026',
            'sheet_qty_pcs': 200,
            'pkt_rim_qty': 0
        })
        self.assertEqual(response.status_code, 302)
        # Check that no demand is created in database directly
        self.assertEqual(StockDemand.objects.count(), 0)
        # Check that a pending change request is created
        self.assertEqual(ChangeRequest.objects.count(), 1)
        req = ChangeRequest.objects.first()
        self.assertEqual(req.status, 'PENDING')
        self.assertEqual(req.action, 'CREATE')
        
    def test_admin_approve_create_request(self):
        from supply_chain.models import ChangeRequest
        req = ChangeRequest.objects.create(
            model_name='StockDemand',
            action='CREATE',
            proposed_data={
                'raw_material_sku_id': self.item.pk,
                'month_str': 'August 2026',
                'sheet_qty_pcs': 500,
                'pkt_rim_qty': 1
            },
            requested_by=self.user,
            status='PENDING'
        )
        self.client.force_login(self.admin)
        response = self.client.post(f'/supply-chain/change-requests/{req.pk}/approve/')
        self.assertEqual(response.status_code, 302)
        req.refresh_from_db()
        self.assertEqual(req.status, 'APPROVED')
        self.assertEqual(StockDemand.objects.count(), 1)
        demand = StockDemand.objects.first()
        self.assertEqual(demand.month_str, 'August 2026')
        self.assertEqual(demand.sheet_qty_pcs, 500)

    def test_admin_bulk_delete(self):
        from supply_chain.models import ChangeRequest
        demand1 = StockDemand.objects.create(raw_material_sku=self.item, month_str='June 2026', sheet_qty_pcs=100)
        demand2 = StockDemand.objects.create(raw_material_sku=self.item, month_str='July 2026', sheet_qty_pcs=200)

        self.client.force_login(self.admin)
        response = self.client.post('/supply-chain/bulk-delete/', {
            'model_name': 'StockDemand',
            'selected_ids': [demand1.pk, demand2.pk],
            'redirect_url': '/supply-chain/monthly-demand/'
        })
        self.assertEqual(response.status_code, 302)
        # Verify both are soft-deleted (is_active = False)
        demand1.refresh_from_db()
        demand2.refresh_from_db()
        self.assertFalse(demand1.is_active)
        self.assertFalse(demand2.is_active)
        
        # Verify approved ChangeRequests are logged
        self.assertEqual(ChangeRequest.objects.filter(status='APPROVED', action='DELETE').count(), 2)

    def test_user_bulk_delete_request(self):
        from supply_chain.models import ChangeRequest
        demand1 = StockDemand.objects.create(raw_material_sku=self.item, month_str='June 2026', sheet_qty_pcs=100)
        demand2 = StockDemand.objects.create(raw_material_sku=self.item, month_str='July 2026', sheet_qty_pcs=200)

        self.client.force_login(self.user)
        response = self.client.post('/supply-chain/bulk-delete/', {
            'model_name': 'StockDemand',
            'selected_ids': [demand1.pk, demand2.pk],
            'redirect_url': '/supply-chain/monthly-demand/'
        })
        self.assertEqual(response.status_code, 302)
        # Verify they are NOT soft-deleted
        demand1.refresh_from_db()
        demand2.refresh_from_db()
        self.assertTrue(demand1.is_active)
        self.assertTrue(demand2.is_active)
        
        # Verify pending ChangeRequests are logged
        self.assertEqual(ChangeRequest.objects.filter(status='PENDING', action='DELETE').count(), 2)


class IssuanceQueueAndApprovalTests(TestCase):
    def setUp(self):
        from django.contrib.auth.models import User
        self.admin = User.objects.create_superuser(username='admin', password='password')
        
        # Setup roles using UserProfile
        from core.models import UserProfile
        admin_profile = self.admin.profile
        admin_profile.role = 'admin'
        admin_profile.save()

        material = Material.objects.create(name='Offset Paper 75')
        self.item = _create_raw_sku(material, 'ITM-0012', uom='Rim')
        machine = Machine.objects.create(name='KBA 1')
        self.job_card = JobCard.objects.create(
            job_card_no='JC-1001',
            SKU='SKU-1',
            material=material,
            purchase_sheet_size='25x36',
            order_qty=1000,
            total_impressions_required=1000,
            total_sheet_quantity=500,
            total_colors=4,
            plate_set_no='PS-1',
            po_date=date(2026, 6, 1),
            status='released',
            machine_name=machine,
        )

    def test_production_sync_creates_pending_transaction(self):
        production = Production.objects.create(
            job_card=self.job_card,
            date=date(2026, 6, 15),
            shift='A',
            machine=self.job_card.machine_name,
            output_sheets=120,
            waste_sheets=10,
            impressions=120,
            planned_time=60,
            run_time=55,
        )

        txn = sync_issuance_from_production(production)
        self.assertIsNotNone(txn)
        self.assertEqual(txn.transaction_type, 'ISSUANCE')
        self.assertEqual(txn.source, 'JOB_CARD')
        self.assertFalse(txn.is_approved)  # Should default to False

        # Verify it is not factored into stock levels
        row = build_dashboard_data([self.item])[0]
        self.assertEqual(row['issuance'], 0)
        self.assertEqual(row['closing'], 0)

        # Approve single transaction
        self.client.force_login(self.admin)
        response = self.client.get(f'/supply-chain/issuance/{txn.pk}/approve/')
        self.assertEqual(response.status_code, 302)

        txn.refresh_from_db()
        self.assertTrue(txn.is_approved)

        # Verify it is now factored into stock levels
        row = build_dashboard_data([self.item])[0]
        self.assertEqual(row['issuance'], 130)
        self.assertEqual(row['closing'], -130)

    def test_bulk_approval(self):
        production1 = Production.objects.create(
            job_card=self.job_card,
            date=date(2026, 6, 15),
            shift='A',
            machine=self.job_card.machine_name,
            output_sheets=50,
            waste_sheets=0,
            impressions=50,
            planned_time=60,
            run_time=55,
        )
        production2 = Production.objects.create(
            job_card=self.job_card,
            date=date(2026, 6, 16),
            shift='B',
            machine=self.job_card.machine_name,
            output_sheets=100,
            waste_sheets=10,
            impressions=100,
            planned_time=60,
            run_time=55,
        )

        txn1 = sync_issuance_from_production(production1)
        txn2 = sync_issuance_from_production(production2)

        self.assertFalse(txn1.is_approved)
        self.assertFalse(txn2.is_approved)

        self.client.force_login(self.admin)
        response = self.client.post('/supply-chain/issuance/bulk-approve/', {
            'selected_ids': [txn1.pk, txn2.pk]
        })
        self.assertEqual(response.status_code, 302)

        txn1.refresh_from_db()
        txn2.refresh_from_db()
        self.assertTrue(txn1.is_approved)
        self.assertTrue(txn2.is_approved)

        row = build_dashboard_data([self.item])[0]
        self.assertEqual(row['issuance'], 160)


