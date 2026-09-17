from decimal import Decimal

from django.contrib.auth.models import User
from django.test import TestCase

from planning.models import PlanningJob, SkuRecipe

from .models import (
    BomTemplate, BomTemplateLine, ItemCategory, ProcessStep, QuantityBasis, RawItem,
    SkuBom, SkuBomLine, UnitOfMeasure,
)
from .requirements import explode_requirements, raise_item_requests_from_shortfalls


class RequirementsExplosionTests(TestCase):
    def setUp(self):
        self.kg = UnitOfMeasure.objects.create(code='KG', name='Kilogram', decimal_places=3)
        self.cat = ItemCategory.objects.create(name='Ink', code='INK', default_uom=self.kg)
        self.step = ProcessStep.objects.create(code='PRINTING', label='Printing', sequence=10)
        self.basis = QuantityBasis.objects.create(code='per_1000_pcs', label='Per 1000 Pieces', formula_key='order_qty')
        self.ink = RawItem.objects.create(item_code='INK-001', name='Black Ink', category=self.cat, uom=self.kg, unit_cost=Decimal('500'), safety_stock=Decimal('2'))
        self.template = BomTemplate.objects.create(code='T1', name='T1', status='approved', priority=1)
        self.template_line = BomTemplateLine.objects.create(
            template=self.template, line_code='ink', sequence=1, process_step=self.step,
            bom_item_type=self.cat, resolution_mode='fixed', raw_item=self.ink,
            quantity_basis=self.basis, entry_mode='formula', qty_expression='1',
        )
        self.recipe = SkuRecipe.objects.create(sku='REQ-001', material='Art Card')
        self.bom = SkuBom.objects.create(sku_recipe=self.recipe, version=1, is_current=True, status='approved', source_template=self.template)
        SkuBomLine.objects.create(
            bom=self.bom, line_code='ink', process_step=self.step, raw_item=self.ink, uom_code='KG',
            qty_per_unit=Decimal('1'), gross_qty=Decimal('1'), unit_cost_snapshot=Decimal('500'),
            line_cost=Decimal('500'), source_template_line=self.template_line,
        )

    def test_aggregates_two_jobs_sharing_a_raw_item(self):
        PlanningJob.objects.create(jc_number='JC-1', sku='REQ-001', order_qty=1000)
        PlanningJob.objects.create(jc_number='JC-2', sku='REQ-001', order_qty=2000)
        rows = explode_requirements(PlanningJob.objects.all())
        self.assertEqual(len(rows), 1)
        row = rows[0]
        # per_1000_pcs basis: gross_qty(1) * (order_qty/1000) summed over jobs = 1*1 + 1*2 = 3
        self.assertEqual(row['required_qty'], Decimal('3'))
        self.assertEqual(row['job_count'], 2)

    def test_shortfall_when_on_hand_unknown_and_no_bridge(self):
        PlanningJob.objects.create(jc_number='JC-3', sku='REQ-001', order_qty=1000)
        rows = explode_requirements(PlanningJob.objects.all())
        row = rows[0]
        self.assertIsNone(row['on_hand'])  # not bridged to a RawMaterialSku
        self.assertIsNone(row['shortfall'])

    def test_job_without_recipe_is_skipped_not_erroring(self):
        PlanningJob.objects.create(jc_number='JC-4', sku='NO-SUCH-SKU', order_qty=500)
        rows = explode_requirements(PlanningJob.objects.all())
        self.assertEqual(rows, [])

    def test_job_without_approved_bom_is_skipped(self):
        self.bom.status = 'draft'
        self.bom.save()
        PlanningJob.objects.create(jc_number='JC-5', sku='REQ-001', order_qty=500)
        rows = explode_requirements(PlanningJob.objects.all())
        self.assertEqual(rows, [])


class RaiseItemRequestTests(TestCase):
    def setUp(self):
        from supply_chain.models import ItemRequestDepartment, ItemRequestType

        self.kg = UnitOfMeasure.objects.create(code='KG', name='Kilogram', decimal_places=3)
        self.cat = ItemCategory.objects.create(name='Ink', code='INK', default_uom=self.kg)
        self.ink = RawItem.objects.create(item_code='INK-002', name='Cyan Ink', category=self.cat, uom=self.kg, unit_cost=Decimal('600'))
        self.user = User.objects.create_user('buyer1', password='x')
        self.request_type, _ = ItemRequestType.objects.get_or_create(code='RM', defaults={'name': 'Raw Material (Test)'})
        self.department, _ = ItemRequestDepartment.objects.get_or_create(name='Printing')

    def test_raises_one_item_request_per_row(self):
        rows = [{
            'raw_item': self.ink, 'required_qty': Decimal('10'), 'on_hand': Decimal('0'),
            'safety_stock': Decimal('0'), 'shortfall': Decimal('10'), 'unit_cost': Decimal('600'),
            'extended_cost': Decimal('6000'), 'lead_time_days': 5, 'suggested_order_qty': Decimal('10'),
            'job_count': 1,
        }]
        created = raise_item_requests_from_shortfalls(rows, self.user, self.request_type, self.department)
        self.assertEqual(len(created), 1)
        ir = created[0]
        self.assertTrue(ir.request_no.startswith('IR-RM-'))
        self.assertEqual(ir.required_quantity, Decimal('10'))
        self.assertEqual(ir.raised_by, self.user)

    def test_zero_quantity_row_is_skipped(self):
        rows = [{
            'raw_item': self.ink, 'required_qty': Decimal('0'), 'on_hand': Decimal('10'),
            'safety_stock': Decimal('0'), 'shortfall': Decimal('0'), 'unit_cost': Decimal('600'),
            'extended_cost': Decimal('0'), 'lead_time_days': 5, 'suggested_order_qty': Decimal('0'),
            'job_count': 1,
        }]
        created = raise_item_requests_from_shortfalls(rows, self.user, self.request_type, self.department)
        self.assertEqual(created, [])
