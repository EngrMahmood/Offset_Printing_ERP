from decimal import Decimal

from django.contrib.auth.models import User
from django.test import TestCase

from planning.models import SkuRecipe

from .models import BomTemplate, ItemCategory, ProcessStep, QuantityBasis, RawItem, SkuBom, UnitOfMeasure
from .services import approve_bom, reject_bom, review_bom, soft_delete_bom, submit_bom
from .spec import available_sku_fields, resolve_sku_spec


class SpecResolutionTests(TestCase):
    def test_available_sku_fields_matches_real_model(self):
        fields = available_sku_fields()
        self.assertIn('material', fields)
        self.assertIn('size_w_mm', fields)
        self.assertNotIn('created_by', fields)  # not on the allow-list

    def test_gsm_parsed_from_material_name(self):
        recipe = SkuRecipe.objects.create(sku='SPEC-001', material='Art Card 300gsm')
        spec = resolve_sku_spec(recipe)
        self.assertEqual(spec['gsm'], '300')

    def test_gsm_none_when_unparseable(self):
        recipe = SkuRecipe.objects.create(sku='SPEC-002', material='Some Random Material')
        spec = resolve_sku_spec(recipe)
        self.assertIsNone(spec['gsm'])

    def test_total_colors_sums_plus_form(self):
        recipe = SkuRecipe.objects.create(sku='SPEC-003', color_spec='1+1')
        spec = resolve_sku_spec(recipe)
        self.assertEqual(spec['total_colors'], '2')

    def test_piece_area_sqm(self):
        recipe = SkuRecipe.objects.create(sku='SPEC-004', size_w_mm=1000, size_h_mm=500)
        spec = resolve_sku_spec(recipe)
        self.assertEqual(Decimal(spec['piece_area_sqm']), Decimal('0.5'))


class BomWorkflowTests(TestCase):
    def setUp(self):
        self.user = User.objects.create_user('approver1', password='x')
        self.recipe = SkuRecipe.objects.create(sku='WF-001', material='Art Card')
        self.bom = SkuBom.objects.create(sku_recipe=self.recipe, version=1, is_current=True, status='draft')

    def test_submit_review_approve_flow(self):
        submit_bom(self.bom, self.user)
        self.assertEqual(self.bom.status, 'pending_review')
        review_bom(self.bom, self.user)
        self.assertEqual(self.bom.status, 'reviewed')
        self.assertEqual(self.bom.reviewed_by, self.user)
        approve_bom(self.bom, self.user)
        self.assertEqual(self.bom.status, 'approved')
        self.assertEqual(self.bom.approved_by, self.user)

    def test_reject_returns_to_draft_with_comment(self):
        submit_bom(self.bom, self.user)
        reject_bom(self.bom, self.user, comment='Wrong ink quantity')
        self.assertEqual(self.bom.status, 'draft')
        self.assertEqual(self.bom.rejection_comment, 'Wrong ink quantity')

    def test_soft_delete(self):
        soft_delete_bom(self.bom, self.user)
        self.assertFalse(self.bom.is_active)


class BomTemplateCleanTests(TestCase):
    """Regression: BomTemplateForm.is_valid() runs instance.full_clean() on the
    BomTemplate BEFORE it has a pk (and therefore before the condition/line
    formsets have saved anything) — exactly what the Recipe Wizard's "Create
    recipe from this combination" flow does on first save. clean() must not
    try to read self.lines (a reverse FK manager) on an unsaved instance."""

    def test_full_clean_succeeds_on_unsaved_new_template(self):
        template = BomTemplate(code='NEW-TPL', name='New', priority=0)
        template.full_clean()  # must not raise ValueError

    def test_cycle_detection_still_works_once_saved(self):
        from .models import BomTemplateLine, ItemCategory as _Cat, ProcessStep as _Step, QuantityBasis as _QB
        kg = UnitOfMeasure.objects.create(code='KG2', name='Kilogram')
        cat = _Cat.objects.create(name='Ink2', code='INK2', default_uom=kg)
        step = _Step.objects.create(code='STEP2', label='Step', sequence=1)
        basis = _QB.objects.create(code='basis2', label='Basis', formula_key='order_qty')
        item = RawItem.objects.create(item_code='INK-200', name='Ink', category=cat, uom=kg)

        template = BomTemplate.objects.create(code='CYCLE-TPL', name='Cycle', priority=0)
        line_a = BomTemplateLine.objects.create(
            template=template, line_code='a', process_step=step, bom_item_type=cat,
            resolution_mode='fixed', raw_item=item, quantity_basis=basis, entry_mode='simple', fixed_qty=Decimal('1'),
        )
        line_a.driver_line_id = None
        line_b = BomTemplateLine.objects.create(
            template=template, line_code='b', process_step=step, bom_item_type=cat,
            resolution_mode='fixed', raw_item=item, quantity_basis=basis, entry_mode='simple', fixed_qty=Decimal('1'),
            driver_line=line_a,
        )
        line_a.driver_line = line_b
        line_a.save()
        with self.assertRaises(Exception):
            template.full_clean()


class RawItemCostHistoryTests(TestCase):
    def test_record_cost_change_writes_history(self):
        kg = UnitOfMeasure.objects.create(code='KG', name='Kilogram')
        cat = ItemCategory.objects.create(name='Ink', code='INK', default_uom=kg)
        item = RawItem.objects.create(item_code='INK-100', name='Test Ink', category=cat, uom=kg, unit_cost=Decimal('10'))
        user = User.objects.create_user('costadmin', password='x')
        item.record_cost_change(Decimal('15'), user=user, source='manual')
        item.refresh_from_db()
        self.assertEqual(item.unit_cost, Decimal('15'))
        self.assertEqual(item.cost_history.count(), 1)
        self.assertEqual(item.cost_history.first().unit_cost, Decimal('15'))
