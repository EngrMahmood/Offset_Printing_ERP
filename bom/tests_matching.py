from decimal import Decimal

from django.contrib.auth.models import User
from django.test import TestCase

from planning.models import SkuRecipe

from .matching import best_template, evaluate_condition, find_matching_templates
from .models import (
    BomTemplate, BomTemplateCondition, BomTemplateLine, ItemCategory, ProcessStep,
    QuantityBasis, RawItem, SpecAttribute, UnitOfMeasure,
)
from .services import generate_bom_for_sku


def _seed_masters():
    kg = UnitOfMeasure.objects.create(code='KG', name='Kilogram', decimal_places=3)
    sheet = UnitOfMeasure.objects.create(code='SHEET', name='Sheet', decimal_places=0)
    paper_cat = ItemCategory.objects.create(name='Paper/Board', code='PAP', default_uom=sheet)
    ink_cat = ItemCategory.objects.create(name='Ink', code='INK', default_uom=kg)
    step = ProcessStep.objects.create(code='PRINTING', label='Printing', sequence=10)
    basis_1000pcs = QuantityBasis.objects.create(code='per_1000_pcs', label='Per 1000 Pieces', formula_key='order_qty')
    basis_driver = QuantityBasis.objects.create(code='per_unit_of_line', label='Per Unit of Driver', formula_key='driver_qty')

    attr_material = SpecAttribute.objects.create(code='material', label='Material', source_kind='sku_field', source_field='material', normalizer='material_name')
    attr_app = SpecAttribute.objects.create(code='application', label='Application', source_kind='sku_field', source_field='application', normalizer='lower_trim')
    attr_colors = SpecAttribute.objects.create(code='total_colors', label='Total Colors', source_kind='derived', derived_key='total_colors', value_type='number')
    return {
        'kg': kg, 'sheet': sheet, 'paper_cat': paper_cat, 'ink_cat': ink_cat, 'step': step,
        'basis_1000pcs': basis_1000pcs, 'basis_driver': basis_driver,
        'attr_material': attr_material, 'attr_app': attr_app, 'attr_colors': attr_colors,
    }


class ConditionOperatorTests(TestCase):
    def setUp(self):
        self.masters = _seed_masters()
        self.recipe = SkuRecipe.objects.create(sku='TEST-001', material='Art Card 300gsm', application='Carton Box', color_spec='4', size_w_mm=100, size_h_mm=200, ups=4)

    def _condition(self, **kwargs):
        return BomTemplateCondition(attribute=self.masters['attr_app'], **kwargs)

    def test_eq_operator(self):
        cond = self._condition(operator='eq', value_text='Carton Box')
        self.assertTrue(evaluate_condition(cond, self.recipe))
        cond2 = self._condition(operator='eq', value_text='Something Else')
        self.assertFalse(evaluate_condition(cond2, self.recipe))

    def test_neq_operator(self):
        cond = self._condition(operator='neq', value_text='Something Else')
        self.assertTrue(evaluate_condition(cond, self.recipe))

    def test_contains_operator(self):
        cond = self._condition(operator='contains', value_text='carton')
        self.assertTrue(evaluate_condition(cond, self.recipe))

    def test_in_operator(self):
        cond = self._condition(operator='in', value_list=['carton box', 'flap box'])
        self.assertTrue(evaluate_condition(cond, self.recipe))

    def test_not_in_operator(self):
        cond = self._condition(operator='not_in', value_list=['flap box'])
        self.assertTrue(evaluate_condition(cond, self.recipe))

    def test_numeric_operators(self):
        cond = BomTemplateCondition(attribute=self.masters['attr_colors'], operator='gt', value_text='2')
        self.assertTrue(evaluate_condition(cond, self.recipe))  # total_colors = 4
        cond = BomTemplateCondition(attribute=self.masters['attr_colors'], operator='lt', value_text='2')
        self.assertFalse(evaluate_condition(cond, self.recipe))
        cond = BomTemplateCondition(attribute=self.masters['attr_colors'], operator='gte', value_text='4')
        self.assertTrue(evaluate_condition(cond, self.recipe))
        cond = BomTemplateCondition(attribute=self.masters['attr_colors'], operator='lte', value_text='4')
        self.assertTrue(evaluate_condition(cond, self.recipe))
        cond = BomTemplateCondition(attribute=self.masters['attr_colors'], operator='between', value_min='1', value_max='5')
        self.assertTrue(evaluate_condition(cond, self.recipe))

    def test_any_wildcard(self):
        cond = self._condition(operator='any')
        self.assertTrue(evaluate_condition(cond, self.recipe))

    def test_missing_attribute_value_fails_numeric_ops(self):
        recipe_no_colors = SkuRecipe.objects.create(sku='TEST-002', color_spec='')
        cond = BomTemplateCondition(attribute=self.masters['attr_colors'], operator='gt', value_text='0')
        self.assertFalse(evaluate_condition(cond, recipe_no_colors))


class TemplateSelectionTests(TestCase):
    def setUp(self):
        self.masters = _seed_masters()
        self.recipe = SkuRecipe.objects.create(sku='TEST-003', material='Art Card 300gsm', application='Carton Box')

    def _make_template(self, code, priority, production_line='OFFSET', status='approved'):
        return BomTemplate.objects.create(code=code, name=code, priority=priority, status=status, production_line=production_line)

    def test_priority_wins(self):
        low = self._make_template('LOW', priority=1)
        high = self._make_template('HIGH', priority=10)
        result = best_template(self.recipe)
        self.assertEqual(result.template.pk, high.pk)
        self.assertFalse(result.is_ambiguous)

    def test_specificity_breaks_priority_tie(self):
        general = self._make_template('GEN', priority=5)
        specific = self._make_template('SPEC', priority=5)
        BomTemplateCondition.objects.create(template=specific, attribute=self.masters['attr_app'], operator='eq', value_text='Carton Box')
        result = best_template(self.recipe)
        self.assertEqual(result.template.pk, specific.pk)
        self.assertFalse(result.is_ambiguous)

    def test_true_tie_is_ambiguous(self):
        a = self._make_template('A', priority=5)
        b = self._make_template('B', priority=5)
        # Force identical updated_at ordering isn't guaranteed, but same priority + same
        # condition count (zero) means specificity also ties.
        result = best_template(self.recipe)
        self.assertTrue(result.is_ambiguous)

    def test_unapproved_template_never_matches(self):
        self._make_template('DRAFT', priority=100, status='draft')
        result = best_template(self.recipe)
        self.assertIsNone(result.template)

    def test_no_match_when_conditions_fail(self):
        t = self._make_template('NOPE', priority=100)
        BomTemplateCondition.objects.create(template=t, attribute=self.masters['attr_app'], operator='eq', value_text='Not This')
        self.assertEqual(find_matching_templates(self.recipe), [])


class GenerationTests(TestCase):
    def setUp(self):
        self.masters = _seed_masters()
        self.user = User.objects.create_user('planner1', password='x')
        self.ink = RawItem.objects.create(
            item_code='INK-001', name='Black Ink', category=self.masters['ink_cat'], uom=self.masters['kg'],
            unit_cost=Decimal('500'),
        )
        self.paper = RawItem.objects.create(
            item_code='PAP-001', name='Art Card 300gsm', category=self.masters['paper_cat'], uom=self.masters['sheet'],
            unit_cost=Decimal('10'), item_role='substrate', attributes={'gsm': 300},
        )

    def _approved_template(self, code='TPL-1'):
        template = BomTemplate.objects.create(code=code, name=code, priority=1, status='approved')
        BomTemplateLine.objects.create(
            template=template, line_code='ink', sequence=1, process_step=self.masters['step'],
            bom_item_type=self.masters['ink_cat'], resolution_mode='fixed', raw_item=self.ink,
            quantity_basis=self.masters['basis_1000pcs'], entry_mode='formula', qty_expression='0.01 * total_colors',
            wastage_expression='0.01', scale_by_colors=False,
        )
        return template

    def test_generate_bom_basic(self):
        template = self._approved_template()
        recipe = SkuRecipe.objects.create(sku='GEN-001', material='Art Card 300gsm', color_spec='4')
        bom = generate_bom_for_sku(recipe, user=self.user)
        self.assertIsNotNone(bom)
        self.assertEqual(bom.status, 'draft')
        self.assertEqual(bom.version, 1)
        self.assertTrue(bom.is_current)
        line = bom.lines.get(line_code='ink')
        # qty_expression: 0.01 * total_colors(=4) = 0.04; wastage 0.01 -> gross = 0.04 * 1.01
        self.assertAlmostEqual(float(line.qty_per_unit), 0.04, places=6)
        self.assertAlmostEqual(float(line.gross_qty), 0.04 * 1.01, places=6)
        self.assertEqual(line.unit_cost_snapshot, Decimal('500'))

    def test_regeneration_creates_new_version_and_supersedes(self):
        self._approved_template()
        recipe = SkuRecipe.objects.create(sku='GEN-002', material='Art Card 300gsm', color_spec='2')
        bom_v1 = generate_bom_for_sku(recipe, user=self.user)
        bom_v2 = generate_bom_for_sku(recipe, user=self.user, mode='manual')
        bom_v1.refresh_from_db()
        self.assertEqual(bom_v2.version, 2)
        self.assertTrue(bom_v2.is_current)
        self.assertFalse(bom_v1.is_current)

    def test_no_matching_template_returns_none(self):
        recipe = SkuRecipe.objects.create(sku='GEN-003', material='Unknown Material')
        bom = generate_bom_for_sku(recipe, user=self.user)
        self.assertIsNone(bom)

    def test_approved_bom_not_mutated_by_spec_drift(self):
        self._approved_template()
        recipe = SkuRecipe.objects.create(sku='GEN-004', material='Art Card 300gsm', color_spec='4')
        bom = generate_bom_for_sku(recipe, user=self.user)
        bom.status = 'approved'
        bom.save(update_fields=['status'])
        original_snapshot = bom.spec_snapshot
        recipe.color_spec = '8'
        recipe.save()
        bom.refresh_from_db()
        self.assertEqual(bom.spec_snapshot, original_snapshot)  # untouched
        self.assertTrue(bom.needs_regeneration)  # but flagged stale

    def test_by_selector_resolution(self):
        template = BomTemplate.objects.create(code='SEL-1', name='Selector', priority=1, status='approved')
        BomTemplateLine.objects.create(
            template=template, line_code='substrate', sequence=1, process_step=self.masters['step'],
            bom_item_type=self.masters['paper_cat'], resolution_mode='by_selector', item_role='substrate',
            selector={'gsm': 'gsm'}, quantity_basis=self.masters['basis_1000pcs'], entry_mode='formula', qty_expression='1',
        )
        recipe = SkuRecipe.objects.create(sku='SEL-001', material='Art Card 300gsm')
        bom = generate_bom_for_sku(recipe, user=self.user)
        self.assertIsNotNone(bom)
        line = bom.lines.get(line_code='substrate')
        self.assertEqual(line.raw_item_id, self.paper.pk)

    def test_by_selector_no_match_warns_not_crashes(self):
        template = BomTemplate.objects.create(code='SEL-2', name='Selector2', priority=1, status='approved')
        BomTemplateLine.objects.create(
            template=template, line_code='substrate', sequence=1, process_step=self.masters['step'],
            bom_item_type=self.masters['paper_cat'], resolution_mode='by_selector', item_role='substrate',
            selector={'gsm': 'gsm'}, quantity_basis=self.masters['basis_1000pcs'], entry_mode='formula', qty_expression='1',
        )
        recipe = SkuRecipe.objects.create(sku='SEL-002', material='Duplex Board 999gsm')
        bom = generate_bom_for_sku(recipe, user=self.user)
        self.assertIsNotNone(bom)
        self.assertFalse(bom.lines.filter(line_code='substrate').exists())
        self.assertTrue(any('substrate' in w for w in bom.warnings))


class DriverLineTests(TestCase):
    """Reproduces the carton reference finding: a consumable scales off another
    line's resolved quantity (e.g. starch = 0.0528 * fluting paper kg), not off
    the FG quantity directly."""

    def setUp(self):
        self.masters = _seed_masters()
        self.user = User.objects.create_user('planner2', password='x')
        self.starch_cat = ItemCategory.objects.create(name='Chemical', code='CHM', default_uom=self.masters['kg'])
        self.flute_item = RawItem.objects.create(
            item_code='FLT-001', name='Fluting Paper', category=self.masters['paper_cat'], uom=self.masters['kg'], unit_cost=Decimal('80'),
        )
        self.starch_item = RawItem.objects.create(
            item_code='STA-001', name='Corn Starch', category=self.starch_cat, uom=self.masters['kg'], unit_cost=Decimal('120'),
        )

    def test_driver_line_ratio(self):
        template = BomTemplate.objects.create(code='DRV-1', name='Driver', priority=1, status='approved', production_line='OFFSET')
        flute_line = BomTemplateLine.objects.create(
            template=template, line_code='flute', sequence=1, process_step=self.masters['step'],
            bom_item_type=self.masters['paper_cat'], resolution_mode='fixed', raw_item=self.flute_item,
            quantity_basis=self.masters['basis_1000pcs'], entry_mode='formula', qty_expression='0.882925', wastage_expression='0',
        )
        starch_line = BomTemplateLine.objects.create(
            template=template, line_code='starch', sequence=2, process_step=self.masters['step'],
            bom_item_type=self.starch_cat, resolution_mode='fixed', raw_item=self.starch_item,
            quantity_basis=self.masters['basis_driver'], entry_mode='formula', qty_expression='driver_qty * 0.0528',
            wastage_expression='0', driver_line=flute_line,
        )
        recipe = SkuRecipe.objects.create(sku='DRV-001', material='Board', product_type='CARTON')
        bom = generate_bom_for_sku(recipe, user=self.user)
        self.assertIsNotNone(bom)
        flute_gross = bom.lines.get(line_code='flute').gross_qty
        starch_gross = bom.lines.get(line_code='starch').gross_qty
        self.assertAlmostEqual(float(starch_gross / flute_gross), 0.0528, places=4)

    def test_driver_cycle_rejected_at_clean(self):
        from django.core.exceptions import ValidationError

        template = BomTemplate.objects.create(code='CYCLE-1', name='Cycle', priority=1, status='draft')
        line_a = BomTemplateLine.objects.create(
            template=template, line_code='a', sequence=1, process_step=self.masters['step'],
            bom_item_type=self.masters['paper_cat'], resolution_mode='fixed', raw_item=self.flute_item,
            quantity_basis=self.masters['basis_1000pcs'], entry_mode='formula', qty_expression='1',
        )
        line_b = BomTemplateLine.objects.create(
            template=template, line_code='b', sequence=2, process_step=self.masters['step'],
            bom_item_type=self.masters['paper_cat'], resolution_mode='fixed', raw_item=self.flute_item,
            quantity_basis=self.masters['basis_1000pcs'], entry_mode='formula', qty_expression='1', driver_line=line_a,
        )
        line_a.driver_line = line_b
        line_a.save()
        with self.assertRaises(ValidationError):
            template.clean()


class ColorPassScalingAndCostTests(TestCase):
    def setUp(self):
        self.masters = _seed_masters()
        self.user = User.objects.create_user('planner3', password='x')
        self.plate_cat = ItemCategory.objects.create(name='Printing Plate', code='PLT', default_uom=self.masters['sheet'])
        self.plate_item = RawItem.objects.create(
            item_code='PLT-001', name='Plate', category=self.plate_cat, uom=self.masters['sheet'], unit_cost=Decimal('250'),
        )

    def test_scale_by_colors_and_passes(self):
        template = BomTemplate.objects.create(code='PLT-T', name='Plate template', priority=1, status='approved')
        BomTemplateLine.objects.create(
            template=template, line_code='plates', sequence=1, process_step=self.masters['step'],
            bom_item_type=self.plate_cat, resolution_mode='fixed', raw_item=self.plate_item,
            quantity_basis=self.masters['basis_1000pcs'], entry_mode='formula', qty_expression='1', wastage_expression='0',
            scale_by_colors=True, scale_by_passes=True,
        )
        recipe = SkuRecipe.objects.create(sku='PLT-001-SKU', material='Art Card', color_spec='4', print_passes=2)
        bom = generate_bom_for_sku(recipe, user=self.user)
        line = bom.lines.get(line_code='plates')
        # 1 * total_colors(4) * print_passes(2) = 8
        self.assertEqual(line.qty_per_unit, Decimal('8.00000000'))
        self.assertEqual(line.line_cost, Decimal('2000.0000'))
        self.assertEqual(bom.total_material_cost, Decimal('2000.0000'))


class SimpleEntryModeTests(TestCase):
    """entry_mode='simple' — a plain per-unit number typed against one exact
    spec combination, as built for the Recipe Wizard's quick-entry flow. No
    formula evaluation, no driver dependency: matches the carton department's
    (and the real offset BOM sheet's) manual-Excel consumption style."""

    def setUp(self):
        self.masters = _seed_masters()
        self.user = User.objects.create_user('planner-simple', password='x')
        self.glue = RawItem.objects.create(
            item_code='ADH-001', name='GMSA Glue 50gm', category=self.masters['ink_cat'], uom=self.masters['kg'],
            unit_cost=Decimal('300'),
        )

    def _template_with_simple_line(self, fixed_qty='0.000625', fixed_wastage_percent='0.05'):
        template = BomTemplate.objects.create(code='SIMPLE-T', name='Simple line template', priority=1, status='approved')
        BomTemplateLine.objects.create(
            template=template, line_code='glue', sequence=1, process_step=self.masters['step'],
            bom_item_type=self.masters['ink_cat'], resolution_mode='fixed', raw_item=self.glue,
            quantity_basis=self.masters['basis_1000pcs'], entry_mode='simple',
            fixed_qty=Decimal(fixed_qty), fixed_wastage_percent=Decimal(fixed_wastage_percent),
        )
        return template

    def test_simple_line_uses_fixed_values_directly(self):
        self._template_with_simple_line()
        recipe = SkuRecipe.objects.create(sku='SIMPLE-001', material='Art Card', color_spec='4')
        bom = generate_bom_for_sku(recipe, user=self.user)
        line = bom.lines.get(line_code='glue')
        self.assertEqual(line.qty_per_unit, Decimal('0.000625').quantize(Decimal('0.00000001')))
        self.assertEqual(line.wastage_percent, Decimal('0.05'))
        expected_gross = (Decimal('0.000625') * Decimal('1.05')).quantize(Decimal('0.00000001'))
        self.assertEqual(line.gross_qty, expected_gross)
        self.assertIn('fixed:', line.resolved_formula)

    def test_simple_line_ignores_sku_size_and_colors(self):
        """A simple line's qty is a constant per spec combination regardless of
        the SKU's own size/color — two different SKUs matching the same
        template get identical consumption, exactly like the reference sheet."""
        self._template_with_simple_line()
        small = SkuRecipe.objects.create(sku='SIMPLE-002', material='Art Card', color_spec='1', size_w_mm=50, size_h_mm=50)
        large = SkuRecipe.objects.create(sku='SIMPLE-003', material='Art Card', color_spec='4', size_w_mm=500, size_h_mm=500)
        bom_small = generate_bom_for_sku(small, user=self.user)
        bom_large = generate_bom_for_sku(large, user=self.user)
        self.assertEqual(bom_small.lines.get(line_code='glue').qty_per_unit, bom_large.lines.get(line_code='glue').qty_per_unit)

    def test_clean_requires_fixed_qty_for_simple_mode(self):
        template = BomTemplate.objects.create(code='SIMPLE-BAD', name='Bad', priority=1)
        line = BomTemplateLine(
            template=template, line_code='glue', sequence=1, process_step=self.masters['step'],
            bom_item_type=self.masters['ink_cat'], resolution_mode='fixed', raw_item=self.glue,
            quantity_basis=self.masters['basis_1000pcs'], entry_mode='simple', fixed_qty=None,
        )
        with self.assertRaises(Exception):
            line.full_clean()

    def test_clean_requires_qty_expression_for_formula_mode(self):
        template = BomTemplate.objects.create(code='FORMULA-BAD', name='Bad', priority=1)
        line = BomTemplateLine(
            template=template, line_code='glue', sequence=1, process_step=self.masters['step'],
            bom_item_type=self.masters['ink_cat'], resolution_mode='fixed', raw_item=self.glue,
            quantity_basis=self.masters['basis_1000pcs'], entry_mode='formula', qty_expression='',
        )
        with self.assertRaises(Exception):
            line.full_clean()


class DictBasedMatchingTests(TestCase):
    """matching.py's entry points also accept an already-resolved spec dict
    (no backing SkuRecipe) — this is what the Recipe Wizard uses to preview a
    match from manually-picked attribute values before any SKU exists."""

    def setUp(self):
        self.masters = _seed_masters()

    def test_evaluate_condition_against_dict(self):
        cond = BomTemplateCondition(attribute=self.masters['attr_app'], operator='eq', value_text='Carton Box')
        self.assertTrue(evaluate_condition(cond, {'application': 'Carton Box'}))
        self.assertFalse(evaluate_condition(cond, {'application': 'Something Else'}))
        self.assertFalse(evaluate_condition(cond, {}))  # missing key -> None -> no match

    def test_find_matching_templates_against_dict_matches_sku_recipe_path(self):
        template = BomTemplate.objects.create(code='DICT-T', name='Dict template', priority=1, status='approved')
        BomTemplateCondition.objects.create(template=template, attribute=self.masters['attr_app'], operator='eq', value_text='Carton Box')
        BomTemplateCondition.objects.create(template=template, attribute=self.masters['attr_colors'], operator='eq', value_text='4')

        recipe = SkuRecipe.objects.create(sku='DICT-SKU', material='Art Card', application='Carton Box', color_spec='4')
        via_recipe = find_matching_templates(recipe)

        from .spec import resolve_sku_spec
        spec = resolve_sku_spec(recipe)
        via_dict = find_matching_templates(spec)

        self.assertEqual([t.pk for t in via_recipe], [t.pk for t in via_dict])
        self.assertEqual(via_dict, [template])

    def test_best_template_against_dict(self):
        template = BomTemplate.objects.create(code='DICT-BEST', name='Dict best', priority=5, status='approved')
        BomTemplateCondition.objects.create(template=template, attribute=self.masters['attr_app'], operator='eq', value_text='Carton Box')
        result = best_template({'application': 'carton box'})
        self.assertEqual(result.template, template)
        self.assertFalse(result.is_ambiguous)
