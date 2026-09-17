"""BOM app models.

Layers, bottom to top:

  UnitOfMeasure / ItemCategory / ProcessStep   -- small lookups
  RawItem / RawItemCostHistory                 -- unified raw material + consumable master
  SpecAttribute                                -- soft-coded SKU spec signature (matching keys)
  QuantityBasis                                -- soft-coded quantity formula families
  BomTemplate / BomTemplateCondition / BomTemplateLine
                                                -- spec-matched recipe templates (the "rules")
  SkuBom / SkuBomLine                          -- generated, costed BOM for one SKU

Everything here is additive to the existing schema: SkuRecipe, RawMaterialSku,
StockTransaction and ItemRequest are read and bridged to, never modified.
"""
from __future__ import annotations

from decimal import Decimal

from django.conf import settings
from django.db import models
from django.utils import timezone

PRODUCTION_LINE_CHOICES = [
    ('OFFSET', 'Offset'),
    ('CARTON', 'Carton'),
]

MASTER_DATA_STATUS_CHOICES = [
    ('draft', 'Draft'),
    ('pending_review', 'Pending Review'),
    ('reviewed', 'Pending Approval (Manager)'),
    ('approved', 'Approved'),
]


# ---------------------------------------------------------------------------
# Small lookups
# ---------------------------------------------------------------------------

class UnitOfMeasure(models.Model):
    code = models.CharField(max_length=20, unique=True)
    name = models.CharField(max_length=60)
    decimal_places = models.PositiveSmallIntegerField(default=3)
    is_active = models.BooleanField(default=True, db_index=True)

    class Meta:
        ordering = ['code']
        verbose_name = 'Unit of Measure'
        verbose_name_plural = 'Units of Measure'

    def __str__(self):
        return self.code


class ItemCategory(models.Model):
    name = models.CharField(max_length=100, unique=True)
    code = models.CharField(max_length=10, unique=True, help_text='Prefix used in the item code, e.g. INK, PLT')
    default_uom = models.ForeignKey(UnitOfMeasure, on_delete=models.PROTECT, related_name='categories', null=True, blank=True)
    sort_order = models.PositiveIntegerField(default=0)
    is_active = models.BooleanField(default=True, db_index=True)

    class Meta:
        ordering = ['sort_order', 'name']
        verbose_name = 'Item Category'
        verbose_name_plural = 'Item Categories'

    def __str__(self):
        return self.name


class ProcessStep(models.Model):
    code = models.CharField(max_length=30, unique=True)
    label = models.CharField(max_length=100)
    sequence = models.PositiveIntegerField(default=0)
    production_line = models.CharField(max_length=10, choices=PRODUCTION_LINE_CHOICES, default='OFFSET')
    is_active = models.BooleanField(default=True, db_index=True)

    class Meta:
        ordering = ['production_line', 'sequence']

    def __str__(self):
        return self.label


# ---------------------------------------------------------------------------
# Raw item / consumable master
# ---------------------------------------------------------------------------

class RawItem(models.Model):
    """Unified raw material + consumable master (all categories, all production lines).

    Paper/board items are bridged to the existing ``supply_chain.RawMaterialSku``
    via ``raw_material_sku`` so stock transactions, demand-gap and procurement
    keep working against the single existing stock identity for paper.
    """

    item_code = models.CharField(max_length=30, unique=True)
    name = models.CharField(max_length=200)
    category = models.ForeignKey(ItemCategory, on_delete=models.PROTECT, related_name='raw_items')
    uom = models.ForeignKey(UnitOfMeasure, on_delete=models.PROTECT, related_name='raw_items')
    specification = models.CharField(max_length=255, blank=True)
    brand_grade = models.CharField(max_length=120, blank=True)
    production_line = models.CharField(max_length=10, choices=PRODUCTION_LINE_CHOICES, default='OFFSET')

    item_role = models.SlugField(
        max_length=50, blank=True, db_index=True,
        help_text='Machine-matchable role, e.g. substrate, ink, plate, lam_film, glue. '
                   'Used by template lines with resolution_mode=by_selector.',
    )
    attributes = models.JSONField(
        default=dict, blank=True,
        help_text='Free-form matchable attributes, e.g. {"gsm": 300, "sheet_size": "20*30", "shade": "black"}.',
    )

    unit_cost = models.DecimalField(max_digits=14, decimal_places=6, default=Decimal('0'))
    cost_updated_at = models.DateTimeField(null=True, blank=True)

    default_vendor = models.ForeignKey('core.Vendor', on_delete=models.SET_NULL, null=True, blank=True, related_name='bom_raw_items')

    pack_size = models.DecimalField(max_digits=14, decimal_places=4, default=Decimal('1'))
    moq = models.DecimalField(max_digits=14, decimal_places=4, default=Decimal('0'))
    safety_stock = models.DecimalField(max_digits=14, decimal_places=4, default=Decimal('0'))
    max_stock_level = models.DecimalField(max_digits=14, decimal_places=4, default=Decimal('0'))
    lead_time_days = models.PositiveIntegerField(default=1)

    raw_material_sku = models.ForeignKey(
        'supply_chain.RawMaterialSku', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='bom_raw_items',
        help_text='Bridge to the existing paper/board stock identity in Supply Chain.',
    )

    is_active = models.BooleanField(default=True, db_index=True)
    deleted_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='bom_raw_items_deleted')
    deleted_at = models.DateTimeField(null=True, blank=True)

    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='bom_raw_items_created')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['category__sort_order', 'name']
        verbose_name = 'Raw Item'
        verbose_name_plural = 'Raw Items'

    def __str__(self):
        return f'{self.item_code} — {self.name}'

    def record_cost_change(self, unit_cost, user=None, source='manual'):
        """Update unit_cost and append an audit row. Wrapped in a transaction by the caller."""
        self.unit_cost = unit_cost
        self.cost_updated_at = timezone.now()
        self.save(update_fields=['unit_cost', 'cost_updated_at', 'updated_at'])
        RawItemCostHistory.objects.create(
            raw_item=self, unit_cost=unit_cost, effective_from=timezone.now(),
            source=source, created_by=user,
        )


class RawItemCostHistory(models.Model):
    raw_item = models.ForeignKey(RawItem, on_delete=models.CASCADE, related_name='cost_history')
    unit_cost = models.DecimalField(max_digits=14, decimal_places=6)
    effective_from = models.DateTimeField(default=timezone.now)
    source = models.CharField(max_length=30, default='manual')
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='bom_cost_changes')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-effective_from']
        verbose_name_plural = 'Raw Item Cost History'

    def __str__(self):
        return f'{self.raw_item.item_code} @ {self.unit_cost} ({self.effective_from:%Y-%m-%d})'


# ---------------------------------------------------------------------------
# Soft-coded spec attribute engine
# ---------------------------------------------------------------------------

class SpecAttribute(models.Model):
    """A named, soft-coded slot in a SKU's 'spec signature'.

    Adding a new match key (e.g. machine_name) is a data operation: pick an
    existing SkuRecipe field from the dropdown (validated in clean()) or
    define a regex capture / derived resolver. No code change, no migration.
    """
    SOURCE_KIND_CHOICES = [
        ('sku_field', 'SkuRecipe field/property'),
        ('regex_capture', 'Regex capture from SKU code'),
        ('derived', 'Derived resolver'),
    ]
    VALUE_TYPE_CHOICES = [
        ('text', 'Text'),
        ('number', 'Number'),
        ('bool', 'Boolean'),
        ('choice', 'Choice'),
    ]
    NORMALIZER_CHOICES = [
        ('none', 'None'),
        ('lower_trim', 'Lowercase + trim'),
        ('material_name', 'Material name'),
        ('sheet_size', 'Sheet size'),
        ('yes_no', 'Yes/No'),
    ]

    code = models.SlugField(max_length=60, unique=True)
    label = models.CharField(max_length=120)
    sort_order = models.PositiveIntegerField(default=0)
    is_active = models.BooleanField(default=True, db_index=True)
    is_match_key = models.BooleanField(default=True, help_text='Available as a BomTemplateCondition attribute.')
    is_wizard_step = models.BooleanField(
        default=False,
        help_text='Shown as a cascade step in the Recipe Wizard (in sort_order). '
                   'Leave off for refinement attributes (e.g. GSM, total colors) that are '
                   'still valid template match keys but implied by an earlier step.',
    )

    source_kind = models.CharField(max_length=20, choices=SOURCE_KIND_CHOICES, default='sku_field')
    source_field = models.CharField(max_length=100, blank=True, help_text='For sku_field: a field/property name on planning.SkuRecipe.')
    pattern = models.CharField(max_length=255, blank=True, help_text='For regex_capture: a regex applied to SkuRecipe.sku.')
    capture_group = models.PositiveSmallIntegerField(default=1)
    derived_key = models.CharField(max_length=60, blank=True, help_text='For derived: a key registered in bom.spec.DERIVED_RESOLVERS.')

    value_type = models.CharField(max_length=10, choices=VALUE_TYPE_CHOICES, default='text')
    normalizer = models.CharField(max_length=20, choices=NORMALIZER_CHOICES, default='none')

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['sort_order', 'label']

    def __str__(self):
        return self.label

    def clean(self):
        from django.core.exceptions import ValidationError
        from . import spec as spec_module

        errors = {}
        if self.source_kind == 'sku_field':
            if not self.source_field:
                errors['source_field'] = 'Required when source is a SkuRecipe field.'
            elif self.source_field not in spec_module.available_sku_fields():
                errors['source_field'] = f'"{self.source_field}" is not a recognised SkuRecipe field or property.'
        elif self.source_kind == 'regex_capture':
            if not self.pattern:
                errors['pattern'] = 'Required when source is a regex capture.'
            else:
                import re
                try:
                    re.compile(self.pattern)
                except re.error as exc:
                    errors['pattern'] = f'Invalid regex: {exc}'
        elif self.source_kind == 'derived':
            if not self.derived_key:
                errors['derived_key'] = 'Required when source is a derived resolver.'
            elif self.derived_key not in spec_module.DERIVED_RESOLVERS:
                errors['derived_key'] = f'"{self.derived_key}" is not a registered derived resolver.'
        if errors:
            raise ValidationError(errors)


class QuantityBasis(models.Model):
    """A named family of consumption formulas (soft-coded label over bom.formula usage)."""
    code = models.SlugField(max_length=60, unique=True)
    label = models.CharField(max_length=120)
    formula_key = models.CharField(max_length=60, help_text='Driver variable this basis exposes to line formulas, e.g. order_qty.')
    unit_hint = models.CharField(max_length=60, blank=True)
    is_active = models.BooleanField(default=True, db_index=True)

    class Meta:
        ordering = ['label']
        verbose_name_plural = 'Quantity Bases'

    def __str__(self):
        return self.label


# ---------------------------------------------------------------------------
# Templates (the "rules")
# ---------------------------------------------------------------------------

class BomTemplate(models.Model):
    code = models.SlugField(max_length=60, unique=True)
    name = models.CharField(max_length=150)
    description = models.TextField(blank=True)
    production_line = models.CharField(max_length=10, choices=PRODUCTION_LINE_CHOICES, default='OFFSET')
    priority = models.IntegerField(default=0, help_text='Higher priority wins when multiple templates match a SKU.')

    status = models.CharField(max_length=20, choices=MASTER_DATA_STATUS_CHOICES, default='draft')
    reviewed_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='bom_templates_reviewed')
    reviewed_at = models.DateTimeField(null=True, blank=True)
    approved_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='bom_templates_approved')
    approved_at = models.DateTimeField(null=True, blank=True)
    rejection_comment = models.TextField(blank=True)

    is_active = models.BooleanField(default=True, db_index=True)
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='bom_templates_created')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-priority', 'name']

    def __str__(self):
        return f'{self.code} — {self.name}'

    def clean(self):
        from django.core.exceptions import ValidationError
        if self.pk is None:
            # A brand-new template has no saved lines yet (the line formset
            # saves after the template itself) — nothing to check for cycles.
            return
        cycle = self._find_driver_cycle()
        if cycle:
            raise ValidationError(f'Circular driver_line reference among lines: {" -> ".join(cycle)}')

    def _find_driver_cycle(self):
        """Return a list of line_codes forming a cycle, or None."""
        if self.pk is None:
            return None
        lines = {l.line_code: l for l in self.lines.all()}
        WHITE, GRAY, BLACK = 0, 1, 2
        color = {code: WHITE for code in lines}
        path = []

        def visit(code):
            color[code] = GRAY
            path.append(code)
            line = lines.get(code)
            driver_code = line.driver_line.line_code if line and line.driver_line_id else None
            if driver_code:
                if driver_code not in lines:
                    pass
                elif color.get(driver_code) == GRAY:
                    return path[path.index(driver_code):] + [driver_code]
                elif color.get(driver_code) == WHITE:
                    result = visit(driver_code)
                    if result:
                        return result
            path.pop()
            color[code] = BLACK
            return None

        for code in list(lines):
            if color[code] == WHITE:
                result = visit(code)
                if result:
                    return result
        return None

    def ordered_lines(self):
        """Lines in topological order of driver_line (drivers evaluated before dependents)."""
        lines = list(self.lines.select_related('driver_line').all())
        by_code = {l.line_code: l for l in lines}
        resolved = []
        seen = set()

        def resolve(line):
            if line.line_code in seen:
                return
            if line.driver_line_id and line.driver_line.line_code in by_code:
                resolve(by_code[line.driver_line.line_code])
            seen.add(line.line_code)
            resolved.append(line)

        for line in sorted(lines, key=lambda l: l.sequence):
            resolve(line)
        return resolved


OPERATOR_CHOICES = [
    ('eq', 'Equals'),
    ('neq', 'Not equals'),
    ('in', 'In list'),
    ('not_in', 'Not in list'),
    ('contains', 'Contains'),
    ('gt', 'Greater than'),
    ('gte', 'Greater or equal'),
    ('lt', 'Less than'),
    ('lte', 'Less or equal'),
    ('between', 'Between'),
    ('any', 'Any (wildcard)'),
]


class BomTemplateCondition(models.Model):
    template = models.ForeignKey(BomTemplate, on_delete=models.CASCADE, related_name='conditions')
    attribute = models.ForeignKey(SpecAttribute, on_delete=models.PROTECT, related_name='template_conditions')
    operator = models.CharField(max_length=10, choices=OPERATOR_CHOICES, default='eq')
    value_text = models.CharField(max_length=255, blank=True)
    value_min = models.CharField(max_length=100, blank=True)
    value_max = models.CharField(max_length=100, blank=True)
    value_list = models.JSONField(default=list, blank=True)

    class Meta:
        ordering = ['attribute__sort_order']

    def __str__(self):
        return f'{self.attribute.code} {self.operator} {self.value_text or self.value_list or ""}'


class BomTemplateLine(models.Model):
    RESOLUTION_MODE_CHOICES = [
        ('fixed', 'Fixed raw item'),
        ('by_selector', 'Select by role + attributes'),
        ('from_sku_substrate', "SKU's own substrate (paper bridge)"),
    ]

    template = models.ForeignKey(BomTemplate, on_delete=models.CASCADE, related_name='lines')
    line_code = models.SlugField(max_length=60)
    sequence = models.PositiveIntegerField(default=0)
    process_step = models.ForeignKey(ProcessStep, on_delete=models.PROTECT, related_name='template_lines')
    bom_item_type = models.ForeignKey(ItemCategory, on_delete=models.PROTECT, related_name='template_lines')

    resolution_mode = models.CharField(max_length=20, choices=RESOLUTION_MODE_CHOICES, default='fixed')
    raw_item = models.ForeignKey(RawItem, on_delete=models.PROTECT, null=True, blank=True, related_name='template_lines')
    item_role = models.SlugField(max_length=50, blank=True)
    selector = models.JSONField(default=dict, blank=True, help_text='For by_selector: {attribute_name: formula_expression}.')

    ENTRY_MODE_CHOICES = [
        ('simple', 'Simple (plain number)'),
        ('formula', 'Formula'),
    ]
    entry_mode = models.CharField(
        max_length=10, choices=ENTRY_MODE_CHOICES, default='simple',
        help_text='Simple: type a fixed qty/wastage for this exact spec combination. '
                   'Formula: an expression that scales across every size/spec matching this template.',
    )
    fixed_qty = models.DecimalField(
        max_digits=18, decimal_places=8, null=True, blank=True,
        help_text='Used when entry_mode=simple: consumption per one FG unit.',
    )
    fixed_wastage_percent = models.DecimalField(
        max_digits=7, decimal_places=4, null=True, blank=True,
        help_text='Used when entry_mode=simple: e.g. 0.05 for 5%.',
    )

    quantity_basis = models.ForeignKey(QuantityBasis, on_delete=models.PROTECT, related_name='template_lines')
    qty_expression = models.CharField(
        max_length=500, blank=True,
        help_text='Used when entry_mode=formula: evaluated against the SKU spec + line_qty() of prior lines.',
    )
    driver_line = models.ForeignKey('self', on_delete=models.SET_NULL, null=True, blank=True, related_name='dependents')
    scale_by_colors = models.BooleanField(default=False)
    scale_by_passes = models.BooleanField(default=False)

    wastage_expression = models.CharField(
        max_length=200, blank=True, default='0',
        help_text='Used when entry_mode=formula: formula or constant, e.g. 0.05.',
    )

    is_optional = models.BooleanField(default=False)
    notes = models.CharField(max_length=255, blank=True)

    class Meta:
        ordering = ['sequence']
        unique_together = ('template', 'line_code')

    def __str__(self):
        return f'{self.template.code}:{self.line_code}'

    def clean(self):
        from django.core.exceptions import ValidationError
        errors = {}
        if self.resolution_mode == 'fixed' and not self.raw_item_id:
            errors['raw_item'] = 'Required when resolution mode is "Fixed raw item".'
        if self.resolution_mode == 'by_selector' and not self.item_role:
            errors['item_role'] = 'Required when resolution mode is "Select by role + attributes".'
        if self.driver_line_id and self.driver_line_id == self.pk:
            errors['driver_line'] = 'A line cannot drive itself.'
        if self.entry_mode == 'simple' and self.fixed_qty is None:
            errors['fixed_qty'] = 'Required when entry mode is "Simple".'
        if self.entry_mode == 'formula' and not self.qty_expression:
            errors['qty_expression'] = 'Required when entry mode is "Formula".'
        if errors:
            raise ValidationError(errors)


# ---------------------------------------------------------------------------
# Generated, per-SKU BOM
# ---------------------------------------------------------------------------

class SkuBom(models.Model):
    GENERATION_MODE_CHOICES = [
        ('auto', 'Auto-generated'),
        ('manual', 'Manually created'),
        ('auto_edited', 'Auto-generated, edited'),
        ('imported', 'Imported'),
    ]

    sku_recipe = models.ForeignKey('planning.SkuRecipe', on_delete=models.PROTECT, related_name='boms')
    version = models.PositiveIntegerField(default=1)
    is_current = models.BooleanField(default=True, db_index=True)
    source_template = models.ForeignKey(BomTemplate, on_delete=models.SET_NULL, null=True, blank=True, related_name='generated_boms')
    generation_mode = models.CharField(max_length=20, choices=GENERATION_MODE_CHOICES, default='auto')
    spec_snapshot = models.JSONField(default=dict, blank=True)

    status = models.CharField(max_length=20, choices=MASTER_DATA_STATUS_CHOICES, default='draft')
    reviewed_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='sku_boms_reviewed')
    reviewed_at = models.DateTimeField(null=True, blank=True)
    approved_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='sku_boms_approved')
    approved_at = models.DateTimeField(null=True, blank=True)
    rejection_comment = models.TextField(blank=True)

    total_material_cost = models.DecimalField(max_digits=14, decimal_places=4, default=Decimal('0'))
    cost_computed_at = models.DateTimeField(null=True, blank=True)
    warnings = models.JSONField(default=list, blank=True)

    is_active = models.BooleanField(default=True, db_index=True)
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='sku_boms_created')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']
        unique_together = ('sku_recipe', 'version')

    def __str__(self):
        return f'{self.sku_recipe.sku} v{self.version} ({self.status})'

    @property
    def needs_regeneration(self):
        if not self.spec_snapshot:
            return False
        from .spec import resolve_sku_spec
        current = resolve_sku_spec(self.sku_recipe)
        return current != self.spec_snapshot


class SkuBomLine(models.Model):
    bom = models.ForeignKey(SkuBom, on_delete=models.CASCADE, related_name='lines')
    sequence = models.PositiveIntegerField(default=0)
    line_code = models.SlugField(max_length=60, blank=True)
    process_step = models.ForeignKey(ProcessStep, on_delete=models.PROTECT, related_name='sku_bom_lines')
    raw_item = models.ForeignKey(RawItem, on_delete=models.PROTECT, related_name='sku_bom_lines')
    uom_code = models.CharField(max_length=20, blank=True)

    qty_per_unit = models.DecimalField(max_digits=18, decimal_places=8, default=Decimal('0'))
    wastage_percent = models.DecimalField(max_digits=7, decimal_places=4, default=Decimal('0'))
    gross_qty = models.DecimalField(max_digits=18, decimal_places=8, default=Decimal('0'))
    unit_cost_snapshot = models.DecimalField(max_digits=14, decimal_places=6, default=Decimal('0'))
    line_cost = models.DecimalField(max_digits=14, decimal_places=4, default=Decimal('0'))

    source_template_line = models.ForeignKey(BomTemplateLine, on_delete=models.SET_NULL, null=True, blank=True, related_name='generated_lines')
    resolved_formula = models.CharField(max_length=500, blank=True)
    is_manual_override = models.BooleanField(default=False)
    notes = models.CharField(max_length=255, blank=True)

    class Meta:
        ordering = ['sequence']

    def __str__(self):
        return f'{self.bom} / {self.raw_item.item_code}'
