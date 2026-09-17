from django.core.management.base import BaseCommand
from django.db import transaction

from bom.models import ItemCategory, ProcessStep, QuantityBasis, SpecAttribute, UnitOfMeasure

UOMS = [
    ('KG', 'Kilogram', 3), ('SHEET', 'Sheet', 0), ('PCS', 'Pieces', 0),
    ('MTR', 'Metre', 2), ('SQM', 'Square Metre', 4), ('LTR', 'Litre', 3),
    ('ROLL', 'Roll', 3), ('SET', 'Set', 0),
]

CATEGORIES = [
    ('Paper/Board', 'PAP', 'SHEET', 10),
    ('Ink', 'INK', 'KG', 20),
    ('Printing Plate', 'PLT', 'SET', 30),
    ('Lamination Film', 'LAM', 'SQM', 40),
    ('Adhesive', 'ADH', 'KG', 50),
    ('Coating/Varnish', 'CTG', 'KG', 60),
    ('Die/Tooling', 'DIE', 'SET', 70),
    ('Packaging', 'PKG', 'PCS', 80),
    ('Press Consumable', 'CON', 'PCS', 90),
    ('Service', 'SVC', 'SET', 100),
]

PROCESS_STEPS = [
    ('PREPRESS', 'Prepress', 10, 'OFFSET'),
    ('PRINTING', 'Printing', 20, 'OFFSET'),
    ('LAMINATION', 'Lamination', 30, 'OFFSET'),
    ('DIE_CUTTING', 'Die Cutting', 40, 'OFFSET'),
    ('PASTING', 'Pasting', 50, 'OFFSET'),
    ('PACKING', 'Packing', 60, 'OFFSET'),
    ('DISPATCH', 'Dispatch', 70, 'OFFSET'),
]

QUANTITY_BASES = [
    ('per_1000_pcs', 'Per 1000 Pieces', 'order_qty'),
    ('per_1000_sheets', 'Per 1000 Sheets', 'print_sheets'),
    ('per_1000_impressions', 'Per 1000 Impressions', 'planned_total_impressions'),
    ('per_job', 'Per Job', 'job_count'),
    ('per_sqm_printed', 'Per SQM Printed', 'print_sheet_area_sqm'),
    ('per_plate_set', 'Per Plate Set', 'plate_sets'),
    ('per_unit_of_line', 'Per Unit of Driver Line', 'driver_qty'),
]

# code, label, source_kind, source_field/derived_key, value_type, normalizer
# Order here is the Recipe Wizard's cascade order (SpecAttribute.sort_order),
# matching the sequence management actually asked for: product type, then
# material, then size, then color, then application, then die. The remaining
# attributes stay available as template match keys but sit later in the
# cascade since they refine the recipe rather than typically ganging it. This
# is data, not code — reorder or add rows any time from Spec Attributes.
# code, label, source_kind, source_field/derived_key, value_type, normalizer, is_wizard_step
SPEC_ATTRIBUTES = [
    ('product_type', 'Product Type', 'sku_field', 'product_type', 'text', 'lower_trim', True),
    ('material', 'Material', 'sku_field', 'material', 'text', 'material_name', True),
    ('size_w_mm', 'Width (mm)', 'sku_field', 'size_w_mm', 'number', 'none', True),
    ('size_h_mm', 'Height (mm)', 'sku_field', 'size_h_mm', 'number', 'none', True),
    ('color_spec', 'Color Spec', 'sku_field', 'color_spec', 'text', 'lower_trim', True),
    ('application', 'Application', 'sku_field', 'application', 'text', 'lower_trim', True),
    ('die_cutting', 'Die Cutting', 'sku_field', 'die_cutting', 'text', 'yes_no', True),
    ('gsm', 'GSM', 'derived', 'gsm', 'number', 'none', False),
    ('total_colors', 'Total Colors', 'derived', 'total_colors', 'number', 'none', False),
    ('print_passes', 'Print Passes', 'sku_field', 'print_passes', 'number', 'none', False),
    ('lamination_front_and_back', 'Lamination (F&B)', 'sku_field', 'lamination_front_and_back', 'bool', 'none', False),
    ('ups', 'Ups', 'sku_field', 'ups', 'number', 'none', False),
    ('print_sheet_size', 'Print Sheet Size', 'sku_field', 'print_sheet_size', 'text', 'sheet_size', False),
    ('purchase_sheet_size', 'Purchase Sheet Size', 'sku_field', 'purchase_sheet_size', 'text', 'sheet_size', False),
]


class Command(BaseCommand):
    help = 'Idempotently seeds BOM master data: UOMs, categories, process steps, quantity bases, spec attributes.'

    @transaction.atomic
    def handle(self, *args, **options):
        uom_by_code = {}
        for code, name, dp in UOMS:
            uom, created = UnitOfMeasure.objects.update_or_create(
                code=code, defaults={'name': name, 'decimal_places': dp},
            )
            uom_by_code[code] = uom
            self.stdout.write(f"{'Created' if created else 'Updated'} UOM: {code}")

        for order, (name, code, default_uom_code, sort_order) in enumerate(CATEGORIES):
            ItemCategory.objects.update_or_create(
                code=code,
                defaults={'name': name, 'default_uom': uom_by_code.get(default_uom_code), 'sort_order': sort_order},
            )
        self.stdout.write(f'Categories: {len(CATEGORIES)}')

        for code, label, sequence, line in PROCESS_STEPS:
            ProcessStep.objects.update_or_create(
                code=code, defaults={'label': label, 'sequence': sequence, 'production_line': line},
            )
        self.stdout.write(f'Process steps: {len(PROCESS_STEPS)}')

        for code, label, formula_key in QUANTITY_BASES:
            QuantityBasis.objects.update_or_create(
                code=code, defaults={'label': label, 'formula_key': formula_key},
            )
        self.stdout.write(f'Quantity bases: {len(QUANTITY_BASES)}')

        for sort_order, (code, label, source_kind, ref, value_type, normalizer, is_wizard_step) in enumerate(SPEC_ATTRIBUTES):
            defaults = {
                'label': label, 'source_kind': source_kind, 'value_type': value_type,
                'normalizer': normalizer, 'sort_order': sort_order, 'is_wizard_step': is_wizard_step,
            }
            if source_kind == 'sku_field':
                defaults['source_field'] = ref
            elif source_kind == 'derived':
                defaults['derived_key'] = ref
            SpecAttribute.objects.update_or_create(code=code, defaults=defaults)
        self.stdout.write(f'Spec attributes: {len(SPEC_ATTRIBUTES)}')

        self.stdout.write(self.style.SUCCESS('BOM master seed complete.'))
