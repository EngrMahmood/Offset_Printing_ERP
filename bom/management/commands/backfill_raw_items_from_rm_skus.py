import re

from django.core.management.base import BaseCommand
from django.db import transaction

from bom.models import ItemCategory, RawItem, UnitOfMeasure
from supply_chain.models import RawMaterialSku

_GSM_RE = re.compile(r'(\d{2,4})\s*gsm', re.IGNORECASE)


class Command(BaseCommand):
    help = 'Creates a bridged Paper/Board RawItem for every active RawMaterialSku (idempotent).'

    @transaction.atomic
    def handle(self, *args, **options):
        category = ItemCategory.objects.filter(code='PAP').first()
        if category is None:
            self.stderr.write(self.style.ERROR('Run seed_bom_masters first (Paper/Board category missing).'))
            return
        sheet_uom, _ = UnitOfMeasure.objects.get_or_create(code='SHEET', defaults={'name': 'Sheet', 'decimal_places': 0})

        created_count = 0
        updated_count = 0
        for rm_sku in RawMaterialSku.objects.filter(is_active=True).select_related('material'):
            gsm_match = _GSM_RE.search(rm_sku.material.name or '')
            attributes = {'sheet_size': rm_sku.purchase_sheet_size}
            if gsm_match:
                attributes['gsm'] = int(gsm_match.group(1))

            raw_item, created = RawItem.objects.get_or_create(
                raw_material_sku=rm_sku,
                defaults={
                    'item_code': f'PAP-{rm_sku.sku}',
                    'name': f'{rm_sku.material.name} — {rm_sku.purchase_sheet_size}',
                    'category': category,
                    'uom': sheet_uom,
                    'specification': rm_sku.purchase_sheet_size,
                    'production_line': 'OFFSET',
                    'item_role': 'substrate',
                    'attributes': attributes,
                    'unit_cost': rm_sku.unit_cost,
                    'pack_size': rm_sku.sheet_packing_pcs or 1,
                    'safety_stock': rm_sku.safety_stock,
                    'max_stock_level': rm_sku.max_stock_level,
                    'lead_time_days': rm_sku.lead_time_days,
                    'is_active': rm_sku.is_active,
                },
            )
            if created:
                created_count += 1
            else:
                raw_item.unit_cost = rm_sku.unit_cost
                raw_item.attributes = attributes
                raw_item.is_active = rm_sku.is_active
                raw_item.save(update_fields=['unit_cost', 'attributes', 'is_active', 'updated_at'])
                updated_count += 1

        self.stdout.write(self.style.SUCCESS(f'Backfill complete: {created_count} created, {updated_count} updated.'))
