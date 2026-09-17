from django.core.management.base import BaseCommand

from bom.models import SkuBom
from bom.services import generate_bom_for_sku
from planning.models import SkuRecipe


class Command(BaseCommand):
    help = 'Generate BOMs for SKUs against approved templates.'

    def add_arguments(self, parser):
        parser.add_argument('--missing-only', action='store_true', help='Only generate for SKUs with no current BOM.')

    def handle(self, *args, **options):
        qs = SkuRecipe.objects.filter(is_active=True)
        if options['missing_only']:
            has_bom_ids = set(SkuBom.objects.filter(is_current=True).values_list('sku_recipe_id', flat=True))
            qs = qs.exclude(pk__in=has_bom_ids)

        generated, skipped = 0, 0
        for sku_recipe in qs.iterator():
            bom = generate_bom_for_sku(sku_recipe, user=None, mode='auto')
            if bom is None:
                skipped += 1
            else:
                generated += 1

        self.stdout.write(self.style.SUCCESS(f'Generated {generated} BOM(s); {skipped} SKU(s) had no matching template.'))
