from django.db import migrations


PRODUCT_TYPE_UNIT_DEFAULTS = {
    # A4/A3 rim paper: fixed 500 sheets/rim across the whole category.
    'A4 Rim': ('rim', 500),
    'A3 Rim': ('rim', 500),
    # Report Books are always books, but the page count varies per SKU
    # (a plain 100-page book vs. a 50-page-per-ply NCR triplicate), so no
    # default_pcs_per_unit — each SKU master sets its own.
    'Report Books': ('book', None),
}


def seed_unit_defaults(apps, schema_editor):
    ProductType = apps.get_model('core', 'ProductType')
    for name, (unit_type, pcs_per_unit) in PRODUCT_TYPE_UNIT_DEFAULTS.items():
        ProductType.objects.filter(name__iexact=name).update(
            default_unit_type=unit_type,
            default_pcs_per_unit=pcs_per_unit,
        )


def unseed_unit_defaults(apps, schema_editor):
    ProductType = apps.get_model('core', 'ProductType')
    for name in PRODUCT_TYPE_UNIT_DEFAULTS:
        ProductType.objects.filter(name__iexact=name).update(
            default_unit_type='pcs',
            default_pcs_per_unit=None,
        )


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0088_producttype_default_pcs_per_unit_and_more'),
    ]

    operations = [
        migrations.RunPython(seed_unit_defaults, unseed_unit_defaults),
    ]
