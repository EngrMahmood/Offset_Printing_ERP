from django.db import migrations


def backfill_sku_type(apps, schema_editor):
    """Every existing SKU predates the type field and is a raw material."""
    RawMaterialSku = apps.get_model('supply_chain', 'RawMaterialSku')
    ItemRequestType = apps.get_model('supply_chain', 'ItemRequestType')

    if not RawMaterialSku.objects.filter(sku_type__isnull=True).exists():
        return

    raw_material = (
        ItemRequestType.objects.filter(name__iexact='Raw Material').first()
        or ItemRequestType.objects.filter(code__iexact='RM').first()
    )
    if raw_material is None:
        raw_material = ItemRequestType.objects.create(
            name='Raw Material', code='RM', is_active=True,
        )

    RawMaterialSku.objects.filter(sku_type__isnull=True).update(sku_type=raw_material)


def unset_sku_type(apps, schema_editor):
    RawMaterialSku = apps.get_model('supply_chain', 'RawMaterialSku')
    RawMaterialSku.objects.update(sku_type=None)


class Migration(migrations.Migration):

    dependencies = [
        ('supply_chain', '0013_itemprocurementtimeline_sku_and_more'),
    ]

    operations = [
        migrations.RunPython(backfill_sku_type, unset_sku_type),
    ]
