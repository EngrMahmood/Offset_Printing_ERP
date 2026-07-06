from django.db import migrations


def repair_plate_making_stages(apps, schema_editor):
    from planning.sku_classification import repair_inconsistent_plate_making_stages

    repair_inconsistent_plate_making_stages()


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0037_skurecipe_legacy_produced'),
    ]

    operations = [
        migrations.RunPython(repair_plate_making_stages, migrations.RunPython.noop),
    ]
