from django.db import migrations


def update_order_fulfillment_target(apps, schema_editor):
    KPITarget = apps.get_model('reports', 'KPITarget')
    KPITarget.objects.filter(kpi_slug='order_fulfillment', year=2026).update(
        min_value=95, target_value=100, max_value=150,
    )


def revert_order_fulfillment_target(apps, schema_editor):
    KPITarget = apps.get_model('reports', 'KPITarget')
    KPITarget.objects.filter(kpi_slug='order_fulfillment', year=2026).update(
        min_value=80, target_value=85, max_value=100,
    )


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0004_seed_kpi_targets_2026'),
    ]

    operations = [
        migrations.RunPython(update_order_fulfillment_target, revert_order_fulfillment_target),
    ]
