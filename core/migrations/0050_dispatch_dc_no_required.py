from django.db import migrations, models
from django.db.models import Q


def backfill_missing_dc_numbers(apps, schema_editor):
    Dispatch = apps.get_model('core', 'Dispatch')
    for row in Dispatch.objects.filter(Q(dc_no__isnull=True) | Q(dc_no='')):
        row.dc_no = f'LEGACY-{row.id}'
        row.save(update_fields=['dc_no'])


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0049_deliverylocation'),
    ]

    operations = [
        migrations.RunPython(backfill_missing_dc_numbers, migrations.RunPython.noop),
        migrations.AlterField(
            model_name='dispatch',
            name='dc_no',
            field=models.CharField(
                help_text='Dispatch Challan / DR number (required; can be shared across multiple Job Cards and SKUs)',
                max_length=50,
            ),
        ),
    ]
