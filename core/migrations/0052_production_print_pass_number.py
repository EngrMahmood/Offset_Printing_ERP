from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0051_sorter_production_packing_fields'),
    ]

    operations = [
        migrations.AddField(
            model_name='production',
            name='print_pass_number',
            field=models.PositiveSmallIntegerField(
                blank=True,
                help_text='Which print pass this run belongs to (1..N from planning).',
                null=True,
            ),
        ),
    ]
