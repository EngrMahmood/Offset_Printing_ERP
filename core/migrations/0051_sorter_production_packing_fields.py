from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0050_dispatch_dc_no_required'),
    ]

    operations = [
        migrations.CreateModel(
            name='Sorter',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=100)),
                ('employee_code', models.CharField(blank=True, max_length=50, null=True)),
                ('is_active', models.BooleanField(default=True)),
            ],
        ),
        migrations.AddField(
            model_name='production',
            name='entry_type',
            field=models.CharField(
                choices=[('printing', 'Printing'), ('packing', 'Packing')],
                db_index=True,
                default='printing',
                max_length=20,
            ),
        ),
        migrations.AddField(
            model_name='production',
            name='packing_qty',
            field=models.PositiveIntegerField(default=0, help_text='Good pieces packed (dispatchable)'),
        ),
        migrations.AddField(
            model_name='production',
            name='sorting_waste_qty',
            field=models.PositiveIntegerField(default=0, help_text='Pieces rejected during sorting'),
        ),
        migrations.AddField(
            model_name='production',
            name='sorter',
            field=models.ForeignKey(
                blank=True,
                limit_choices_to={'is_active': True},
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                to='core.sorter',
            ),
        ),
        migrations.AlterField(
            model_name='production',
            name='impressions',
            field=models.PositiveIntegerField(
                default=0,
                help_text='Total impressions produced (sheets × passes)',
            ),
        ),
        migrations.AlterField(
            model_name='production',
            name='machine',
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.PROTECT,
                to='core.machine',
            ),
        ),
        migrations.AlterField(
            model_name='production',
            name='output_sheets',
            field=models.PositiveIntegerField(default=0),
        ),
        migrations.AlterField(
            model_name='production',
            name='planned_time',
            field=models.FloatField(default=0, help_text='in minutes'),
        ),
        migrations.AlterField(
            model_name='production',
            name='run_time',
            field=models.FloatField(default=0, help_text='in minutes'),
        ),
    ]
