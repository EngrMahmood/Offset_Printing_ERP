from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0032_rename_planning_stage_in_production_to_planning_done'),
    ]

    operations = [
        migrations.AddField(
            model_name='skurecipe',
            name='product_type',
            field=models.CharField(blank=True, max_length=100),
        ),
    ]
