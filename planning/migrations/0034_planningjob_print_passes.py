from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0033_skurecipe_product_type'),
    ]

    operations = [
        migrations.AddField(
            model_name='planningjob',
            name='print_passes',
            field=models.PositiveSmallIntegerField(
                blank=True,
                help_text='Number of press passes (1, 2, or 3). Total impressions = print sheets × passes.',
                null=True,
            ),
        ),
    ]
