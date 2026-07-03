from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0053_printcolor'),
    ]

    operations = [
        migrations.AddField(
            model_name='machine',
            name='plate_life_impressions',
            field=models.PositiveIntegerField(
                default=25000,
                help_text='Impressions one plate set can run before a replacement set is needed',
            ),
        ),
    ]
