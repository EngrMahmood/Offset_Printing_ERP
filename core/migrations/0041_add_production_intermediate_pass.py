from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0040_production_change_reason_production_counter_end_and_more'),
    ]

    operations = [
        migrations.AddField(
            model_name='production',
            name='intermediate_pass',
            field=models.BooleanField(default=False, help_text='Mark as an intermediate print pass with no final usable output'),
        ),
    ]
