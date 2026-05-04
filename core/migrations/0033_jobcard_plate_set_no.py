from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0032_jobcard_total_colors'),
    ]

    operations = [
        migrations.AddField(
            model_name='jobcard',
            name='plate_set_no',
            field=models.CharField(max_length=120, null=True, blank=True),
        ),
    ]
