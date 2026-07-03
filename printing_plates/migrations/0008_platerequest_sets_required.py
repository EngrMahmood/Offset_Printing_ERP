from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('printing_plates', '0007_rename_printing_pl_source_idx_printing_pl_source_f00284_idx'),
    ]

    operations = [
        migrations.AddField(
            model_name='platerequest',
            name='sets_required',
            field=models.PositiveIntegerField(
                blank=True,
                help_text='Number of plate sets ordered (all issued to production together for now)',
                null=True,
            ),
        ),
    ]
