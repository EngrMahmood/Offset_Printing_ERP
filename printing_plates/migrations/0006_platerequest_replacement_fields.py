from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('printing_plates', '0005_alter_platerequest_planning_job'),
    ]

    operations = [
        migrations.AddField(
            model_name='platerequest',
            name='damaged_colors',
            field=models.CharField(blank=True, help_text='Colours needing remake, e.g. Cyan, Magenta', max_length=120),
        ),
        migrations.AddField(
            model_name='platerequest',
            name='replacement_reason',
            field=models.CharField(
                blank=True,
                choices=[
                    ('damaged_during_run', 'Damaged during run'),
                    ('damaged_before_printing', 'Damaged before printing'),
                    ('wrong_plates_received', 'Wrong plates received'),
                    ('worn_out', 'Worn out'),
                    ('vendor_defect', 'Vendor defect'),
                    ('other', 'Other'),
                ],
                default='',
                max_length=40,
            ),
        ),
        migrations.AddField(
            model_name='platerequest',
            name='replaces_request',
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='replacement_requests',
                to='printing_plates.platerequest',
            ),
        ),
        migrations.AddIndex(
            model_name='platerequest',
            index=models.Index(fields=['source'], name='printing_pl_source_idx'),
        ),
    ]
