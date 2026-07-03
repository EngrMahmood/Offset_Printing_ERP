from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0034_planningjob_print_passes'),
    ]

    operations = [
        migrations.AddField(
            model_name='planningjob',
            name='job_process_type',
            field=models.CharField(
                choices=[
                    ('print_and_pack', 'Print + Pack'),
                    ('cut_and_pack', 'Cut & Pack (no printing)'),
                ],
                db_index=True,
                default='print_and_pack',
                max_length=20,
            ),
        ),
    ]
