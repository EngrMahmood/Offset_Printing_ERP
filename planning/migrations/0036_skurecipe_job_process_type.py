from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0035_planningjob_job_process_type'),
    ]

    operations = [
        migrations.AddField(
            model_name='skurecipe',
            name='job_process_type',
            field=models.CharField(
                choices=[
                    ('print_and_pack', 'Print + Pack'),
                    ('cut_and_pack', 'Cut & Pack (no printing)'),
                ],
                default='print_and_pack',
                help_text='Default process for jobs using this SKU. Cut & Pack skips printing/plates.',
                max_length=20,
            ),
        ),
    ]
