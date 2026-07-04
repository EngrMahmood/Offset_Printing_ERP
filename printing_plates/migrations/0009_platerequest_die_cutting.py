from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('printing_plates', '0008_platerequest_sets_required'),
    ]

    operations = [
        migrations.AddField(
            model_name='platerequest',
            name='die_cutting',
            field=models.CharField(
                blank=True,
                choices=[('', 'Select'), ('YES', 'Yes'), ('NO', 'No')],
                default='',
                help_text='Die cutting required: Yes or No only.',
                max_length=10,
            ),
        ),
    ]
