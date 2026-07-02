from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0047_alter_jobcard_purchase_sheet_ups_alter_jobcard_ups'),
    ]

    operations = [
        migrations.CreateModel(
            name='ProductType',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=100, unique=True)),
            ],
            options={
                'verbose_name': 'Product Type',
                'verbose_name_plural': 'Product Types',
                'ordering': ['name'],
            },
        ),
    ]
