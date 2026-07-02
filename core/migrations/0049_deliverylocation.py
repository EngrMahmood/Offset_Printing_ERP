from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0048_producttype'),
    ]

    operations = [
        migrations.CreateModel(
            name='DeliveryLocation',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=120, unique=True)),
            ],
            options={
                'verbose_name': 'Delivery Location',
                'verbose_name_plural': 'Delivery Locations',
                'ordering': ['name'],
            },
        ),
    ]
