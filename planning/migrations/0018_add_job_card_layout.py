from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0017_phase_separation_purchase_origin_delivery_date'),
    ]

    operations = [
        migrations.CreateModel(
            name='JobCardLayout',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(default='Job Card Layout', max_length=120)),
                ('layout', models.JSONField(blank=True, default=list)),
                ('is_active', models.BooleanField(default=False)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('updated_at', models.DateTimeField(auto_now=True)),
            ],
            options={
                'ordering': ['-is_active', '-updated_at'],
                'verbose_name': 'Job Card Layout',
                'verbose_name_plural': 'Job Card Layouts',
            },
        ),
    ]
