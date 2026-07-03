from django.db import migrations, models


DEFAULT_PRINT_COLORS = [
    (1, '1'),
    (2, '2'),
    (3, '3'),
    (4, '4'),
    (5, '1+0'),
    (6, '1+1'),
    (7, '2+0'),
    (8, '2+1'),
    (9, '2+2'),
    (10, '3+0'),
    (11, '3+1'),
    (12, '4+0'),
    (13, '4+1'),
    (14, '4+4'),
]


def seed_print_colors(apps, schema_editor):
    PrintColor = apps.get_model('core', 'PrintColor')
    for sort_order, name in DEFAULT_PRINT_COLORS:
        PrintColor.objects.get_or_create(
            name=name,
            defaults={'sort_order': sort_order, 'is_active': True},
        )


def unseed_print_colors(apps, schema_editor):
    PrintColor = apps.get_model('core', 'PrintColor')
    PrintColor.objects.filter(name__in=[name for _, name in DEFAULT_PRINT_COLORS]).delete()


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0052_production_print_pass_number'),
    ]

    operations = [
        migrations.CreateModel(
            name='PrintColor',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(help_text='Production colour pattern, e.g. 1, 2, 4, 1+1, 2+1', max_length=20, unique=True)),
                ('is_active', models.BooleanField(default=True)),
                ('sort_order', models.PositiveSmallIntegerField(default=0)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
            ],
            options={
                'verbose_name': 'Print Color',
                'verbose_name_plural': 'Print Colors',
                'ordering': ['sort_order', 'name'],
            },
        ),
        migrations.RunPython(seed_print_colors, unseed_print_colors),
    ]
