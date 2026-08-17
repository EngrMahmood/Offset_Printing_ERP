from django.db import migrations


def seed_print_pass_options(apps, schema_editor):
    PrintPassOption = apps.get_model('core', 'PrintPassOption')
    for sort_order, value in enumerate(['1', '2', '3', '4'], start=1):
        PrintPassOption.objects.get_or_create(
            name=value,
            defaults={'is_active': True, 'sort_order': sort_order},
        )


def noop_reverse(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0082_printpassoption'),
    ]

    operations = [
        migrations.RunPython(seed_print_pass_options, noop_reverse),
    ]
