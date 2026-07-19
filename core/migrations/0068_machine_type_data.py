from django.db import migrations


TYPE_BY_NAME_PATTERN = [
    # (substring to match in machine name, lowercased, machine_type)
    ('gto', 'offset_printing'),
    ('sm 74', 'offset_printing'),
    ('sm74', 'offset_printing'),
    ('zongrui', 'offset_printing'),
    ('cutting', 'cutting'),
    ('konica', 'digital_printing'),
    ('sindoh', 'digital_printing'),
]


def forwards(apps, schema_editor):
    Machine = apps.get_model('core', 'Machine')
    for machine in Machine.objects.all():
        lowered = (machine.name or '').lower()
        for pattern, machine_type in TYPE_BY_NAME_PATTERN:
            if pattern in lowered:
                machine.machine_type = machine_type
                machine.save(update_fields=['machine_type'])
                break


def backwards(apps, schema_editor):
    Machine = apps.get_model('core', 'Machine')
    Machine.objects.all().update(machine_type='other')


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0067_machine_type_and_optional_colors'),
    ]

    operations = [
        migrations.RunPython(forwards, backwards),
    ]
