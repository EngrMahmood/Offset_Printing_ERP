from django.db import migrations


# GTO1 (single-colour) sheet max ~ 520x740mm (typical GTO46/52 class), GTO2 same press
# class but running 2-colour. SM74 (5-colour) has the largest sheet size and acts as
# the size-gate overflow machine per MACHINE_PLANNING_UPGRADE_PLAN.md.
MACHINE_SPECS = {
    'GTO 1A': dict(group='GTO1', default_colors=1, operational_colors=1,
                   min_l=210, min_w=270, max_l=520, max_w=740),
    'GTO 1B': dict(group='GTO1', default_colors=1, operational_colors=1,
                   min_l=210, min_w=270, max_l=520, max_w=740),
    'GTO 2A': dict(group='GTO2', default_colors=2, operational_colors=2,
                   min_l=210, min_w=270, max_l=520, max_w=740),
    'GTO 2B': dict(group='GTO2', default_colors=2, operational_colors=2,
                   min_l=210, min_w=270, max_l=520, max_w=740),
    'GTO 2C': dict(group='GTO2', default_colors=2, operational_colors=2,
                   min_l=210, min_w=270, max_l=520, max_w=740),
    'SM 74': dict(group='SM74', default_colors=5, operational_colors=5,
                  min_l=210, min_w=270, max_l=740, max_w=1050),
}


def forwards(apps, schema_editor):
    Machine = apps.get_model('core', 'Machine')
    for name, spec in MACHINE_SPECS.items():
        Machine.objects.filter(name=name).update(
            machine_group_code=spec['group'],
            default_colors=spec['default_colors'],
            operational_colors=spec['operational_colors'],
            min_print_length_mm=spec['min_l'],
            min_print_width_mm=spec['min_w'],
            max_print_length_mm=spec['max_l'],
            max_print_width_mm=spec['max_w'],
        )


def backwards(apps, schema_editor):
    Machine = apps.get_model('core', 'Machine')
    Machine.objects.filter(name__in=MACHINE_SPECS.keys()).update(
        machine_group_code='',
        default_colors=1,
        operational_colors=1,
        min_print_length_mm=None,
        min_print_width_mm=None,
        max_print_length_mm=None,
        max_print_width_mm=None,
    )


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0065_machine_color_size_fields'),
    ]

    operations = [
        migrations.RunPython(forwards, backwards),
    ]
