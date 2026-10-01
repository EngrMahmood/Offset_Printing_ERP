from django.db import migrations


def seed_print_pass_option_5(apps, schema_editor):
    """Adds '5' passes to the admin-editable master list (Master Data ->
    Print Pass Options), needed for multi-form books whose Inner pages take
    more passes than their Cover (e.g. a 10-page inner section printed 2-up
    across 5 press passes). Not a hard-coded limit — admins can still add
    further values from the UI; this just seeds the one value already known
    to be needed."""
    PrintPassOption = apps.get_model('core', 'PrintPassOption')
    PrintPassOption.objects.get_or_create(
        name='5', defaults={'is_active': True, 'sort_order': 5},
    )


def noop_reverse(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0090_alter_machine_machine_type'),
    ]

    operations = [
        migrations.RunPython(seed_print_pass_option_5, noop_reverse),
    ]
