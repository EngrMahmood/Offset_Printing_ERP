from django.db import migrations

SEED_VALUES = ['UV', 'Lamination Gloss', 'Lamination Matt', 'NO']


def forwards(apps, schema_editor):
    ApplicationType = apps.get_model('core', 'ApplicationType')
    for name in SEED_VALUES:
        ApplicationType.objects.get_or_create(name=name)


def backwards(apps, schema_editor):
    ApplicationType = apps.get_model('core', 'ApplicationType')
    ApplicationType.objects.filter(name__in=SEED_VALUES).delete()


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0069_application_type'),
    ]

    operations = [
        migrations.RunPython(forwards, backwards),
    ]
