from django.db import migrations

CATEGORIES = ['Mechanical', 'Electrical', 'Electronic', 'Pneumatic', 'Hydraulic', 'Other']


def seed_categories(apps, schema_editor):
    FaultCategory = apps.get_model('maintenance', 'FaultCategory')
    for name in CATEGORIES:
        FaultCategory.objects.get_or_create(name=name)


def remove_categories(apps, schema_editor):
    FaultCategory = apps.get_model('maintenance', 'FaultCategory')
    FaultCategory.objects.filter(name__in=CATEGORIES).delete()


class Migration(migrations.Migration):

    dependencies = [
        ('maintenance', '0001_initial'),
    ]

    operations = [
        migrations.RunPython(seed_categories, remove_categories),
    ]
