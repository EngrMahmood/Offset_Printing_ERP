from django.db import migrations

SEED_TYPES = [
    ('Raw Material', 'RM'),
    ('Consumable', 'CON'),
    ('Maintenance Item', 'MNT'),
    ('Service', 'SRV'),
    ('Spare Part', 'SPR'),
]


def seed_types(apps, schema_editor):
    ItemRequestType = apps.get_model('supply_chain', 'ItemRequestType')
    for name, code in SEED_TYPES:
        ItemRequestType.objects.get_or_create(name=name, defaults={'code': code, 'is_active': True})


def unseed_types(apps, schema_editor):
    ItemRequestType = apps.get_model('supply_chain', 'ItemRequestType')
    ItemRequestType.objects.filter(name__in=[n for n, _ in SEED_TYPES]).delete()


class Migration(migrations.Migration):

    dependencies = [
        ('supply_chain', '0007_itemrequesttype_itemrequest_itemprocurementtimeline_and_more'),
    ]

    operations = [
        migrations.RunPython(seed_types, unseed_types),
    ]
