from django.db import migrations


def merge_types(apps, schema_editor):
    ItemRequestType = apps.get_model('supply_chain', 'ItemRequestType')
    ItemRequest = apps.get_model('supply_chain', 'ItemRequest')

    maintenance = ItemRequestType.objects.filter(name='Maintenance Item').first()
    spare = ItemRequestType.objects.filter(name='Spare Part').first()

    if maintenance and spare:
        ItemRequest.objects.filter(request_type=spare).update(request_type=maintenance)
        spare.delete()

    if maintenance:
        maintenance.name = 'Maintenance Item / Spare Part'
        maintenance.save(update_fields=['name'])


def unmerge_types(apps, schema_editor):
    ItemRequestType = apps.get_model('supply_chain', 'ItemRequestType')
    merged = ItemRequestType.objects.filter(name='Maintenance Item / Spare Part').first()
    if merged:
        merged.name = 'Maintenance Item'
        merged.save(update_fields=['name'])


class Migration(migrations.Migration):

    dependencies = [
        ('supply_chain', '0009_itemrequestdepartment_alter_itemrequest_status_and_more'),
    ]

    operations = [
        migrations.RunPython(merge_types, unmerge_types),
    ]
