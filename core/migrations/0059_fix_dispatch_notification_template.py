from django.db import migrations


def fix_dispatch_template(apps, schema_editor):
    NotificationEvent = apps.get_model('core', 'NotificationEvent')
    try:
        event = NotificationEvent.objects.get(code='dispatch.created')
        event.message_template = 'Dispatch of {{ instance.dispatch_qty }} units registered to location {{ instance.delivery_location.name|default:"client" }}.'
        event.save()
    except NotificationEvent.DoesNotExist:
        pass


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0058_seed_all_erp_notifications'),
    ]

    operations = [
        migrations.RunPython(fix_dispatch_template),
    ]
