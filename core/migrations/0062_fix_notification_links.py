from django.db import migrations

def fix_notification_event_links(apps, schema_editor):
    NotificationEvent = apps.get_model('core', 'NotificationEvent')
    Notification = apps.get_model('core', 'Notification')

    # 1. Update link_template for Planning Job events
    events = NotificationEvent.objects.filter(code__in=['job.pending_qc', 'job.qc_approved', 'job.released'])
    for event in events:
        if event.link_template:
            event.link_template = event.link_template.replace('/planning/jobs/', '/planning/job/')
            event.save()

    # 2. Correct existing notification links
    notifications = Notification.objects.filter(link__contains='/planning/jobs/')
    for notification in notifications:
        notification.link = notification.link.replace('/planning/jobs/', '/planning/job/')
        notification.save()


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0061_passwordresetrequest'),
    ]

    operations = [
        migrations.RunPython(fix_notification_event_links),
    ]
