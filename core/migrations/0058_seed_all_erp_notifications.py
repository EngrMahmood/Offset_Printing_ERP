from django.db import migrations


def seed_extended_erp_notifications(apps, schema_editor):
    NotificationEvent = apps.get_model('core', 'NotificationEvent')
    NotificationRule = apps.get_model('core', 'NotificationRule')

    # Seed events
    events_data = [
        {
            'code': 'job.pending_qc',
            'name': 'Job Card Pending QC',
            'description': 'Fired when a planner submits a job card for QC inspection.',
            'module': 'Planning',
            'title_template': 'Job Card pending QC: {{ instance.jc_number }}',
            'message_template': 'Job {{ instance.job_name }} (Qty: {{ instance.qty }}) is ready for QC approval.',
            'link_template': '/planning/job/{{ instance.id }}/',
        },
        {
            'code': 'job.qc_approved',
            'name': 'Job Card QC Approved',
            'description': 'Fired when a job card is approved by QC.',
            'module': 'QC',
            'title_template': 'Job Card approved: {{ instance.jc_number }}',
            'message_template': 'Job {{ instance.job_name }} approved by QC.',
            'link_template': '/planning/job/{{ instance.id }}/',
        },
        {
            'code': 'job.released',
            'name': 'Job Card Released',
            'description': 'Fired when a job card is released to production.',
            'module': 'Planning',
            'title_template': 'Job Card released: {{ instance.jc_number }}',
            'message_template': 'Job {{ instance.job_name }} is released for production.',
            'link_template': '/planning/job/{{ instance.id }}/',
        },
        {
            'code': 'dispatch.created',
            'name': 'New Dispatch Created',
            'description': 'Fired when a new dispatch record is saved.',
            'module': 'Dispatch',
            'title_template': 'Dispatch created for JC: {{ instance.job_card.jc_number }}',
            'message_template': 'Dispatch of {{ instance.quantity }} units registered to location {{ instance.delivery_location.name|default:"client" }}.',
            'link_template': '/dispatch-records/',
        },
        {
            'code': 'override.requested',
            'name': 'Edit Override Requested',
            'description': 'Fired when a user requests an edit override.',
            'module': 'Core',
            'title_template': 'Override requested: {{ instance.entity_type }} (ID: {{ instance.record_id }})',
            'message_template': 'Reason: {{ instance.reason }}',
            'link_template': '/override-requests/',
        },
        {
            'code': 'production.submitted',
            'name': 'Production Entry Submitted',
            'description': 'Fired when a production shift run entry is submitted.',
            'module': 'Production',
            'title_template': 'Production submitted: {{ instance.job_card.jc_number }}',
            'message_template': 'Shift {{ instance.shift }} run by Operator {{ instance.operator.name|default:"Staff" }} submitted.',
            'link_template': '/production-records/',
        },
    ]

    events_map = {}
    for data in events_data:
        event, _ = NotificationEvent.objects.update_or_create(
            code=data['code'],
            defaults=data
        )
        events_map[data['code']] = event

    # Seed rules
    rules_data = [
        # job.pending_qc
        ('job.pending_qc', 'role', 'qc'),
        ('job.pending_qc', 'role', 'admin'),
        
        # job.qc_approved
        ('job.qc_approved', 'role', 'planner'),
        ('job.qc_approved', 'role', 'graphics_designer'),
        ('job.qc_approved', 'role', 'admin'),
        
        # job.released
        ('job.released', 'role', 'production'),
        ('job.released', 'role', 'production_manager'),
        ('job.released', 'role', 'admin'),
        
        # dispatch.created
        ('dispatch.created', 'role', 'dispatch'),
        ('dispatch.created', 'role', 'admin'),
        ('dispatch.created', 'role', 'manager'),
        
        # override.requested
        ('override.requested', 'role', 'admin'),
        ('override.requested', 'role', 'manager'),

        # production.submitted
        ('production.submitted', 'role', 'production_manager'),
        ('production.submitted', 'role', 'admin'),
        ('production.submitted', 'role', 'manager'),
    ]

    for event_code, recipient_type, target in rules_data:
        event = events_map.get(event_code)
        if not event:
            continue
        
        defaults = {
            'enabled': True,
            'recipient_type': recipient_type,
            'exclude_actor': True,
            'in_app_enabled': True,
        }
        if recipient_type == 'role':
            defaults['role'] = target
            
        NotificationRule.objects.get_or_create(
            event=event,
            recipient_type=recipient_type,
            role=target if recipient_type == 'role' else None,
            defaults=defaults
        )


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0057_notificationevent_userprofile_department_and_more'),
    ]

    operations = [
        migrations.RunPython(seed_extended_erp_notifications),
    ]
