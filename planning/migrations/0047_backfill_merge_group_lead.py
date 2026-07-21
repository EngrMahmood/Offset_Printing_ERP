from django.db import migrations


def backfill_leads(apps, schema_editor):
    """Groups accepted before the lead-job model need a lead assigned.

    Without one every member counts as a follower and is blocked from plate
    making, so the group would be stuck. The member with the most allocated ups
    leads, matching how new groups are created.
    """
    MergeGroup = apps.get_model('planning', 'MergeGroup')

    for group in MergeGroup.objects.filter(lead_job__isnull=True):
        lead_item = group.items.order_by('-allocated_ups', 'id').first()
        if not lead_item:
            continue
        group.items.update(is_lead=False)
        lead_item.is_lead = True
        lead_item.save(update_fields=['is_lead'])
        group.lead_job_id = lead_item.planning_job_id
        group.save(update_fields=['lead_job'])


def noop(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0046_mergegroup_lead_job_mergegroupitem_is_lead_and_more'),
    ]

    operations = [
        migrations.RunPython(backfill_leads, noop),
    ]
