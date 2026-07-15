from django.db import migrations


def repair_plate_making_stages(apps, schema_editor):
    PlanningJob = apps.get_model('planning', 'PlanningJob')
    PLATE_MAKING_STAGES = {'new_plate_making', 'repeat_plate_making'}
    qs = PlanningJob.objects.filter(planning_stage__in=PLATE_MAKING_STAGES)
    for job in qs:
        stage = (job.planning_stage or '').strip()
        expected = 'new_plate_making' if (job.repeat_flag or '').strip() == 'New' else 'repeat_plate_making'
        if stage != expected:
            job.planning_stage = expected
            job.save(update_fields=['planning_stage'])


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0037_skurecipe_legacy_produced'),
    ]

    operations = [
        migrations.RunPython(repair_plate_making_stages, migrations.RunPython.noop),
    ]
