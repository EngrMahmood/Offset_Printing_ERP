from django.db import migrations, models


def migrate_planning_stage_values(apps, schema_editor):
    PlanningJob = apps.get_model('planning', 'PlanningJob')
    JobCard = apps.get_model('core', 'JobCard')
    PlanningJob.objects.filter(planning_stage='in_production').update(planning_stage='planning_done')
    released_planning_job_ids = JobCard.objects.filter(
        status__in=['released', 'in_production', 'completed', 'closed'],
        planning_job_id__isnull=False,
    ).values_list('planning_job_id', flat=True)
    PlanningJob.objects.filter(
        id__in=released_planning_job_ids,
    ).exclude(planning_stage='planning_done').update(planning_stage='planning_done')


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0025_planningjob_master_sync_request'),
    ]

    operations = [
        migrations.RunPython(migrate_planning_stage_values, migrations.RunPython.noop),
        migrations.AlterField(
            model_name='planningjob',
            name='planning_stage',
            field=models.CharField(
                blank=True,
                choices=[
                    ('', 'Not Set'),
                    ('jc_ready', 'JC Ready'),
                    ('new_plate_making', 'New Plate Making'),
                    ('repeat_plate_making', 'Repeat Plate Making'),
                    ('planning_done', 'Planning Done'),
                ],
                default='',
                max_length=40,
            ),
        ),
    ]
