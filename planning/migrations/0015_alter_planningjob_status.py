from django.db import migrations, models


def normalize_planning_job_statuses(apps, schema_editor):
    PlanningJob = apps.get_model('planning', 'PlanningJob')
    valid_statuses = {'draft', 'pending_qc', 'qc_approved', 'released', 'in_production', 'completed'}
    legacy_map = {
        'open': 'draft',
        'pending': 'draft',
        'reviewed': 'pending_qc',
        'approved': 'qc_approved',
        'closed': 'completed',
    }

    for job in PlanningJob.objects.all().iterator():
        raw_status = (job.status or '').strip().lower()
        normalized_status = legacy_map.get(raw_status, raw_status or 'draft')
        if normalized_status not in valid_statuses:
            normalized_status = 'draft'

        if raw_status != normalized_status:
            PlanningJob.objects.filter(pk=job.pk).update(status=normalized_status)


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0014_remove_machine_from_sku_recipe'),
    ]

    operations = [
        migrations.AlterField(
            model_name='planningjob',
            name='status',
            field=models.CharField(blank=True, choices=[('draft', 'Draft'), ('pending_qc', 'Pending QC'), ('qc_approved', 'QC Approved'), ('released', 'Released'), ('in_production', 'In Production'), ('completed', 'Completed')], default='draft', max_length=40),
        ),
        migrations.RunPython(normalize_planning_job_statuses, migrations.RunPython.noop),
    ]