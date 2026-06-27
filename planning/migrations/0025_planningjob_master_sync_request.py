from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('planning', '0024_planningjob_planning_stage_changed_by'),
    ]

    operations = [
        migrations.AddField(
            model_name='planningjob',
            name='master_sync_requested',
            field=models.BooleanField(default=False),
        ),
        migrations.AddField(
            model_name='planningjob',
            name='master_sync_reason',
            field=models.TextField(blank=True),
        ),
        migrations.AddField(
            model_name='planningjob',
            name='master_sync_requested_by',
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='planning_jobs_master_sync_requested',
                to=settings.AUTH_USER_MODEL,
            ),
        ),
        migrations.AddField(
            model_name='planningjob',
            name='master_sync_requested_at',
            field=models.DateTimeField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='planningjob',
            name='master_sync_applied_by',
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='planning_jobs_master_sync_applied',
                to=settings.AUTH_USER_MODEL,
            ),
        ),
        migrations.AddField(
            model_name='planningjob',
            name='master_sync_applied_at',
            field=models.DateTimeField(blank=True, null=True),
        ),
    ]
