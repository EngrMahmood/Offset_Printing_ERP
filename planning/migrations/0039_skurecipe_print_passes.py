from django.db import migrations, models


def backfill_skurecipe_print_passes(apps, schema_editor):
    SkuRecipe = apps.get_model('planning', 'SkuRecipe')
    PlanningJob = apps.get_model('planning', 'PlanningJob')

    for recipe in SkuRecipe.objects.all().iterator():
        if (recipe.job_process_type or 'print_and_pack') == 'cut_and_pack':
            continue
        if recipe.print_passes is not None:
            continue
        job = (
            PlanningJob.objects.filter(sku__iexact=recipe.sku, print_passes__isnull=False)
            .order_by('-updated_at', '-id')
            .first()
        )
        if not job:
            continue
        recipe.print_passes = job.print_passes
        recipe.save(update_fields=['print_passes'])

    open_statuses = ['draft', 'pending_qc']
    for job in PlanningJob.objects.filter(status__in=open_statuses).iterator():
        if (job.job_process_type or 'print_and_pack') == 'cut_and_pack':
            if job.print_passes is not None:
                job.print_passes = None
                job.save(update_fields=['print_passes'])
            continue
        recipe = (
            SkuRecipe.objects.filter(sku__iexact=job.sku, master_data_status='approved')
            .order_by('id')
            .first()
        )
        if not recipe or not recipe.print_passes:
            continue
        master_passes = int(recipe.print_passes)
        if job.print_passes != master_passes:
            job.print_passes = master_passes
            job.save(update_fields=['print_passes'])


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0038_sync_plate_making_stage_with_repeat_flag'),
    ]

    operations = [
        migrations.AddField(
            model_name='skurecipe',
            name='print_passes',
            field=models.PositiveSmallIntegerField(
                blank=True,
                help_text='Number of press passes (1, 2, or 3) for Print + Pack SKUs.',
                null=True,
            ),
        ),
        migrations.RunPython(backfill_skurecipe_print_passes, migrations.RunPython.noop),
    ]
