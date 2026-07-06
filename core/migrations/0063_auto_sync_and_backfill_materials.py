import re
from django.db import migrations

def _resolve_by_name(model_class, raw_value):
    raw_value = (raw_value or '').strip()
    if not raw_value:
        return None

    normalized_value = re.sub(r'[\s_-]+', ' ', raw_value).strip()
    exact_match = model_class.objects.filter(name__iexact=normalized_value).first()
    if exact_match:
        return exact_match

    startswith_matches = model_class.objects.filter(name__istartswith=normalized_value)
    if startswith_matches.count() == 1:
        return startswith_matches.first()

    contains_matches = model_class.objects.filter(name__icontains=normalized_value)
    if contains_matches.count() == 1:
        return contains_matches.first()

    return model_class.objects.filter(name__iexact=raw_value).first()

def sync_and_backfill_materials(apps, schema_editor):
    PlanningJob = apps.get_model('planning', 'PlanningJob')
    SkuRecipe = apps.get_model('planning', 'SkuRecipe')
    JobCard = apps.get_model('core', 'JobCard')
    Material = apps.get_model('core', 'Material')

    # Gather unique material names
    names = set()
    for mat in PlanningJob.objects.exclude(material='').values_list('material', flat=True).distinct():
        cleaned = (mat or '').strip()
        if cleaned:
            names.add(cleaned)

    for mat in SkuRecipe.objects.exclude(material='').values_list('material', flat=True).distinct():
        cleaned = (mat or '').strip()
        if cleaned:
            names.add(cleaned)

    # Create Material master records
    for name in sorted(names):
        Material.objects.get_or_create(name=name)

    # Backfill JobCards
    for job_card in JobCard.objects.all():
        raw_mat = ''
        if job_card.planning_job:
            raw_mat = job_card.planning_job.material or ''
            if not raw_mat and job_card.planning_job.sku_recipe:
                raw_mat = job_card.planning_job.sku_recipe.material or ''

        if raw_mat:
            matched_material = _resolve_by_name(Material, raw_mat)
            if matched_material:
                job_card.material = matched_material
                job_card.save(update_fields=['material'])


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0062_fix_notification_links'),
    ]

    operations = [
        migrations.RunPython(sync_and_backfill_materials),
    ]
