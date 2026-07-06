from django.db import migrations, models

LEGACY_BULK_CORE_FIELDS = (
    'material',
    'color_spec',
    'ups',
    'print_sheet_size',
    'job_name',
)


def _recipe_is_bulk_like(recipe):
    filled = sum(
        1
        for field in LEGACY_BULK_CORE_FIELDS
        if str(getattr(recipe, field, '') or '').strip()
    )
    return filled >= 3


def backfill_legacy_produced(apps, schema_editor):
    SkuRecipe = apps.get_model('planning', 'SkuRecipe')
    to_update = []
    for recipe in SkuRecipe.objects.iterator():
        if _recipe_is_bulk_like(recipe):
            recipe.legacy_produced = True
            to_update.append(recipe)
        if len(to_update) >= 500:
            SkuRecipe.objects.bulk_update(to_update, ['legacy_produced'])
            to_update = []
    if to_update:
        SkuRecipe.objects.bulk_update(to_update, ['legacy_produced'])


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0036_skurecipe_job_process_type'),
    ]

    operations = [
        migrations.AddField(
            model_name='skurecipe',
            name='legacy_produced',
            field=models.BooleanField(
                default=False,
                help_text='True when SKU master came from Google Sheet / bulk upload and was produced before ERP go-live.',
            ),
        ),
        migrations.RunPython(backfill_legacy_produced, migrations.RunPython.noop),
    ]
