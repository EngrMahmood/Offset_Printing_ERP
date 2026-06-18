from django.db import migrations, models


def copy_legacy_to_new_columns(apps, schema_editor):
    Production = apps.get_model('core', 'Production')
    for production in Production.objects.all():
        if hasattr(production, 'setup_time') and production.setup_time is not None:
            production.make_ready_time = production.setup_time
        if hasattr(production, 'downtime') and production.downtime is not None:
            production.downtime_minutes = production.downtime
        production.save(update_fields=['make_ready_time', 'downtime_minutes'])


def noop(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0041_add_production_intermediate_pass'),
    ]

    operations = [
        migrations.AddField(
            model_name='production',
            name='make_ready_time',
            field=models.FloatField(default=0, help_text='in minutes'),
        ),
        migrations.AddField(
            model_name='production',
            name='downtime_minutes',
            field=models.FloatField(default=0, help_text='in minutes'),
        ),
        migrations.RunPython(copy_legacy_to_new_columns, noop),
    ]
