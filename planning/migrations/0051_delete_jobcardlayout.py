from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('planning', '0050_alter_mergegroup_id_alter_mergegroupitem_id'),
    ]

    operations = [
        migrations.DeleteModel(
            name='JobCardLayout',
        ),
    ]
