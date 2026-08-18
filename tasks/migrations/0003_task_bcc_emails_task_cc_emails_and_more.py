# Hand-edited from the auto-generated version: makemigrations proposed a
# DeleteModel(TaskReminderLog) + CreateModel(TaskNotificationLog) pair, which
# would drop any reminder-log rows already written in production by the
# scheduler since deploy. Replaced with RenameModel + RenameField so existing
# data survives.
import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('tasks', '0002_tasknotificationsettings_taskreminderlog'),
    ]

    operations = [
        migrations.AddField(
            model_name='task',
            name='bcc_emails',
            field=models.TextField(blank=True, help_text="BCC on this task's assignment/reminder emails. Comma or newline separated."),
        ),
        migrations.AddField(
            model_name='task',
            name='cc_emails',
            field=models.TextField(blank=True, help_text="CC on this task's assignment/reminder emails. Comma or newline separated."),
        ),
        migrations.AddField(
            model_name='task',
            name='reminder_interval_days',
            field=models.PositiveSmallIntegerField(blank=True, help_text='Override the global reminder interval for this task specifically. Blank uses the default set in Tasks → Automation.', null=True),
        ),
        migrations.RenameModel(
            old_name='TaskReminderLog',
            new_name='TaskNotificationLog',
        ),
        migrations.RenameField(
            model_name='tasknotificationlog',
            old_name='recipients',
            new_name='recipients_to',
        ),
        migrations.AlterField(
            model_name='tasknotificationlog',
            name='task',
            field=models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='notification_logs', to='tasks.task'),
        ),
        migrations.AddField(
            model_name='tasknotificationlog',
            name='kind',
            field=models.CharField(choices=[('ASSIGNMENT', 'Assignment'), ('REMINDER', 'Reminder')], default='REMINDER', max_length=10),
        ),
        migrations.AddField(
            model_name='tasknotificationlog',
            name='recipients_cc',
            field=models.TextField(blank=True),
        ),
        migrations.AddField(
            model_name='tasknotificationlog',
            name='recipients_bcc',
            field=models.TextField(blank=True),
        ),
        migrations.AlterModelOptions(
            name='tasknotificationlog',
            options={'ordering': ['-sent_at'], 'verbose_name': 'Task Notification Log', 'verbose_name_plural': 'Task Notification Logs'},
        ),
        migrations.RenameIndex(
            model_name='tasknotificationlog',
            old_name='tasks_taskr_task_id_ab9a58_idx',
            new_name='tasks_taskn_task_id_dfdd6b_idx',
        ),
    ]
