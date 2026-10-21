import datetime
from io import StringIO

from django.contrib.auth import get_user_model
from django.core.management import call_command
from django.test import TestCase, override_settings
from django.utils import timezone

from bot.models import BotAutomation, BotExecution
from core.data_retention import BODY_PLACEHOLDER, prune_old_data
from core.models import Notification


class DataRetentionTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user('ret_user', password='pw12345678')
        self.now = timezone.now()

    def notification(self, days_old, read=False):
        row = Notification.objects.create(user=self.user, event_type='production.submitted', title='t', is_read=read)
        Notification.objects.filter(pk=row.pk).update(created_at=self.now - datetime.timedelta(days=days_old, hours=1))
        return row.pk

    def execution(self, days_old, body='<html>' + 'x' * 1000 + '</html>'):
        bot = BotAutomation.objects.first() or BotAutomation.objects.create(name='Retention bot', code='retention-bot')
        row = BotExecution.objects.create(bot=bot, status='SENT', rendered_subject='Daily report', rendered_body=body,
                                          recipients_to='a@b.c')
        BotExecution.objects.filter(pk=row.pk).update(started_at=self.now - datetime.timedelta(days=days_old, hours=1))
        return row.pk

    def test_notifications_older_than_the_window_are_deleted_read_or_not(self):
        old_unread, old_read = self.notification(45), self.notification(31, read=True)
        recent_unread, recent_read = self.notification(29), self.notification(1, read=True)
        result = prune_old_data(notification_days=30, bot_body_days=0, now=self.now)
        self.assertEqual(result['notifications'], 2)
        self.assertEqual(set(Notification.objects.values_list('pk', flat=True)), {recent_unread, recent_read})
        self.assertNotIn(old_unread, Notification.objects.values_list('pk', flat=True))

    def test_old_bot_email_bodies_are_cleared_but_the_record_is_kept(self):
        old, recent = self.execution(20), self.execution(3)
        result = prune_old_data(notification_days=0, bot_body_days=14, now=self.now)
        self.assertEqual(result['bot_bodies'], 1)
        old_row, recent_row = BotExecution.objects.get(pk=old), BotExecution.objects.get(pk=recent)
        self.assertEqual(old_row.rendered_body, BODY_PLACEHOLDER.format(days=14))
        self.assertEqual((old_row.rendered_subject, old_row.status, old_row.recipients_to), ('Daily report', 'SENT', 'a@b.c'))
        self.assertTrue(recent_row.rendered_body.startswith('<html>'))

    def test_running_it_twice_changes_nothing_the_second_time(self):
        self.execution(20)
        self.notification(40)
        prune_old_data(notification_days=30, bot_body_days=14, now=self.now)
        second = prune_old_data(notification_days=30, bot_body_days=14, now=self.now)
        self.assertEqual(second, {'notifications': 0, 'bot_bodies': 0})

    def test_zero_disables_a_rule(self):
        self.notification(400)
        self.execution(400)
        result = prune_old_data(notification_days=0, bot_body_days=0, now=self.now)
        self.assertEqual(result, {'notifications': 0, 'bot_bodies': 0})
        self.assertEqual(Notification.objects.count(), 1)

    def test_dry_run_reports_without_changing_anything(self):
        self.notification(40)
        self.execution(20)
        result = prune_old_data(notification_days=30, bot_body_days=14, now=self.now, dry_run=True)
        self.assertEqual(result, {'notifications': 1, 'bot_bodies': 1})
        self.assertEqual(Notification.objects.count(), 1)
        self.assertTrue(BotExecution.objects.get().rendered_body.startswith('<html>'))

    def test_a_large_backlog_is_removed_in_batches(self):
        import core.data_retention as module

        for _ in range(7):
            self.notification(60)
        original = module.BATCH
        module.BATCH = 3
        try:
            self.assertEqual(prune_old_data(notification_days=30, bot_body_days=0, now=self.now)['notifications'], 7)
        finally:
            module.BATCH = original
        self.assertEqual(Notification.objects.count(), 0)

    @override_settings(NOTIFICATION_RETENTION_DAYS=10, BOT_EMAIL_BODY_RETENTION_DAYS=5)
    def test_settings_supply_the_defaults(self):
        self.notification(11)
        self.notification(9)
        self.assertEqual(prune_old_data(now=self.now)['notifications'], 1)

    def test_management_command_reports_what_it_did(self):
        self.notification(60)
        out = StringIO()
        call_command('prune_old_data', '--dry-run', stdout=out)
        self.assertIn('Would remove 1 old notification', out.getvalue())
        out = StringIO()
        call_command('prune_old_data', stdout=out)
        self.assertIn('Removed 1 old notification', out.getvalue())
