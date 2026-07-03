from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from core.models import Notification
from core.notifications import notify_roles, notify_users

User = get_user_model()


class NotificationSystemTests(TestCase):
    def setUp(self):
        self.designer = User.objects.create_user(username='notify_designer', password='pass')
        profile = self.designer.profile
        profile.role = 'graphics_designer'
        profile.save(update_fields=['role'])

        self.planner = User.objects.create_user(username='notify_planner', password='pass')
        planner_profile = self.planner.profile
        planner_profile.role = 'planner'
        planner_profile.save(update_fields=['role'])

        self.actor = User.objects.create_user(username='notify_actor', password='pass')

    def test_notify_roles_creates_rows_and_skips_actor(self):
        created = notify_roles(
            ('graphics_designer', 'planner'),
            event_type='test.event',
            title='Test title',
            message='Hello',
            link='/planning/',
            actor=self.planner,
        )
        self.assertEqual(created, 1)
        self.assertTrue(Notification.objects.filter(user=self.designer, title='Test title').exists())
        self.assertFalse(Notification.objects.filter(user=self.planner).exists())

    def test_notification_api_list_and_mark_read(self):
        notify_users(
            [self.designer],
            event_type='test.event',
            title='API item',
            message='Body',
            link='/planning/',
            actor=self.actor,
        )
        self.client.login(username='notify_designer', password='pass')
        response = self.client.get(reverse('notification_list'))
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data['ok'])
        self.assertEqual(data['unread_count'], 1)
        self.assertEqual(data['items'][0]['title'], 'API item')

        item_id = data['items'][0]['id']
        mark = self.client.post(reverse('notification_mark_read', args=[item_id]))
        self.assertEqual(mark.status_code, 200)
        self.assertEqual(mark.json()['unread_count'], 0)
        self.assertTrue(Notification.objects.get(pk=item_id).is_read)
