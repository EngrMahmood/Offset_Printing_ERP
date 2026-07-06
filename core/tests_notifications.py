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


from core.models import NotificationEvent, NotificationRule, WorkflowTransition, NotificationRuleAuditLog, Department
from core.notifications import notify_event

class RuleBasedNotificationTests(TestCase):
    def setUp(self):
        self.qc_user = User.objects.create_user(username='test_qc', password='pass')
        self.qc_user.profile.role = 'qc'
        self.qc_user.profile.save()

        self.dept = Department.objects.create(name='Graphics Dept')
        self.graphics_user = User.objects.create_user(username='test_graphics', password='pass')
        self.graphics_user.profile.role = 'graphics_designer'
        self.graphics_user.profile.department = self.dept
        self.graphics_user.profile.save()

        self.manager_user = User.objects.create_user(username='test_manager', password='pass')
        self.manager_user.profile.role = 'manager'
        self.manager_user.profile.save()

        self.creator_user = User.objects.create_user(username='test_creator', password='pass')
        self.creator_user.profile.manager = self.manager_user
        self.creator_user.profile.save()

        # Create a mock instance
        from planning.models import SkuRecipe
        self.recipe = SkuRecipe.objects.create(
            sku='TEST-SKU-999',
            job_name='Super Job Name',
            created_by=self.creator_user
        )

        # Event
        self.event = NotificationEvent.objects.create(
            code='test.custom_event',
            name='Test Custom Event',
            title_template='SKU Review: {{ instance.sku }}',
            message_template='Submitted by {{ actor.username }} for {{ instance.job_name }}',
            link_template='/sku/{{ instance.pk }}/'
        )

    def test_routing_by_role(self):
        rule = NotificationRule.objects.create(
            event=self.event,
            recipient_type='role',
            role='qc',
            in_app_enabled=True
        )
        created = notify_event('test.custom_event', instance=self.recipe, actor=self.creator_user)
        self.assertEqual(created, 1)
        self.assertTrue(Notification.objects.filter(user=self.qc_user, title='SKU Review: TEST-SKU-999').exists())
        notif = Notification.objects.get(user=self.qc_user)
        self.assertEqual(notif.message, 'Submitted by test_creator for Super Job Name')
        self.assertEqual(notif.link, '/sku/%d/' % self.recipe.pk)

    def test_routing_by_department(self):
        rule = NotificationRule.objects.create(
            event=self.event,
            recipient_type='department',
            department=self.dept,
            in_app_enabled=True
        )
        created = notify_event('test.custom_event', instance=self.recipe, actor=self.creator_user)
        self.assertEqual(created, 1)
        self.assertTrue(Notification.objects.filter(user=self.graphics_user).exists())

    def test_routing_to_creator_and_manager(self):
        rule = NotificationRule.objects.create(
            event=self.event,
            recipient_type='role',
            role='qc',
            send_to_creator=True,
            send_to_manager=True,
            in_app_enabled=True
        )
        created = notify_event('test.custom_event', instance=self.recipe, actor=self.graphics_user)
        self.assertEqual(created, 3)
        self.assertTrue(Notification.objects.filter(user=self.creator_user).exists())
        self.assertTrue(Notification.objects.filter(user=self.manager_user).exists())
        self.assertTrue(Notification.objects.filter(user=self.qc_user).exists())

    def test_workflow_transition_routing(self):
        WorkflowTransition.objects.create(
            module='SkuRecipe',
            current_stage='draft',
            action='Submit',
            next_stage='pending_review',
            notify_role='qc'
        )
        rule = NotificationRule.objects.create(
            event=self.event,
            recipient_type='next_stage',
            in_app_enabled=True
        )
        self.recipe.master_data_status = 'draft'
        self.recipe.save()
        created = notify_event('test.custom_event', instance=self.recipe, actor=self.creator_user)
        self.assertEqual(created, 1)
        self.assertTrue(Notification.objects.filter(user=self.qc_user).exists())

    def test_audit_logging_mechanism(self):
        from core.notifications import log_rule_change
        rule = NotificationRule.objects.create(
            event=self.event,
            recipient_type='role',
            role='qc',
            in_app_enabled=True
        )
        log_rule_change(self.manager_user, rule, 'create')
        self.assertEqual(NotificationRuleAuditLog.objects.count(), 1)
        log = NotificationRuleAuditLog.objects.first()
        self.assertEqual(log.action, 'create')
        self.assertEqual(log.changed_by, self.manager_user)
        self.assertEqual(log.new_values['recipient_type'], 'role')


class NotificationSettingsViewsTests(TestCase):
    def setUp(self):
        self.admin_user = User.objects.create_user(username='admin_user', password='pass')
        self.admin_user.profile.role = 'admin'
        self.admin_user.profile.save()

        self.normal_user = User.objects.create_user(username='normal_user', password='pass')
        self.normal_user.profile.role = 'planner'
        self.normal_user.profile.save()

        self.event = NotificationEvent.objects.create(
            code='test.view_event',
            name='Test View Event',
            is_active=True
        )

    def test_forgot_password_views(self):
        response = self.client.get(reverse('forgot_password'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Reset Password")

        response = self.client.post(reverse('forgot_password'), {'username_or_email': 'normal_user'})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Password reset request logged for")
        self.assertContains(response, "normal_user")

        from core.models import PasswordResetRequest
        self.assertEqual(PasswordResetRequest.objects.filter(username_or_email='normal_user').count(), 1)

    def test_settings_home_access(self):
        response = self.client.get(reverse('notification_settings_home'))
        self.assertEqual(response.status_code, 302)

        self.client.login(username='normal_user', password='pass')
        response = self.client.get(reverse('notification_settings_home'))
        self.assertEqual(response.status_code, 302)
        self.client.logout()

        self.client.login(username='admin_user', password='pass')
        response = self.client.get(reverse('notification_settings_home'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "System Settings")
        self.assertContains(response, "Active Notification Rules")

    def test_rule_add_and_delete_views(self):
        self.client.login(username='admin_user', password='pass')
        
        initial_count = NotificationRule.objects.count()
        initial_audit_count = NotificationRuleAuditLog.objects.count()

        post_data = {
            'event': self.event.id,
            'recipient_type': 'role',
            'role': 'qc',
            'enabled': 'on',
            'exclude_actor': 'on',
            'in_app_enabled': 'on',
            'priority': 'high'
        }
        response = self.client.post(reverse('notification_rule_add'), post_data)
        self.assertEqual(response.status_code, 302)
        
        self.assertEqual(NotificationRule.objects.count(), initial_count + 1)
        rule = NotificationRule.objects.order_by('-id').first()
        self.assertEqual(rule.role, 'qc')
        self.assertEqual(rule.priority, 'high')

        self.assertEqual(NotificationRuleAuditLog.objects.count(), initial_audit_count + 1)
        
        response = self.client.post(reverse('notification_rule_delete', args=[rule.id]))
        self.assertEqual(response.status_code, 302)
        self.assertEqual(NotificationRule.objects.count(), initial_count)

    def test_workflow_transition_add_and_delete_views(self):
        self.client.login(username='admin_user', password='pass')

        initial_count = WorkflowTransition.objects.count()

        post_data = {
            'module': 'JobCard',
            'current_stage': 'draft',
            'action': 'Release',
            'next_stage': 'released',
            'notify_role': 'production'
        }
        response = self.client.post(reverse('workflow_transition_add'), post_data)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(WorkflowTransition.objects.count(), initial_count + 1)
        
        transition = WorkflowTransition.objects.order_by('-id').first()
        self.assertEqual(transition.module, 'JobCard')

        response = self.client.post(reverse('workflow_transition_delete', args=[transition.id]))
        self.assertEqual(response.status_code, 302)
        self.assertEqual(WorkflowTransition.objects.count(), initial_count)
