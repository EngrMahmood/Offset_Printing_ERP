from django.test import TestCase
from django.contrib.auth.models import User
from django.utils import timezone
from datetime import timedelta
from .models import Team, Task

class TaskScoringTests(TestCase):

    def setUp(self):
        self.creator = User.objects.create_user(username='creator', password='pwd')
        self.user = User.objects.create_user(username='employee', password='pwd')
        self.team = Team.objects.create(name='QC Team', description='Quality Control')
        self.team.members.add(self.user)

    def test_on_time_task_score(self):
        # Due tomorrow, completed today
        due_date = timezone.now().date() + timedelta(days=1)
        task = Task.objects.create(
            title="On Time Task",
            description="Testing 100 score",
            assignee=self.user,
            due_date=due_date,
            created_by=self.creator,
            status='pending'
        )
        # Complete task
        task.status = 'completed'
        task.save()
        
        self.assertEqual(task.score, 100)
        self.assertIsNotNone(task.completed_at)

    def test_delayed_task_score_penalty(self):
        # Due 3 days ago, completed today
        due_date = timezone.now().date() - timedelta(days=3)
        task = Task.objects.create(
            title="Late Task",
            description="Testing penalty score",
            assignee=self.user,
            due_date=due_date,
            created_by=self.creator,
            status='pending'
        )
        task.status = 'completed'
        task.save()
        
        # 100 - (3 * 10) = 70
        self.assertEqual(task.score, 70)

    def test_minimum_delayed_score_floor(self):
        # Due 10 days ago, completed today
        due_date = timezone.now().date() - timedelta(days=10)
        task = Task.objects.create(
            title="Very Late Task",
            description="Testing score floor",
            assignee=self.user,
            due_date=due_date,
            created_by=self.creator,
            status='pending'
        )
        task.status = 'completed'
        task.save()
        
        # 100 - (10 * 10) = 0, but floor is 40
        self.assertEqual(task.score, 40)

    def test_team_task_assignment(self):
        due_date = timezone.now().date() + timedelta(days=2)
        task = Task.objects.create(
            title="Team Task",
            description="Assigned to QC Team",
            assigned_team=self.team,
            due_date=due_date,
            created_by=self.creator,
            status='pending'
        )
        
        self.assertEqual(task.assigned_team.name, "QC Team")
        self.assertIn(self.user, task.assigned_team.members.all())


class TeamEditTests(TestCase):
    """edit_team view — teams_list's own POST handler only ever creates a
    new Team, with no way to update one after creation."""

    def setUp(self):
        from django.urls import reverse
        from core.models import Permission, Role, UserPermissionOverride, UserProfile

        self.reverse = reverse
        Role.objects.get_or_create(slug='operator', defaults={'display_name': 'Operator'})

        self.manager = User.objects.create_user(username='team_manager', password='pass12345')
        profile, _ = UserProfile.objects.get_or_create(user=self.manager, defaults={'role': 'operator'})
        permission, _ = Permission.objects.get_or_create(
            code='action.manage_tasks', defaults={'name': 'Manage Tasks'},
        )
        UserPermissionOverride.objects.get_or_create(
            user=self.manager, permission=permission, defaults={'granted': True},
        )

        self.plain_user = User.objects.create_user(username='plain_employee', password='pass12345')
        self.member_a = User.objects.create_user(username='member_a', password='pass12345')
        self.member_b = User.objects.create_user(username='member_b', password='pass12345')

        self.team = Team.objects.create(name='Original Team', description='Original description')
        self.team.members.add(self.member_a)

    def test_manager_can_edit_team_name_description_and_members(self):
        self.client.login(username='team_manager', password='pass12345')

        response = self.client.post(self.reverse('tasks:team_edit', args=[self.team.pk]), {
            'name': 'Renamed Team',
            'description': 'Updated description',
            'members': [self.member_b.id],
        })

        self.assertRedirects(response, self.reverse('tasks:teams'))
        self.team.refresh_from_db()
        self.assertEqual(self.team.name, 'Renamed Team')
        self.assertEqual(self.team.description, 'Updated description')
        self.assertEqual(list(self.team.members.all()), [self.member_b])

    def test_get_edit_team_prefills_form(self):
        self.client.login(username='team_manager', password='pass12345')

        response = self.client.get(self.reverse('tasks:team_edit', args=[self.team.pk]))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Original Team')

    def test_non_manager_cannot_edit_team(self):
        self.client.login(username='plain_employee', password='pass12345')

        response = self.client.post(self.reverse('tasks:team_edit', args=[self.team.pk]), {
            'name': 'Hijacked Name', 'description': '', 'members': [],
        })

        self.assertRedirects(response, self.reverse('tasks:teams'))
        self.team.refresh_from_db()
        self.assertEqual(self.team.name, 'Original Team')
