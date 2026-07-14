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
