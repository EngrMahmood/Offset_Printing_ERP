from django.core import mail
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from django.utils import timezone
from django.contrib.auth.models import User
from datetime import timedelta
from .models import Team, Task, TaskAttachment

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


class TaskVarianceTests(TestCase):
    """Task.variance_days / variance_kind drive the Tasks dashboard's
    Completed / Due Days audit columns."""

    def setUp(self):
        self.creator = User.objects.create_user(username='variance_creator', password='pwd')
        self.user = User.objects.create_user(username='variance_employee', password='pwd')

    def test_completed_late_reports_positive_days(self):
        due_date = timezone.now().date() - timedelta(days=10)
        task = Task.objects.create(
            title="Late", description="x", assignee=self.user, due_date=due_date,
            created_by=self.creator, status='pending',
        )
        task.status = 'completed'
        task.save()
        Task.objects.filter(pk=task.pk).update(completed_at=timezone.now() - timedelta(days=5))
        task.refresh_from_db()

        self.assertEqual(task.variance_kind, 'late')
        self.assertEqual(task.variance_days, 5)

    def test_completed_early_reports_negative_direction(self):
        due_date = timezone.now().date() + timedelta(days=5)
        task = Task.objects.create(
            title="Early", description="x", assignee=self.user, due_date=due_date,
            created_by=self.creator, status='completed',
        )
        self.assertEqual(task.variance_kind, 'early')
        self.assertEqual(task.variance_days, 5)

    def test_completed_on_due_date_is_on_time(self):
        due_date = timezone.now().date()
        task = Task.objects.create(
            title="OnTime", description="x", assignee=self.user, due_date=due_date,
            created_by=self.creator, status='completed',
        )
        self.assertEqual(task.variance_kind, 'on_time')
        self.assertEqual(task.variance_days, 0)

    def test_pending_and_overdue_reports_overdue_kind(self):
        due_date = timezone.now().date() - timedelta(days=4)
        task = Task.objects.create(
            title="StillOpen", description="x", assignee=self.user, due_date=due_date,
            created_by=self.creator, status='in_progress',
        )
        self.assertEqual(task.variance_kind, 'overdue')
        self.assertEqual(task.variance_days, 4)

    def test_pending_and_not_yet_due_reports_none(self):
        due_date = timezone.now().date() + timedelta(days=4)
        task = Task.objects.create(
            title="NotDueYet", description="x", assignee=self.user, due_date=due_date,
            created_by=self.creator, status='pending',
        )
        self.assertEqual(task.variance_kind, 'none')
        self.assertIsNone(task.variance_days)


class TaskDashboardAuditColumnsTests(TestCase):
    """Regression for the dashboard request to surface Completed date, Due
    Days variance, and who assigned the task, for a complete audit trail."""

    def setUp(self):
        self.creator = User.objects.create_user(username='dash_creator', password='pwd')
        self.user = User.objects.create_user(username='dash_employee', password='pwd')
        self.client.force_login(self.creator)

    def test_dashboard_shows_assigned_by_completed_and_due_days(self):
        due_date = timezone.now().date() - timedelta(days=3)
        task = Task.objects.create(
            title="Audit Trail Task", description="x", assignee=self.user, due_date=due_date,
            created_by=self.creator, status='pending',
        )

        response = self.client.get('/tasks/')

        self.assertEqual(response.status_code, 200)
        html = response.content.decode()
        self.assertIn('Due Days', html)
        self.assertIn('Completed', html)
        self.assertIn(f'by {self.creator.username}', html)
        self.assertIn('overdue', html)


class TaskAttachmentTests(TestCase):
    """Pasted screenshots / manually picked files on the create & edit
    forms — the Outlook-style paste-to-attach feature."""

    def setUp(self):
        self.creator = User.objects.create_user(username='attach_creator', password='pwd')
        self.other_user = User.objects.create_user(username='attach_other', password='pwd')
        self.employee = User.objects.create_user(
            username='attach_employee', password='pwd', email='attach_employee@example.com',
        )

    def _image_file(self, name='pasted-screenshot-123.png'):
        # Minimal valid 1x1 PNG.
        png_bytes = (
            b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01'
            b'\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\xcf\xc0'
            b'\x00\x00\x03\x01\x01\x00\x18\xdd\x8d\xb0\x00\x00\x00\x00IEND\xaeB`\x82'
        )
        return SimpleUploadedFile(name, png_bytes, content_type='image/png')

    def test_create_task_saves_uploaded_attachment(self):
        self.client.force_login(self.creator)
        response = self.client.post('/tasks/create/', {
            'title': 'Task with screenshot',
            'description': 'See attached',
            'assignee': self.employee.pk,
            'priority': 'medium',
            'due_date': (timezone.now().date() + timedelta(days=2)).isoformat(),
            'attachments': [self._image_file()],
        })

        self.assertEqual(response.status_code, 302)
        task = Task.objects.get(title='Task with screenshot')
        self.assertEqual(task.attachments.count(), 1)
        attachment = task.attachments.first()
        self.assertTrue(attachment.is_image)
        self.assertEqual(attachment.uploaded_by, self.creator)

    def test_assignment_email_includes_the_pasted_attachment(self):
        """Regression: the assignment email used to fire from Task's post_save
        signal *before* the view got a chance to save the uploaded/pasted
        attachment rows, so task.attachments.all() was still empty by the
        time the email was built — attachments silently never made it into
        the email. create_task now wraps the task save + attachment save in
        one atomic() block, and the signal defers sending via
        transaction.on_commit() so it only runs once both exist."""
        self.client.force_login(self.creator)

        with self.captureOnCommitCallbacks(execute=True):
            response = self.client.post('/tasks/create/', {
                'title': 'Task with emailed screenshot',
                'description': 'See attached',
                'assignee': self.employee.pk,
                'priority': 'medium',
                'due_date': (timezone.now().date() + timedelta(days=2)).isoformat(),
                'attachments': [self._image_file('email-shot.png')],
            })

        self.assertEqual(response.status_code, 302)
        self.assertEqual(len(mail.outbox), 1)
        sent = mail.outbox[0]
        self.assertEqual(len(sent.attachments), 1)
        self.assertEqual(sent.attachments[0][0], 'email-shot.png')

    def test_edit_task_adds_additional_attachment(self):
        task = Task.objects.create(
            title="Existing", description="x", assignee=self.employee,
            due_date=timezone.now().date() + timedelta(days=2), created_by=self.creator,
        )
        self.client.force_login(self.creator)

        response = self.client.post(f'/tasks/{task.pk}/edit/', {
            'title': task.title,
            'description': task.description,
            'assignee': self.employee.pk,
            'priority': 'medium',
            'due_date': task.due_date.isoformat(),
            'attachments': [self._image_file('extra.png')],
        })

        self.assertEqual(response.status_code, 302)
        self.assertEqual(task.attachments.count(), 1)

    def test_only_creator_or_manager_can_delete_attachment(self):
        task = Task.objects.create(
            title="Existing", description="x", assignee=self.employee,
            due_date=timezone.now().date() + timedelta(days=2), created_by=self.creator,
        )
        attachment = TaskAttachment.objects.create(
            task=task, file=self._image_file(), original_filename='shot.png', uploaded_by=self.creator,
        )

        self.client.force_login(self.other_user)
        response = self.client.post(f'/tasks/attachments/{attachment.pk}/delete/')
        self.assertEqual(response.status_code, 302)
        self.assertTrue(TaskAttachment.objects.filter(pk=attachment.pk).exists())

        self.client.force_login(self.creator)
        response = self.client.post(f'/tasks/attachments/{attachment.pk}/delete/')
        self.assertEqual(response.status_code, 302)
        self.assertFalse(TaskAttachment.objects.filter(pk=attachment.pk).exists())


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
