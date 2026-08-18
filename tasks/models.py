from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone

REMIND_FROM_ASSIGNMENT = 'assignment'
REMIND_FROM_OVERDUE = 'overdue'
REMIND_FROM_CHOICES = [
    (REMIND_FROM_ASSIGNMENT, 'From the moment the task is assigned'),
    (REMIND_FROM_OVERDUE, 'Only once the task is overdue (past due date)'),
]


class Team(models.Model):
    name = models.CharField(max_length=100, unique=True)
    description = models.TextField(blank=True)
    members = models.ManyToManyField(User, related_name='erp_teams')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class Task(models.Model):
    PRIORITY_CHOICES = [
        ('low', 'Low'),
        ('medium', 'Medium'),
        ('high', 'High'),
        ('critical', 'Critical'),
    ]
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('in_progress', 'In Progress'),
        ('completed', 'Completed'),
        ('verified', 'Verified'),
        ('cancelled', 'Cancelled'),
    ]

    title = models.CharField(max_length=200)
    description = models.TextField()
    assignee = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='assigned_tasks'
    )
    assigned_team = models.ForeignKey(
        Team,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='assigned_tasks'
    )
    priority = models.CharField(max_length=20, choices=PRIORITY_CHOICES, default='medium')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    due_date = models.DateField()
    
    # Scoring
    score = models.IntegerField(
        null=True,
        blank=True,
        help_text="Automatically calculated task score (40 to 100), can be adjusted manually by managers."
    )
    score_remarks = models.TextField(blank=True)
    scored_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='graded_tasks'
    )
    
    completed_at = models.DateTimeField(null=True, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='created_tasks')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    # Optional per-task overrides for the automation in tasks/reminders.py +
    # tasks/emails.py. Blank means "use the global default" (Tasks -> Automation).
    reminder_interval_days = models.PositiveSmallIntegerField(
        null=True, blank=True,
        help_text="Override the global reminder interval for this task specifically. "
                   "Blank uses the default set in Tasks → Automation."
    )
    cc_emails = models.TextField(
        blank=True,
        help_text="CC on this task's assignment/reminder emails. Comma or newline separated."
    )
    bcc_emails = models.TextField(
        blank=True,
        help_text="BCC on this task's assignment/reminder emails. Comma or newline separated."
    )

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return self.title

    def calculate_auto_score(self):
        """
        Auto-score based on time consumed:
        - Completed on or before due date = 100
        - Completed after due date = Penalty of 10 points per day late, minimum 40.
        """
        if self.completed_at and self.due_date:
            comp_date = self.completed_at.date()
            if comp_date <= self.due_date:
                return 100
            else:
                delay_days = (comp_date - self.due_date).days
                return max(40, 100 - (delay_days * 10))
        return None

    def save(self, *args, **kwargs):
        # Auto-set completed_at if status changed to completed/verified and not set
        if self.status in ['completed', 'verified'] and not self.completed_at:
            self.completed_at = timezone.now()
            
        # Clear completed_at if status moved back
        elif self.status not in ['completed', 'verified']:
            self.completed_at = None
            self.score = None
            
        # Calculate auto-score if score is not set yet
        if self.status in ['completed', 'verified'] and self.score is None:
            self.score = self.calculate_auto_score()
            
        super().save(*args, **kwargs)


class TaskComment(models.Model):
    task = models.ForeignKey(Task, on_delete=models.CASCADE, related_name='comments')
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    comment = models.TextField()
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['created_at']

    def __str__(self):
        return f"Comment by {self.user.username} on {self.task.title}"


class TaskNotificationSettings(models.Model):
    """Singleton row (EmailSettings.get_solo() pattern) controlling task
    assignment/reminder emails, editable from Settings — no restart needed."""

    assignment_email_enabled = models.BooleanField(
        default=True,
        help_text="Send an email immediately when a task is assigned/reassigned to a user or team."
    )
    reminders_enabled = models.BooleanField(
        default=True,
        help_text="Send recurring reminder emails for pending/in-progress tasks."
    )
    reminder_interval_days = models.PositiveSmallIntegerField(
        default=3,
        help_text="Days between reminder emails for the same task."
    )
    remind_from = models.CharField(
        max_length=12, choices=REMIND_FROM_CHOICES, default=REMIND_FROM_ASSIGNMENT,
        help_text="Whether the reminder clock starts at assignment or only once the task is overdue."
    )
    updated_at = models.DateTimeField(auto_now=True)
    updated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)

    class Meta:
        verbose_name = 'Task Notification Settings'
        verbose_name_plural = 'Task Notification Settings'

    def __str__(self):
        return 'Task Notification Settings'

    @classmethod
    def get_solo(cls):
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj


class TaskNotificationLog(models.Model):
    """One row per assignment/reminder email actually sent (or attempted) for
    a task — the dedup anchor for the reminder scheduler, and the activity
    log shown on the Tasks -> Automation page. Renamed from TaskReminderLog:
    now covers assignment sends too, not just reminders."""

    KIND_ASSIGNMENT = 'ASSIGNMENT'
    KIND_REMINDER = 'REMINDER'
    KIND_CHOICES = [(KIND_ASSIGNMENT, 'Assignment'), (KIND_REMINDER, 'Reminder')]

    STATUS_SENT = 'SENT'
    STATUS_FAILED = 'FAILED'
    STATUS_CHOICES = [(STATUS_SENT, 'Sent'), (STATUS_FAILED, 'Failed')]

    task = models.ForeignKey(Task, on_delete=models.CASCADE, related_name='notification_logs')
    kind = models.CharField(max_length=10, choices=KIND_CHOICES, default=KIND_REMINDER)
    sent_at = models.DateTimeField(auto_now_add=True)
    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default=STATUS_SENT)
    recipients_to = models.TextField(blank=True)
    recipients_cc = models.TextField(blank=True)
    recipients_bcc = models.TextField(blank=True)
    error_message = models.TextField(blank=True)

    class Meta:
        verbose_name = 'Task Notification Log'
        verbose_name_plural = 'Task Notification Logs'
        ordering = ['-sent_at']
        indexes = [models.Index(fields=['task', '-sent_at'])]

    def __str__(self):
        return f'{self.task_id} {self.get_kind_display()} @ {self.sent_at:%Y-%m-%d %H:%M} ({self.status})'
