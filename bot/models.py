"""Bot / Automation — scheduled report emails, configured entirely from the ERP UI.

A bot is a *configuration row*, not code: it names a report (by its
`reports.report_registry` slug), the filters to run it with, when to run, who
receives it, and the subject/body templates. Adding automation #2..#N is a new
BotAutomation row, not a new module.
"""
import datetime

from django.conf import settings
from django.db import models


FREQUENCY_DAILY = 'DAILY'
FREQUENCY_WEEKLY = 'WEEKLY'
FREQUENCY_MONTHLY = 'MONTHLY'

FREQUENCY_CHOICES = [
    (FREQUENCY_DAILY, 'Daily'),
    (FREQUENCY_WEEKLY, 'Weekly'),
    (FREQUENCY_MONTHLY, 'Monthly'),
]

ATTACHMENT_FORMAT_CHOICES = [
    ('xlsx', 'Excel (.xlsx)'),
    ('csv', 'CSV (.csv)'),
    ('pdf', 'PDF (.pdf)'),
]

# Mirrors the Period dropdown the reports themselves offer
# (reports/templates/reports/pending-work.html) and the presets understood by
# reports.services._parse_period_filter. Blank means "send whatever the report
# would show by default" — today that is the current month.
PERIOD_DEFAULT = ''
PERIOD_CUSTOM = 'custom'

PERIOD_CHOICES = [
    (PERIOD_DEFAULT, 'Report default'),
    ('today', 'Today'),
    ('week', 'This Week'),
    ('month', 'This Month'),
    (PERIOD_CUSTOM, 'Custom Range'),
    ('all', 'All Time'),
]

WEEKDAY_CHOICES = [
    ('0', 'Monday'),
    ('1', 'Tuesday'),
    ('2', 'Wednesday'),
    ('3', 'Thursday'),
    ('4', 'Friday'),
    ('5', 'Saturday'),
    ('6', 'Sunday'),
]

TRIGGER_AUTO = 'AUTO'
TRIGGER_MANUAL = 'MANUAL'
TRIGGER_TEST = 'TEST'

TRIGGER_CHOICES = [
    (TRIGGER_AUTO, 'Scheduled'),
    (TRIGGER_MANUAL, 'Run Now'),
    (TRIGGER_TEST, 'Test Send'),
]

STATUS_PENDING = 'PENDING'
STATUS_SENT = 'SENT'
STATUS_SKIPPED = 'SKIPPED'
STATUS_FAILED = 'FAILED'

EXECUTION_STATUS_CHOICES = [
    (STATUS_PENDING, 'Running'),
    (STATUS_SENT, 'Sent'),
    (STATUS_SKIPPED, 'Skipped (no records)'),
    (STATUS_FAILED, 'Failed'),
]


DEFAULT_SUBJECT_TEMPLATE = 'Pending Work - Required for Production Release - {{date}}'

DEFAULT_BODY_TEMPLATE = """<p>Dear Planning Team,</p>

<p>Please find below the pending work that is required to be released
for production.</p>

{{report_table}}

<p>Kindly review and release the pending jobs as per priority.</p>

<p>Regards,<br>
Production Printing</p>
"""

# Seed copy for the per-stage production backlog bots (printing / packing /
# dispatch). STAGE is substituted once at seed time rather than left as a
# template variable, so an admin can reword one bot's email from the UI without
# touching the other two.
_STAGE_BODY_TEMPLATE = """<p>Dear Team,</p>

<p>Below is the pending STAGE work as of {{date}} ({{filters_summary}}).</p>

{{report_table}}

<p>Total pending job cards: {{total_records}}</p>

<p>Kindly clear the backlog as per priority.</p>

<p>Regards,<br>
Production Printing</p>
"""


def stage_body_template(stage_label):
    return _STAGE_BODY_TEMPLATE.replace('STAGE', stage_label)


def _split_addresses(raw):
    """Split a comma/newline/semicolon separated address blob into a clean list."""
    if not raw:
        return []
    parts = (raw or '').replace(';', ',').replace('\n', ',').replace('\r', ',').split(',')
    return [part.strip() for part in parts if part.strip()]


class BotAutomation(models.Model):
    """One scheduled automation. Everything a business user needs to change
    lives here; the report query itself stays in the reports app."""

    code = models.CharField(
        max_length=80,
        unique=True,
        help_text="Stable identifier used for seeding/lookup, e.g. PENDING_PRODUCTION_RELEASE.",
    )
    name = models.CharField(max_length=150)
    description = models.TextField(blank=True)
    is_active = models.BooleanField(
        default=False,
        help_text="Only active bots are picked up by the scheduler. Preview and test-send first.",
    )

    # --- What to send -----------------------------------------------------
    report_slug = models.CharField(
        max_length=120,
        db_index=True,
        help_text="Slug of a report registered in reports.report_registry (e.g. pending-work).",
    )
    report_filters = models.JSONField(
        default=dict,
        blank=True,
        help_text='Filters passed to the report, e.g. {"stage": "not_released"}.',
    )
    report_period = models.CharField(
        max_length=10,
        choices=PERIOD_CHOICES,
        blank=True,
        default=PERIOD_DEFAULT,
        help_text="Date range the report covers, same presets as the report screen. "
                  "Blank leaves the report's own default (current month).",
    )
    report_date_from = models.DateField(
        null=True,
        blank=True,
        help_text="Start of the range when the period is Custom Range.",
    )
    report_date_to = models.DateField(
        null=True,
        blank=True,
        help_text="End of the range when the period is Custom Range.",
    )

    # --- When to send -----------------------------------------------------
    frequency = models.CharField(max_length=10, choices=FREQUENCY_CHOICES, default=FREQUENCY_DAILY)
    send_time = models.TimeField(default=datetime.time(8, 30))
    weekdays = models.CharField(
        max_length=20,
        blank=True,
        help_text="Comma-separated weekday numbers (0=Mon .. 6=Sun). Blank means every day.",
    )
    day_of_month = models.PositiveSmallIntegerField(
        null=True,
        blank=True,
        help_text="Day of month for MONTHLY frequency (1-31). Clamped to the last day of short months.",
    )
    start_date = models.DateField(null=True, blank=True, help_text="Optional: do not run before this date.")
    end_date = models.DateField(null=True, blank=True, help_text="Optional: do not run after this date.")

    # --- Who receives it --------------------------------------------------
    email_to = models.TextField(blank=True, help_text="Comma or newline separated addresses.")
    email_cc = models.TextField(blank=True)
    email_bcc = models.TextField(blank=True)
    recipient_roles = models.CharField(
        max_length=255,
        blank=True,
        help_text="Comma-separated role keys (e.g. planner,production_manager). "
                  "Resolved to live user emails at send time.",
    )

    # --- What it looks like -----------------------------------------------
    subject_template = models.CharField(max_length=255, default=DEFAULT_SUBJECT_TEMPLATE)
    body_template = models.TextField(default=DEFAULT_BODY_TEMPLATE)
    max_rows_in_body = models.PositiveIntegerField(
        default=100,
        help_text="Rows shown inline in the email body. The rest go in the attachment.",
    )

    attach_report = models.BooleanField(default=True)
    attachment_format = models.CharField(max_length=10, choices=ATTACHMENT_FORMAT_CHOICES, default='xlsx')
    send_when_empty = models.BooleanField(
        default=False,
        help_text="Send the email even when the report returns zero rows.",
    )

    # --- Reliability ------------------------------------------------------
    retry_count = models.PositiveSmallIntegerField(
        default=2,
        help_text="Extra attempts after a failed run, within the same day.",
    )
    retry_interval_minutes = models.PositiveIntegerField(
        default=15,
        help_text="Minimum wait before retrying a failed run.",
    )

    run_as = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='bots_run_as',
        help_text="Report runs with this user's permissions. Defaults to the first superuser.",
    )

    # --- Bookkeeping ------------------------------------------------------
    last_run_at = models.DateTimeField(null=True, blank=True)
    next_run_at = models.DateTimeField(null=True, blank=True, db_index=True)
    last_status = models.CharField(max_length=10, blank=True)

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='bots_created',
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'Bot Automation'
        verbose_name_plural = 'Bot Automations'
        ordering = ['name']

    def __str__(self):
        return f'{self.name} ({self.code})'

    # --- Parsed accessors -------------------------------------------------

    @property
    def to_addresses(self):
        return _split_addresses(self.email_to)

    @property
    def cc_addresses(self):
        return _split_addresses(self.email_cc)

    @property
    def bcc_addresses(self):
        return _split_addresses(self.email_bcc)

    @property
    def role_keys(self):
        return [item.lower() for item in _split_addresses(self.recipient_roles)]

    @property
    def weekday_numbers(self):
        """Set of ints 0-6. Empty set means 'no restriction'."""
        numbers = set()
        for item in _split_addresses(self.weekdays):
            try:
                value = int(item)
            except (TypeError, ValueError):
                continue
            if 0 <= value <= 6:
                numbers.add(value)
        return numbers

    def effective_filters(self):
        """The full filter dict handed to the report engine.

        `report_filters` stays the general escape hatch for any key a report
        understands; the period control is layered on top of it because that is
        the one an admin can actually see, so it must win over a stale `period`
        left behind in the raw JSON.
        """
        filters = dict(self.report_filters or {})
        if not self.report_period:
            return filters

        filters['period'] = self.report_period
        if self.report_period == PERIOD_CUSTOM:
            if self.report_date_from:
                filters['date_from'] = self.report_date_from.isoformat()
            if self.report_date_to:
                filters['date_to'] = self.report_date_to.isoformat()
        else:
            # A preset covers its own window — leftover explicit dates would
            # only muddy the filter summary shown in the email.
            filters.pop('date_from', None)
            filters.pop('date_to', None)
        return filters

    def resolve_run_as_user(self):
        """The user whose permissions the report executes under."""
        if self.run_as and self.run_as.is_active:
            return self.run_as
        from django.contrib.auth import get_user_model
        return get_user_model().objects.filter(is_superuser=True, is_active=True).order_by('id').first()


class BotExecution(models.Model):
    """One attempt to run a bot. Stores the rendered email verbatim so the
    history screen can answer "what did they actually receive?" without
    regenerating the report."""

    bot = models.ForeignKey(BotAutomation, on_delete=models.CASCADE, related_name='executions')
    trigger = models.CharField(max_length=10, choices=TRIGGER_CHOICES, default=TRIGGER_AUTO)
    status = models.CharField(max_length=10, choices=EXECUTION_STATUS_CHOICES, default=STATUS_PENDING)

    started_at = models.DateTimeField(auto_now_add=True)
    finished_at = models.DateTimeField(null=True, blank=True)
    duration_seconds = models.FloatField(null=True, blank=True)

    record_count = models.IntegerField(null=True, blank=True)
    recipients_to = models.TextField(blank=True)
    recipients_cc = models.TextField(blank=True)
    recipients_bcc = models.TextField(blank=True)

    rendered_subject = models.CharField(max_length=255, blank=True)
    rendered_body = models.TextField(blank=True)
    attachment_name = models.CharField(max_length=255, blank=True)

    error_message = models.TextField(blank=True)
    triggered_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='bot_executions_triggered',
    )

    class Meta:
        verbose_name = 'Bot Execution'
        verbose_name_plural = 'Bot Executions'
        ordering = ['-started_at']
        indexes = [
            models.Index(fields=['bot', '-started_at']),
            models.Index(fields=['status', '-started_at']),
        ]

    def __str__(self):
        return f'{self.bot.name} — {self.started_at:%Y-%m-%d %H:%M} ({self.status})'

    @property
    def is_failure(self):
        return self.status == STATUS_FAILED
