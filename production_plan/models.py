from datetime import time

from django.conf import settings
from django.db import models

SHIFT_CHOICES = [('D', 'Day (8 am - 8 pm)'), ('N', 'Night (8 pm - 8 am)')]
LANE_CHOICES = [
    ('black', 'Black (1 and 1+1 colour)'),
    ('colour_gto', 'Colour - GTO sheet'),
    ('colour_sm', 'Colour - SM 74 sheet'),
]
PRIORITY_LABELS = {0: 'Low', 1: 'Normal', 2: 'Medium', 3: 'High', 4: 'Urgent'}


class PlanSettings(models.Model):
    """Single-row configuration for the Production Plan (id=1)."""

    shift_capacity = models.PositiveIntegerField(default=40000, help_text='Impressions per machine per shift.')
    friday_day_capacity = models.PositiveIntegerField(
        default=35000, help_text='Friday day-shift capacity (1.5 hours go to lunch and Jumma prayer).',
    )
    gto_max_sheet_long_in = models.DecimalField(
        max_digits=5, decimal_places=2, default=18,
        help_text='Longest side of the biggest print sheet the GTOs take. Bigger sheets are SM 74 jobs.',
    )
    gto_max_sheet_short_in = models.DecimalField(max_digits=5, decimal_places=2, default=12.5)
    horizon_working_days = models.PositiveSmallIntegerField(default=2)
    draft_time = models.TimeField(default=time(18, 30), help_text='The evening draft is built at this time.')
    rebalance_time = models.TimeField(default=time(8, 30), help_text='The morning re-balance runs at this time.')
    email_time = models.TimeField(default=time(8, 45), help_text='The plan email goes out at this time.')
    black_machines = models.CharField(max_length=255, default='GTO 1A, GTO 1B, GTO 2A, GTO 2B')
    colour_gto_machines = models.CharField(max_length=255, default='GTO 2C')
    colour_sm_machines = models.CharField(max_length=255, default='SM 74')
    email_to_roles = models.CharField(max_length=255, default='planner', help_text='Roles whose users get the plan email (To).')
    email_to_extra = models.TextField(blank=True, help_text='Extra To addresses, comma or newline separated.')
    email_cc = models.TextField(default='printing@utopia.pk', help_text='CC addresses, comma or newline separated.')
    off_dates = models.TextField(blank=True, help_text='Holidays, one YYYY-MM-DD per line. Sundays are always off.')
    pm_block_shift = models.CharField(
        max_length=1, choices=[('D', 'Day shift'), ('N', 'Night shift'), ('B', 'Both shifts')], default='D',
        help_text='Which shift a preventive-maintenance day blocks on that machine.',
    )
    gate_enabled = models.BooleanField(
        default=True, help_text='Block planners from the planning pages until they approve the evening draft.',
    )
    scheduler_enabled = models.BooleanField(default=True, help_text='Build, re-balance and email the plan automatically.')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'Plan setting'
        verbose_name_plural = 'Plan settings'

    @classmethod
    def get(cls):
        obj, _ = cls.objects.get_or_create(id=1)
        return obj

    def __str__(self):
        return 'Production Plan settings'


class ProductionPlan(models.Model):
    STATUS_CHOICES = [('draft', 'Draft - waiting for planner approval'), ('approved', 'Approved'), ('superseded', 'Superseded')]

    morning_date = models.DateField(db_index=True, help_text='The morning this plan is circulated (first working day it covers).')
    horizon_dates = models.JSONField(default=list)
    status = models.CharField(max_length=12, choices=STATUS_CHOICES, default='draft', db_index=True)
    base_plan = models.ForeignKey('self', null=True, blank=True, on_delete=models.SET_NULL, related_name='followers')
    generated_at = models.DateTimeField(auto_now_add=True)
    approved_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL, related_name='+')
    approved_at = models.DateTimeField(null=True, blank=True)
    rebalanced_at = models.DateTimeField(null=True, blank=True)
    emailed_at = models.DateTimeField(null=True, blank=True)
    email_status = models.CharField(max_length=20, blank=True)
    email_error = models.TextField(blank=True)
    summary = models.JSONField(default=dict, blank=True)

    class Meta:
        ordering = ['-morning_date', '-id']

    def __str__(self):
        return f'Plan for {self.morning_date} ({self.status})'

    @property
    def is_approved(self):
        return self.status == 'approved'


class PlanItem(models.Model):
    plan = models.ForeignKey(ProductionPlan, on_delete=models.CASCADE, related_name='items')
    lane = models.CharField(max_length=12, choices=LANE_CHOICES, db_index=True)
    machine = models.ForeignKey('core.Machine', on_delete=models.PROTECT, related_name='+')
    plan_date = models.DateField(db_index=True)
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES)
    sequence = models.PositiveIntegerField(default=0)
    planning_job = models.ForeignKey('planning.PlanningJob', null=True, blank=True, on_delete=models.SET_NULL, related_name='+')
    jc_number = models.CharField(max_length=50)
    sku = models.CharField(max_length=255, blank=True)
    order_date = models.DateField(null=True, blank=True)
    qty = models.PositiveIntegerField(help_text='Impressions planned in this slot.')
    remaining_at_plan = models.PositiveIntegerField(default=0, help_text="The job's remaining impressions when planned.")
    priority = models.SmallIntegerField(default=1)
    is_new = models.BooleanField(default=False)
    is_chunk = models.BooleanField(default=False, help_text='Part of a job bigger than one shift.')
    pinned = models.BooleanField(default=False, help_text='Kept in place by every re-balance.')
    manual = models.BooleanField(default=False, help_text='Added or moved by a planner.')
    fifo_reason = models.TextField(blank=True, help_text='Why the planner put this job ahead of FIFO.')

    class Meta:
        ordering = ['plan_date', 'shift', 'machine_id', 'sequence', 'id']
        indexes = [models.Index(fields=['plan', 'lane', 'plan_date', 'shift'])]

    def __str__(self):
        return f'{self.jc_number} {self.qty:,} on {self.machine} {self.plan_date} {self.shift}'


class PlanBlock(models.Model):
    """A machine that cannot run in a shift (maintenance, breakdown, holiday on one press)."""

    machine = models.ForeignKey('core.Machine', on_delete=models.CASCADE, related_name='+')
    date = models.DateField(db_index=True)
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES + [('B', 'Both shifts')], default='D')
    reason = models.CharField(max_length=200, blank=True)
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL, related_name='+')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('machine', 'date', 'shift')
        ordering = ['date', 'machine_id']


class PlanEvent(models.Model):
    plan = models.ForeignKey(ProductionPlan, on_delete=models.CASCADE, related_name='events')
    at = models.DateTimeField(auto_now_add=True)
    user = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL, related_name='+')
    action = models.CharField(max_length=40)
    detail = models.TextField(blank=True)

    class Meta:
        ordering = ['-at', '-id']


class PlanRunLog(models.Model):
    """One row per scheduled task per day, so a restart never runs the same task twice."""

    task = models.CharField(max_length=20)
    run_date = models.DateField()
    status = models.CharField(max_length=12, default='running')
    detail = models.TextField(blank=True)
    attempts = models.PositiveSmallIntegerField(default=1)
    started_at = models.DateTimeField(auto_now_add=True)
    finished_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        unique_together = ('task', 'run_date')
        ordering = ['-run_date', '-id']
