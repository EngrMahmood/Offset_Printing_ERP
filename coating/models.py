from django.conf import settings
from django.db import models


class CoatingRun(models.Model):
    """One coating (UV / varnish) run on a job card.

    Coating goes through a press but is not a print pass: it is recorded here, apart from
    core.Production, so printing, packing and OEE figures and the job's impressions stay as they are.
    """

    SHIFT_CHOICES = [('A', 'Shift A'), ('B', 'Shift B')]
    TYPE_CHOICES = [
        ('uv', 'UV'),
        ('matt_varnish', 'Matt Varnish'),
        ('gloss_varnish', 'Gloss Varnish'),
        ('other', 'Other coating'),
    ]
    STATUS_CHOICES = [('in_progress', 'In Progress'), ('completed', 'Completed (coating done)')]

    job_card = models.ForeignKey('core.JobCard', on_delete=models.CASCADE, related_name='coating_runs')
    date = models.DateField(db_index=True)
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES)
    machine = models.ForeignKey('core.Machine', on_delete=models.PROTECT, null=True, blank=True, related_name='coating_runs')
    operator = models.ForeignKey('core.Operator', on_delete=models.SET_NULL, null=True, blank=True, related_name='coating_runs')
    coating_type = models.CharField(max_length=20, choices=TYPE_CHOICES, default='uv')
    coats = models.PositiveSmallIntegerField(default=1, help_text='Coats applied to each sheet (normally 1).')
    sides = models.PositiveSmallIntegerField(default=1, help_text='2 when the job is coated on front and back (each sheet then runs twice).')
    sheets_coated = models.PositiveIntegerField(default=0, help_text='Good sheets coated.')
    waste_sheets = models.PositiveIntegerField(default=0)
    start_time = models.TimeField(null=True, blank=True)
    end_time = models.TimeField(null=True, blank=True)
    run_minutes = models.FloatField(default=0, help_text='Minutes the press ran this coating.')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='in_progress')
    remark_notes = models.TextField(blank=True)
    change_reason = models.TextField(blank=True)
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, editable=False, related_name='+')
    is_active = models.BooleanField(default=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-date', '-id']

    def __str__(self):
        return f'Coating {self.job_card_id} {self.date} {self.shift}'

    @property
    def impressions(self):
        """Press impressions the run used: every sheet put through, good or waste, times the coats."""
        return (self.sheets_coated + self.waste_sheets) * max(self.coats, 1)
