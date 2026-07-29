from django.db import models

from core.models import Production


class DailyTarget(models.Model):
    """Production target for the floor dashboard's Plant Overview / Target
    Achievement screens. Set from the admin; if no row exists for a given
    date (and optionally shift), the dashboard falls back to an estimate
    built from today's active job cards, labelled accordingly."""

    date = models.DateField(db_index=True)
    shift = models.CharField(max_length=1, choices=Production.SHIFT_CHOICES, null=True, blank=True)
    target_qty = models.PositiveIntegerField(help_text="Target output in pcs for the date (and shift, if set)")

    class Meta:
        constraints = [
            models.UniqueConstraint(fields=['date', 'shift'], name='unique_daily_target_per_shift'),
        ]
        ordering = ['-date']

    def __str__(self):
        if self.shift:
            return f'{self.date} Shift {self.shift}: {self.target_qty}'
        return f'{self.date}: {self.target_qty}'
