from __future__ import annotations

from datetime import timedelta

from django.utils import timezone


def calculate_next_run(frequency: str, from_dt=None):
    now = from_dt or timezone.now()
    if frequency == 'daily':
        return now + timedelta(days=1)
    if frequency == 'weekly':
        return now + timedelta(days=7)
    if frequency == 'monthly':
        return now + timedelta(days=30)
    if frequency == 'quarterly':
        return now + timedelta(days=90)
    return now + timedelta(days=1)
