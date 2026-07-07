from __future__ import annotations

from django.utils import timezone

from reports.report_engine import run_report
from reports.scheduler.services import calculate_next_run

try:
    from celery import shared_task
except Exception:  # pragma: no cover
    def shared_task(*args, **kwargs):  # type: ignore
        def decorator(func):
            return func
        return decorator


@shared_task(bind=True)
def execute_scheduled_report(self, schedule_id: int):
    from reports.models import ScheduledReport

    schedule = ScheduledReport.objects.filter(id=schedule_id, is_active=True).first()
    if not schedule:
        return {'ok': False, 'error': 'Schedule not found'}

    # Build a minimal request-like object for engine compatibility.
    class RequestLike:
        GET = {}
        user = None

    req = RequestLike()
    req.GET = schedule.filters or {}
    req.user = schedule.created_by

    run_report(schedule.report_slug, req)

    schedule.last_run_at = timezone.now()
    schedule.next_run_at = calculate_next_run(schedule.frequency, schedule.last_run_at)
    schedule.save(update_fields=['last_run_at', 'next_run_at', 'updated_at'])
    return {'ok': True, 'schedule_id': schedule.id}
