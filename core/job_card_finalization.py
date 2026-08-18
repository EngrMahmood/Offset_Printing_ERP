"""Admin/manager tool to manually close job cards that are stuck near
completion — the dispatch-% auto-complete signal (core.signals) never
reaches 'closed' by itself, and never fires at all once a permanent
wastage/shortfall keeps a job under the 95% dispatch threshold forever.
See core.jobcard_service.close_job_card_manually / reopen_job_card_manually.
"""

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import get_object_or_404, redirect, render
from django.views.decorators.http import require_POST

from .jobcard_service import close_job_card_manually, reopen_job_card_manually
from .models import JobCard
from .services import compute_job_card_wastage_metrics
from .views import require_role

STUCK_DISPATCH_FLOOR_PERCENT = 80
STUCK_LIST_CAP = 500


def _row(job_card):
    wastage = compute_job_card_wastage_metrics(job_card)
    return {
        'job_card': job_card,
        'order_qty': job_card.order_qty,
        'total_dispatch': job_card.total_dispatch,
        'dispatch_completion_percent': job_card.dispatch_completion_percent,
        'total_printed_pcs': job_card.total_printed_pcs,
        'total_packed_pcs': job_card.total_packed_pcs,
        'gap_qty': max(job_card.order_qty - job_card.total_dispatch, 0),
        'wastage': wastage,
    }


@login_required
@require_role('admin', 'manager')
def job_card_finalization_queue(request):
    stuck_jobs = [
        jc for jc in JobCard.objects.filter(is_active=True, status='in_production')
            .select_related('planning_job', 'material')
            .order_by('-updated_at')[:STUCK_LIST_CAP]
        if jc.dispatch_completion_percent >= STUCK_DISPATCH_FLOOR_PERCENT
    ]
    stuck_jobs.sort(key=lambda jc: jc.dispatch_completion_percent, reverse=True)

    completed_not_closed = list(
        JobCard.objects.filter(is_active=True, status='completed')
        .select_related('planning_job', 'material')
        .order_by('-updated_at')[:STUCK_LIST_CAP]
    )

    closed_jobs = list(
        JobCard.objects.filter(is_active=True, status='closed')
        .select_related('planning_job', 'material')
        .order_by('-updated_at')[:STUCK_LIST_CAP]
    )

    context = {
        'stuck_rows': [_row(jc) for jc in stuck_jobs],
        'completed_rows': [_row(jc) for jc in completed_not_closed],
        'closed_rows': [_row(jc) for jc in closed_jobs],
        'stuck_floor_percent': STUCK_DISPATCH_FLOOR_PERCENT,
    }
    return render(request, 'job_card_finalization.html', context)


@login_required
@require_role('admin', 'manager')
@require_POST
def job_card_finalization_close(request):
    reason = (request.POST.get('reason') or '').strip()
    job_card_ids = request.POST.getlist('job_card_ids')

    if not reason:
        messages.error(request, 'A reason is required to close job cards.')
        return redirect('job_card_finalization_queue')
    if not job_card_ids:
        messages.error(request, 'Select at least one job card to close.')
        return redirect('job_card_finalization_queue')

    closed_count = 0
    failed = []
    for job_card_id in job_card_ids:
        try:
            job_card = JobCard.objects.get(pk=job_card_id, is_active=True)
            close_job_card_manually(job_card, actor=request.user, reason=reason)
            closed_count += 1
        except JobCard.DoesNotExist:
            continue
        except Exception as exc:
            failed.append((job_card_id, str(exc)))

    if closed_count:
        messages.success(request, f'Closed {closed_count} job card(s).')
    for job_card_id, error in failed:
        messages.error(request, f'Job card #{job_card_id}: {error}')

    return redirect('job_card_finalization_queue')


@login_required
@require_role('admin', 'manager')
@require_POST
def job_card_finalization_reopen(request):
    reason = (request.POST.get('reason') or '').strip()
    job_card_id = request.POST.get('job_card_id')

    if not reason:
        messages.error(request, 'A reason is required to reopen a job card.')
        return redirect('job_card_finalization_queue')

    job_card = get_object_or_404(JobCard, pk=job_card_id, is_active=True)
    try:
        reopen_job_card_manually(job_card, actor=request.user, reason=reason)
        messages.success(request, f'Reopened {job_card.job_card_no} to In Production.')
    except Exception as exc:
        messages.error(request, f'Could not reopen {job_card.job_card_no}: {exc}')

    return redirect('job_card_finalization_queue')
