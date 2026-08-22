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

from .jobcard_service import close_job_card_manually, job_card_completion_blockers, reopen_job_card_manually
from .models import JobCard
from .services import compute_job_card_wastage_metrics
from .views import permission_required

DEFAULT_DISPATCH_FLOOR_PERCENT = 80
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
        'blockers': job_card_completion_blockers(job_card),
    }


def _to_float(raw_value, default=None):
    raw_value = (raw_value or '').strip()
    if not raw_value:
        return default
    try:
        return float(raw_value)
    except ValueError:
        return default


def _matches_filters(row, search, min_wastage_pct, max_wastage_pct, min_dispatch_pct, max_dispatch_pct):
    if search:
        haystack = ' '.join([
            row['job_card'].job_card_no or '',
            row['job_card'].SKU or '',
            row['job_card'].PO_No or '',
        ]).upper()
        if search.upper() not in haystack:
            return False
    if min_wastage_pct is not None and row['wastage']['total_wastage_pct'] < min_wastage_pct:
        return False
    if max_wastage_pct is not None and row['wastage']['total_wastage_pct'] > max_wastage_pct:
        return False
    if min_dispatch_pct is not None and row['dispatch_completion_percent'] < min_dispatch_pct:
        return False
    if max_dispatch_pct is not None and row['dispatch_completion_percent'] > max_dispatch_pct:
        return False
    return True


@login_required
@permission_required('can_finalize_job_card')
def job_card_finalization_queue(request):
    search = (request.GET.get('q') or '').strip()
    min_wastage_pct = _to_float(request.GET.get('min_wastage_pct'))
    max_wastage_pct = _to_float(request.GET.get('max_wastage_pct'))
    min_dispatch_pct = _to_float(request.GET.get('min_dispatch_pct'), default=DEFAULT_DISPATCH_FLOOR_PERCENT)
    max_dispatch_pct = _to_float(request.GET.get('max_dispatch_pct'))

    # Every _row() computes dispatch/printing/packing/wastage figures from
    # JobCard.total_dispatch / printing_productions / packing_productions,
    # which reuse a prefetched productions/dispatch_set instead of a fresh
    # query per job card (JobCard._is_prefetched) — without this, a page of
    # a few hundred job cards fires 10+ queries each, which is what was
    # making this page slow to load.
    related = dict(
        select_related=('planning_job', 'material'),
        prefetch_related=('productions', 'dispatch_set'),
    )
    stuck_floor = min_dispatch_pct if min_dispatch_pct is not None else DEFAULT_DISPATCH_FLOOR_PERCENT
    stuck_jobs = [
        jc for jc in JobCard.objects.filter(is_active=True, status='in_production')
            .select_related(*related['select_related'])
            .prefetch_related(*related['prefetch_related'])
            .order_by('-updated_at')[:STUCK_LIST_CAP]
        if jc.dispatch_completion_percent >= stuck_floor
    ]
    stuck_jobs.sort(key=lambda jc: jc.dispatch_completion_percent, reverse=True)

    completed_not_closed = list(
        JobCard.objects.filter(is_active=True, status='completed')
        .select_related(*related['select_related'])
        .prefetch_related(*related['prefetch_related'])
        .order_by('-updated_at')[:STUCK_LIST_CAP]
    )

    closed_jobs = list(
        JobCard.objects.filter(is_active=True, status='closed')
        .select_related(*related['select_related'])
        .prefetch_related(*related['prefetch_related'])
        .order_by('-updated_at')[:STUCK_LIST_CAP]
    )

    stuck_rows = [_row(jc) for jc in stuck_jobs]
    completed_rows = [_row(jc) for jc in completed_not_closed]
    closed_rows = [_row(jc) for jc in closed_jobs]

    filter_args = (search, min_wastage_pct, max_wastage_pct, min_dispatch_pct, max_dispatch_pct)
    stuck_rows = [row for row in stuck_rows if _matches_filters(row, *filter_args)]
    completed_rows = [row for row in completed_rows if _matches_filters(row, *filter_args)]

    context = {
        'stuck_rows': stuck_rows,
        'completed_rows': completed_rows,
        'closed_rows': closed_rows,
        'stuck_floor_percent': stuck_floor,
        'filter_q': search,
        'filter_min_wastage_pct': request.GET.get('min_wastage_pct', ''),
        'filter_max_wastage_pct': request.GET.get('max_wastage_pct', ''),
        'filter_min_dispatch_pct': request.GET.get('min_dispatch_pct') or stuck_floor,
        'filter_max_dispatch_pct': request.GET.get('max_dispatch_pct', ''),
    }
    return render(request, 'job_card_finalization.html', context)


@login_required
@permission_required('can_finalize_job_card')
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
@permission_required('can_finalize_job_card')
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
