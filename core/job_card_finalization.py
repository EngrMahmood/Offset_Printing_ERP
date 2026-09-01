"""Admin/manager tool to manually close job cards that are stuck near
completion — the dispatch-% auto-complete signal (core.signals) never
reaches 'closed' by itself, and never fires at all once a permanent
wastage/shortfall keeps a job under the 95% dispatch threshold forever.
See core.jobcard_service.close_job_card_manually / reopen_job_card_manually.
"""

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Q
from django.shortcuts import get_object_or_404, redirect, render
from django.views.decorators.http import require_POST

from .jobcard_service import close_job_card_manually, job_card_completion_blockers, reopen_job_card_manually
from .models import JobCard
from .services import compute_job_card_wastage_metrics
from .views import permission_required

DEFAULT_DISPATCH_FLOOR_PERCENT = 80
STUCK_LIST_CAP = 500
# The closed list is reopen-a-mistake history, not a working queue, and it is
# the largest of the three sections (1000+ rows and growing) — every row costs
# a prefetched productions/dispatch set to compute its wastage figures, which
# dominated this page's load time. Keep the rendered slice small and let the
# search box reach the rest at the database level instead.
CLOSED_LIST_CAP = 100


def _row(job_card, *, with_blockers=True):
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
        # Only the two open sections offer a Close button, so the closed list
        # skips this — it is pure wasted work for a column it never renders.
        'blockers': job_card_completion_blockers(job_card) if with_blockers else [],
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


def _build_filtered_sections(request):
    """Rows for all three queue sections, honouring the request's filters.

    Shared by the page and its exports so a download can never disagree with
    what the planner is looking at on screen.
    """
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

    # Closed jobs are matched in the database rather than in Python: the list is
    # far longer than the rendered slice, so an in-memory filter could only ever
    # search the slice — a closed job outside it was unfindable. The numeric
    # dispatch/wastage filters stay off this section deliberately; they triage
    # the open queues, and applying the default 80% dispatch floor here would
    # silently hide closed history the planner is looking for.
    closed_qs = JobCard.objects.filter(is_active=True, status='closed')
    if search:
        closed_qs = closed_qs.filter(
            Q(job_card_no__icontains=search)
            | Q(SKU__icontains=search)
            | Q(PO_No__icontains=search)
        )
    closed_total = closed_qs.count()
    closed_jobs = list(
        closed_qs
        .select_related(*related['select_related'])
        .prefetch_related(*related['prefetch_related'])
        .order_by('-updated_at')[:CLOSED_LIST_CAP]
    )

    stuck_rows = [_row(jc) for jc in stuck_jobs]
    completed_rows = [_row(jc) for jc in completed_not_closed]
    closed_rows = [_row(jc, with_blockers=False) for jc in closed_jobs]

    filter_args = (search, min_wastage_pct, max_wastage_pct, min_dispatch_pct, max_dispatch_pct)
    stuck_rows = [row for row in stuck_rows if _matches_filters(row, *filter_args)]
    completed_rows = [row for row in completed_rows if _matches_filters(row, *filter_args)]

    return {
        'stuck_rows': stuck_rows,
        'completed_rows': completed_rows,
        'closed_rows': closed_rows,
        'closed_total': closed_total,
        'closed_truncated': closed_total > len(closed_rows),
        'stuck_floor': stuck_floor,
        'search': search,
    }


@login_required
@permission_required('can_finalize_job_card')
def job_card_finalization_queue(request):
    sections = _build_filtered_sections(request)
    stuck_floor = sections['stuck_floor']

    context = {
        'stuck_rows': sections['stuck_rows'],
        'completed_rows': sections['completed_rows'],
        'closed_rows': sections['closed_rows'],
        'closed_total': sections['closed_total'],
        'closed_truncated': sections['closed_truncated'],
        'stuck_floor_percent': stuck_floor,
        'filter_q': sections['search'],
        'filter_min_wastage_pct': request.GET.get('min_wastage_pct', ''),
        'filter_max_wastage_pct': request.GET.get('max_wastage_pct', ''),
        'filter_min_dispatch_pct': request.GET.get('min_dispatch_pct') or stuck_floor,
        'filter_max_dispatch_pct': request.GET.get('max_dispatch_pct', ''),
        'export_qs': request.GET.urlencode(),
    }
    return render(request, 'job_card_finalization.html', context)


# Section key -> (section label, context key, ordered export columns).
# Closed job cards drop the printed/packed/gap columns the other two carry:
# once closed the gap is already settled into confirmed wastage, so those
# figures no longer describe an actionable shortfall.
_EXPORT_COMMON_COLUMNS = [
    ('job_card_no', 'Job Card'),
    ('po_no', 'PO / WO No'),
    ('sku', 'SKU'),
    ('order_qty', 'Order Qty'),
    ('dispatched', 'Dispatched'),
    ('dispatch_pct', 'Dispatch %'),
]
_EXPORT_SECTIONS = {
    'stuck': ('Stuck Near-Complete', 'stuck_rows', _EXPORT_COMMON_COLUMNS + [
        ('printed', 'Printed'),
        ('packed', 'Packed'),
        ('wastage_pct', 'Wastage %'),
        ('gap_qty', 'Gap (to Wastage)'),
        ('close_status', 'Close Status'),
    ]),
    'completed': ('Completed - Not Yet Closed', 'completed_rows', _EXPORT_COMMON_COLUMNS + [
        ('printed', 'Printed'),
        ('packed', 'Packed'),
        ('wastage_pct', 'Wastage %'),
        ('gap_qty', 'Gap (to Wastage)'),
        ('close_status', 'Close Status'),
    ]),
    'closed': ('Closed', 'closed_rows', _EXPORT_COMMON_COLUMNS + [
        ('wastage_pct', 'Wastage %'),
        ('wastage_status', 'Wastage Status'),
    ]),
}


def _export_row(row):
    """Flatten one queue row into the plain scalars the exporters render."""
    job_card = row['job_card']
    blockers = row.get('blockers') or []
    return {
        'job_card_no': job_card.job_card_no or '',
        'po_no': job_card.PO_No or '',
        'sku': job_card.SKU or '',
        'order_qty': row['order_qty'] or 0,
        'dispatched': row['total_dispatch'] or 0,
        'dispatch_pct': row['dispatch_completion_percent'],
        'printed': row['total_printed_pcs'] or 0,
        'packed': row['total_packed_pcs'] or 0,
        'wastage_pct': row['wastage']['total_wastage_pct'],
        'gap_qty': row['gap_qty'] or 0,
        'wastage_status': row['wastage'].get('wastage_status') or '',
        'close_status': ' '.join(blockers) if blockers else 'Ready to close',
    }


@login_required
@permission_required('can_finalize_job_card')
def job_card_finalization_export(request):
    """Excel/PDF of one queue section, filtered exactly as the page is.

    Reuses the shared report exporters (reports.export.services) rather than
    building a second spreadsheet/PDF writer — same branding and layout as
    every other export in the ERP.
    """
    from django.http import HttpResponse
    from django.utils import timezone

    from reports.export.services import export_as_pdf, export_as_xlsx

    section_key = (request.GET.get('section') or 'stuck').strip().lower()
    if section_key not in _EXPORT_SECTIONS:
        section_key = 'stuck'
    section_label, context_key, columns = _EXPORT_SECTIONS[section_key]

    sections = _build_filtered_sections(request)
    rows = [_export_row(row) for row in sections[context_key]]

    payload = {
        'report': {'title': f'Job Card Finalization - {section_label}', 'slug': 'job-card-finalization'},
        'generated_at': timezone.localtime().strftime('%Y-%m-%d %H:%M'),
        'headers': [key for key, _label in columns],
        'header_labels': {key: label for key, label in columns},
        'data': {'export_rows': rows},
    }

    filename = f'job-card-finalization-{section_key}-{timezone.localdate():%Y%m%d}'
    export_type = (request.GET.get('type') or 'xlsx').strip().lower()
    try:
        if export_type == 'pdf':
            response = HttpResponse(export_as_pdf(payload), content_type='application/pdf')
            response['Content-Disposition'] = f'attachment; filename="{filename}.pdf"'
            return response
        response = HttpResponse(
            export_as_xlsx(payload),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}.xlsx"'
        return response
    except RuntimeError as exc:
        # openpyxl/reportlab missing — say so instead of a 500.
        messages.error(request, str(exc))
        return redirect('job_card_finalization_queue')


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
