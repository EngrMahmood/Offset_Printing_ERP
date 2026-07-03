"""Released jobs list and plate replacement requests from production."""

from __future__ import annotations

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.exceptions import ValidationError
from django.core.paginator import Paginator
from django.db.models import Max, Q
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from core.models import JOB_CARD_PRODUCTION_START_STATUSES, JobCard
from core.views import permission_required
from printing_plates.constants import PLATE_INK_OPTIONS
from printing_plates.models import PlateRequest
from printing_plates.services import (
    get_job_card_awc_no,
    get_job_card_plate_set_no,
    get_open_replacement_requests,
    get_plate_remake_count,
    job_is_waiting_for_plates,
    request_plate_remake,
)


def user_can_request_plate_remake(user):
    profile = getattr(user, 'profile', None)
    if not profile:
        return bool(getattr(user, 'is_superuser', False))
    return (
        profile.can_edit_production()
        or profile.can_edit_jobcard()
        or profile.normalized_role in {'graphics_designer', 'admin', 'manager'}
        or getattr(user, 'is_superuser', False)
    )


def released_print_jobs_queryset():
    return (
        JobCard.objects.filter(
            is_active=True,
            is_print_job=True,
            status__in=JOB_CARD_PRODUCTION_START_STATUSES,
        )
        .select_related('planning_job', 'machine_name', 'material')
        .annotate(last_production_date=Max('productions__date'))
        .order_by('-updated_at', '-id')
    )


def build_released_job_row(job_card):
    planning_job = job_card.planning_job
    waiting = job_is_waiting_for_plates(job_card)
    open_requests = list(get_open_replacement_requests(job_card)[:5])
    history = list(
        PlateRequest.objects.filter(
            Q(job_card=job_card) | Q(planning_job=planning_job) if planning_job else Q(job_card=job_card)
        )
        .select_related('requested_by', 'replaces_request')
        .order_by('-requested_at', '-created_at')[:8]
    )
    has_printing_entry = job_card.productions.filter(is_active=True, entry_type='printing').exists()
    return {
        'job_card': job_card,
        'jc_number': job_card.job_card_no,
        'sku': job_card.SKU or '-',
        'customer': job_card.destination or '-',
        'machine': job_card.machine_name_display or '-',
        'plate_set_no': get_job_card_plate_set_no(job_card) or '-',
        'awc_no': get_job_card_awc_no(job_card) or '-',
        'planning_stage': (
            planning_job.get_planning_stage_display()
            if planning_job and planning_job.planning_stage
            else '-'
        ),
        'last_production_date': job_card.last_production_date,
        'plate_status': 'Waiting for plate' if waiting else 'Ready',
        'waiting_for_plate': waiting,
        'remake_count': get_plate_remake_count(job_card),
        'open_requests': open_requests,
        'history': history,
        'has_printing_entry': has_printing_entry,
    }


@login_required
def released_jobs(request):
    if not user_can_request_plate_remake(request.user):
        messages.error(request, 'You do not have permission to view released jobs.')
        return redirect('production_dashboard')

    q = (request.GET.get('q') or '').strip()
    plate_status = (request.GET.get('plate_status') or '').strip().lower()
    qs = released_print_jobs_queryset()
    if q:
        qs = qs.filter(
            Q(job_card_no__icontains=q)
            | Q(SKU__icontains=q)
            | Q(destination__icontains=q)
            | Q(PO_No__icontains=q)
            | Q(planning_job__job_name__icontains=q)
            | Q(plate_set_no__icontains=q)
        )

    rows = [build_released_job_row(job) for job in qs[:500]]
    if plate_status == 'waiting':
        rows = [row for row in rows if row['waiting_for_plate']]
    elif plate_status == 'ready':
        rows = [row for row in rows if not row['waiting_for_plate']]

    paginator = Paginator(rows, 50)
    page_obj = paginator.get_page(request.GET.get('page'))

    waiting_count = sum(1 for row in rows if row['waiting_for_plate'])
    context = {
        'page_obj': page_obj,
        'rows': page_obj.object_list,
        'q': q,
        'plate_status': plate_status,
        'total_count': len(rows),
        'waiting_count': waiting_count,
        'ready_count': len(rows) - waiting_count,
        'reason_choices': PlateRequest.REPLACEMENT_REASON_CHOICES,
        'can_request': user_can_request_plate_remake(request.user),
        'plate_ink_options': PLATE_INK_OPTIONS,
    }
    return render(request, 'production/released_jobs.html', context)


@login_required
@require_POST
def request_plate_remake_view(request):
    if not user_can_request_plate_remake(request.user):
        messages.error(request, 'You do not have permission to request plate replacement.')
        return redirect('released_jobs')

    job_card_id = (request.POST.get('job_card_id') or '').strip()
    reason = (request.POST.get('reason') or '').strip()
    damaged_colors = (request.POST.get('damaged_colors') or '').strip()
    notes = (request.POST.get('notes') or '').strip()

    job_card = get_object_or_404(
        JobCard.objects.select_related('planning_job', 'machine_name'),
        pk=job_card_id,
        is_active=True,
        is_print_job=True,
    )

    warn_no_entry = not job_card.productions.filter(is_active=True, entry_type='printing').exists()

    try:
        plate_request = request_plate_remake(
            job_card,
            actor=request.user,
            reason=reason,
            damaged_colors=damaged_colors,
            notes=notes,
            source=PlateRequest.SOURCE_REPLACEMENT,
        )
    except ValidationError as exc:
        message = exc.messages[0] if getattr(exc, 'messages', None) else str(exc)
        messages.error(request, message)
        return redirect('released_jobs')

    if warn_no_entry:
        messages.warning(
            request,
            'No printing entry logged yet for this job. Replacement request was still sent to graphics.',
        )

    messages.success(
        request,
        f'Plate replacement requested for {job_card.job_card_no}. Status: Waiting for plate. '
        f'Graphics request #{plate_request.pk} created.',
    )
    return redirect(f"{reverse('released_jobs')}?q={job_card.job_card_no}")
