"""Released jobs list and plate replacement requests from production."""

from __future__ import annotations

from collections import defaultdict

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.exceptions import ValidationError
from django.core.paginator import Paginator
from django.db.models import Count, Exists, IntegerField, OuterRef, Q, Subquery, Value
from django.db.models.functions import Coalesce, Lower
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from core.models import JOB_CARD_PRODUCTION_START_STATUSES, JobCard, Production
from planning.models import SkuRecipe
from planning.services import normalize_awc_no
from printing_plates.constants import PLATE_INK_OPTIONS
from printing_plates.models import PlateRequest
from printing_plates.services import (
    PLATE_REQUEST_OPEN_STATUSES,
    REPLACEMENT_SOURCES,
    request_plate_remake,
)

PAGE_SIZE = 50


def user_can_request_plate_remake(user):
    if getattr(user, 'is_superuser', False):
        return True
    profile = getattr(user, 'profile', None)
    return bool(profile and profile.can_view_released_jobs())


def _replacement_q():
    return Q(source__in=REPLACEMENT_SOURCES) | ~Q(replacement_reason='')


def _plate_request_matches_jobcard():
    """Match PlateRequest rows linked by job_card or planning_job (same as job_card_plate_requests_qs)."""
    return Q(job_card_id=OuterRef('pk')) | (
        Q(planning_job_id=OuterRef('planning_job_id')) & Q(planning_job_id__isnull=False)
    )


def released_print_jobs_queryset():
    open_replacement = (
        PlateRequest.objects.filter(_plate_request_matches_jobcard())
        .filter(status__in=PLATE_REQUEST_OPEN_STATUSES)
        .filter(_replacement_q())
    )
    printing_entry = Production.objects.filter(
        job_card_id=OuterRef('pk'),
        is_active=True,
        entry_type='printing',
    )
    remake_count_subquery = (
        PlateRequest.objects.filter(_plate_request_matches_jobcard())
        .filter(_replacement_q())
        .order_by()
        .values(marker=Value(1))
        .annotate(c=Count('id'))
        .values('c')[:1]
    )
    last_production_subquery = (
        Production.objects.filter(job_card_id=OuterRef('pk'))
        .order_by('-date')
        .values('date')[:1]
    )
    return (
        JobCard.objects.filter(
            is_active=True,
            is_print_job=True,
            status__in=JOB_CARD_PRODUCTION_START_STATUSES,
        )
        .select_related('planning_job', 'machine_name')
        .annotate(
            last_production_date=Subquery(last_production_subquery),
            waiting_for_plate=Exists(open_replacement),
            has_printing_entry=Exists(printing_entry),
            remake_count=Coalesce(
                Subquery(remake_count_subquery, output_field=IntegerField()),
                Value(0),
            ),
        )
        .order_by('-updated_at', '-id')
    )


def _batch_sku_recipes(job_cards):
    skus = set()
    for job_card in job_cards:
        sku = (job_card.SKU or '').strip()
        if not sku and job_card.planning_job:
            sku = (job_card.planning_job.sku or '').strip()
        if sku:
            skus.add(sku.upper())
    if not skus:
        return {}
    recipes = SkuRecipe.objects.annotate(sku_lower=Lower('sku')).filter(sku_lower__in=skus)
    return {(recipe.sku or '').strip().upper(): recipe for recipe in recipes}


def _prefetch_plate_history(job_cards):
    """Load recent plate request history for the page in one query, grouped by job card."""
    if not job_cards:
        return {}

    job_card_ids = [jc.pk for jc in job_cards]
    planning_ids = [jc.planning_job_id for jc in job_cards if jc.planning_job_id]
    history_qs = (
        PlateRequest.objects.filter(
            Q(job_card_id__in=job_card_ids) | Q(planning_job_id__in=planning_ids)
        )
        .select_related(
            'requested_by',
            'replaces_request',
            'planning_job',
            'job_card',
            'sku_recipe',
            'job_card__planning_job',
        )
        .order_by('-requested_at', '-created_at')
    )

    by_job_card = defaultdict(list)
    by_planning = defaultdict(list)
    for req in history_qs:
        if req.job_card_id:
            by_job_card[req.job_card_id].append(req)
        if req.planning_job_id:
            by_planning[req.planning_job_id].append(req)

    history_map = {}
    for job_card in job_cards:
        seen = set()
        rows = []
        for req in by_job_card.get(job_card.pk, []):
            if req.pk not in seen:
                seen.add(req.pk)
                rows.append(req)
        if job_card.planning_job_id:
            for req in by_planning.get(job_card.planning_job_id, []):
                if req.pk not in seen:
                    seen.add(req.pk)
                    rows.append(req)
        rows.sort(
            key=lambda r: (
                r.requested_at or r.created_at,
                r.pk or 0,
            ),
            reverse=True,
        )
        history_map[job_card.pk] = rows[:8]
    return history_map


def _resolve_plate_set_no(job_card, history_rows):
    if (job_card.plate_set_no or '').strip():
        return job_card.plate_set_no.strip()
    planning_job = job_card.planning_job
    if planning_job and (planning_job.plate_set_no or '').strip():
        return planning_job.plate_set_no.strip()
    issued_statuses = {PlateRequest.STATUS_AVAILABLE, PlateRequest.STATUS_ARCHIVED}
    for req in history_rows:
        if req.status in issued_statuses:
            value = (req.set_no or req.new_set_no or '').strip()
            if value:
                return value
    for req in history_rows:
        value = (req.set_no or req.new_set_no or '').strip()
        if value:
            return value
    return ''


def _resolve_awc_no(job_card, recipe_map, history_rows):
    planning_job = job_card.planning_job
    sku = (job_card.SKU or '').strip()
    if not sku and planning_job:
        sku = (planning_job.sku or '').strip()
    recipe = recipe_map.get(sku.upper()) if sku else None
    if recipe:
        value = normalize_awc_no(recipe.awc_no)
        if value:
            return value
    for req in history_rows:
        value = normalize_awc_no(req.awc_no)
        if value:
            return value
    return ''


def build_released_job_row(job_card, *, recipe_map=None, history_map=None):
    recipe_map = recipe_map or {}
    history_map = history_map or {}
    planning_job = job_card.planning_job
    history = history_map.get(job_card.pk, [])
    waiting = bool(getattr(job_card, 'waiting_for_plate', False))
    return {
        'job_card': job_card,
        'jc_number': job_card.job_card_no,
        'sku': job_card.SKU or '-',
        'customer': job_card.destination or '-',
        'machine': job_card.machine_name_display or '-',
        'plate_set_no': _resolve_plate_set_no(job_card, history) or '-',
        'awc_no': _resolve_awc_no(job_card, recipe_map, history) or '-',
        'planning_stage': (
            planning_job.get_planning_stage_display()
            if planning_job and planning_job.planning_stage
            else '-'
        ),
        'last_production_date': job_card.last_production_date,
        'plate_status': 'Waiting for plate' if waiting else 'Ready',
        'waiting_for_plate': waiting,
        'remake_count': getattr(job_card, 'remake_count', 0) or 0,
        'history': history,
        'has_printing_entry': bool(getattr(job_card, 'has_printing_entry', False)),
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

    if plate_status == 'waiting':
        qs = qs.filter(waiting_for_plate=True)
    elif plate_status == 'ready':
        qs = qs.filter(waiting_for_plate=False)

    totals = qs.aggregate(
        total_count=Count('id'),
        waiting_count=Count('id', filter=Q(waiting_for_plate=True)),
    )
    total_count = totals['total_count'] or 0
    waiting_count = totals['waiting_count'] or 0
    ready_count = total_count - waiting_count

    paginator = Paginator(qs, PAGE_SIZE)
    page_obj = paginator.get_page(request.GET.get('page'))
    page_jobs = list(page_obj.object_list)

    recipe_map = _batch_sku_recipes(page_jobs)
    history_map = _prefetch_plate_history(page_jobs)
    rows = [
        build_released_job_row(job, recipe_map=recipe_map, history_map=history_map)
        for job in page_jobs
    ]

    filter_params = request.GET.copy()
    filter_params.pop('page', None)
    filter_query = filter_params.urlencode()

    context = {
        'page_obj': page_obj,
        'rows': rows,
        'q': q,
        'plate_status': plate_status,
        'total_count': total_count,
        'waiting_count': waiting_count,
        'ready_count': ready_count,
        'reason_choices': PlateRequest.REPLACEMENT_REASON_CHOICES,
        'can_request': user_can_request_plate_remake(request.user),
        'plate_ink_options': PLATE_INK_OPTIONS,
        'filter_query': filter_query,
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
