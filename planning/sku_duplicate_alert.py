"""Duplicate SKU cluster alerts for planners (same SKU across active POs/JCs)."""

from __future__ import annotations

from datetime import timedelta

from django.db.models import Q
from django.utils import timezone

from core.models import Dispatch, JobCard
from planning.models import PLANNING_STAGE_CHOICES, PlanningJob
from planning.services import _sku_key
from workflow.services import _normalize_status

RECENT_DISPATCH_DAYS = 14


def _sku_filter_q(sku_values):
    sku_values = [value for value in sku_values if (value or '').strip()]
    if not sku_values:
        return Q(pk__in=[])
    query = Q()
    for sku in sku_values:
        query |= Q(sku__iexact=sku.strip())
    return query


def _job_sort_key(job):
    approval = job.po_approval_date
    plan = job.plan_date
    created = job.created_at.date() if job.created_at else None
    return (
        approval is None,
        approval or plan or created,
        plan or created,
        created,
        job.id or 0,
    )


def _printing_started(job_card):
    if not job_card:
        return False
    return job_card.productions.filter(is_active=True, entry_type='printing').exists()


def _member_status_label(job, job_card):
    if job_card:
        return job_card.workflow_status_label
    status = _normalize_status(job.status)
    labels = {
        'draft': 'Draft',
        'pending_qc': 'Pending QC',
        'qc_approved': 'QC Approved',
        'released': 'Released',
        'in_production': 'In Production',
        'completed': 'Completed',
    }
    return labels.get(status, status.replace('_', ' ').title())


def _member_stage_label(job):
    stage = (job.planning_stage or '').strip()
    if not stage:
        return '-'
    return dict(PLANNING_STAGE_CHOICES).get(stage, stage.replace('_', ' ').title())


def _is_active_planning_job(job):
    if not job or not job.is_active:
        return False
    return _normalize_status(job.status) != 'completed'


def active_jobs_for_sku(sku, *, exclude_job_id=None):
    sku_value = (sku or '').strip()
    if not sku_value:
        return PlanningJob.objects.none()
    qs = (
        PlanningJob.objects.filter(is_active=True, sku__iexact=sku_value)
        .exclude(status__iexact='completed')
        .select_related('job_card')
        .order_by('po_approval_date', 'plan_date', 'created_at', 'id')
    )
    if exclude_job_id:
        qs = qs.exclude(pk=exclude_job_id)
    return qs


def recent_dispatches_for_sku(sku, *, days=RECENT_DISPATCH_DAYS):
    sku_value = (sku or '').strip()
    if not sku_value:
        return []
    cutoff = timezone.localdate() - timedelta(days=days)
    rows = (
        Dispatch.objects.filter(
            is_active=True,
            dispatch_date__gte=cutoff,
            job_card__SKU__iexact=sku_value,
        )
        .select_related('job_card', 'job_card__planning_job')
        .order_by('-dispatch_date', '-id')[:8]
    )
    today = timezone.localdate()
    results = []
    for dispatch in rows:
        job_card = dispatch.job_card
        planning_job = getattr(job_card, 'planning_job', None)
        results.append({
            'jc_number': job_card.job_card_no,
            'po_number': (planning_job.po_number if planning_job else job_card.PO_No) or '-',
            'dispatch_date': dispatch.dispatch_date,
            'days_ago': (today - dispatch.dispatch_date).days,
            'dispatch_qty': dispatch.dispatch_qty,
        })
    return results


def _build_member_row(job, *, current_job_id=None, is_primary=False):
    job_card = getattr(job, 'job_card', None)
    printing_started = _printing_started(job_card)
    return {
        'job_id': job.pk,
        'jc_number': job.jc_number or '-',
        'po_number': job.po_number or '-',
        'status_label': _member_status_label(job, job_card),
        'stage_label': _member_stage_label(job),
        'order_qty': job.order_qty,
        'plate_set_no': job.effective_plate_set_no if hasattr(job, 'effective_plate_set_no') else (job.plate_set_no or ''),
        'printing_started': printing_started,
        'printing_label': 'Printing logged' if printing_started else 'Not started',
        'is_current': current_job_id == job.pk if current_job_id else False,
        'is_primary': is_primary,
        'detail_url_name': 'planning:job_detail',
    }


def _compose_sku_alert_payload(
    *,
    sku,
    cluster_jobs,
    recent_dispatches,
    current_job_id=None,
):
    """Build alert dict for duplicate clusters and/or recent-dispatch low priority."""
    sku_value = (sku or '').strip()
    if not sku_value:
        return None

    cluster_jobs = [job for job in cluster_jobs if _is_active_planning_job(job)]
    is_duplicate_cluster = len(cluster_jobs) >= 2
    if not is_duplicate_cluster and not recent_dispatches:
        return None

    cluster_jobs.sort(key=_job_sort_key)
    primary_job = cluster_jobs[0] if cluster_jobs else None
    members = []
    combine_eligible = []
    for job in cluster_jobs:
        is_primary = primary_job is not None and job.pk == primary_job.pk
        row = _build_member_row(job, current_job_id=current_job_id, is_primary=is_primary)
        members.append(row)
        if not row['printing_started']:
            combine_eligible.append(row['jc_number'])

    combine_possible = is_duplicate_cluster and len(combine_eligible) >= 2
    alert_payload = {
        'sku': sku_value,
        'active_count': len(members) if members else 1,
        'is_duplicate_cluster': is_duplicate_cluster,
        'members': members,
        'primary_jc_number': (primary_job.jc_number if primary_job else '-') or '-',
        'primary_po_number': (primary_job.po_number if primary_job else '-') or '-',
        'combine_possible': combine_possible,
        'combine_eligible_jcs': combine_eligible,
        'recent_dispatches': recent_dispatches,
        'show_priority_hint': bool(recent_dispatches),
        'recent_dispatch_days': RECENT_DISPATCH_DAYS,
    }
    if current_job_id is not None:
        alert_payload['current_job_id'] = current_job_id
    return enrich_sku_duplicate_alert(alert_payload)


def build_sku_duplicate_alert(planning_job):
    """
    Return alert context when multiple active jobs share the SKU and/or the SKU
    was dispatched recently, else None.

    Shown on every matching JC so planners see the signal from any search result.
    """
    if not planning_job or not _is_active_planning_job(planning_job):
        return None

    sku = (planning_job.sku or '').strip()
    if not sku:
        return None

    cluster_jobs = list(active_jobs_for_sku(sku))
    recent_dispatches = recent_dispatches_for_sku(sku)
    alert_payload = _compose_sku_alert_payload(
        sku=sku,
        cluster_jobs=cluster_jobs,
        recent_dispatches=recent_dispatches,
        current_job_id=planning_job.pk,
    )
    if not alert_payload:
        return None
    alert_payload['current_jc_number'] = planning_job.jc_number or '-'
    return alert_payload


def build_sku_duplicate_alert_for_sku(sku, *, current_job=None):
    """Alert by SKU when no planning job is loaded yet (e.g. pending master entry)."""
    sku_value = (sku or '').strip()
    if not sku_value:
        return None
    if current_job:
        return build_sku_duplicate_alert(current_job)

    cluster_jobs = list(active_jobs_for_sku(sku_value))
    recent_dispatches = recent_dispatches_for_sku(sku_value)
    return _compose_sku_alert_payload(
        sku=sku_value,
        cluster_jobs=cluster_jobs,
        recent_dispatches=recent_dispatches,
    )


def attach_sku_duplicate_alerts_to_jobs(jobs):
    """Attach job.sku_duplicate_alert for each job in a page/list (bulk-friendly)."""
    job_list = list(jobs)
    for job in job_list:
        job.sku_duplicate_alert = None

    sku_to_jobs_on_page = {}
    for job in job_list:
        key = _sku_key(job.sku)
        if key:
            sku_to_jobs_on_page.setdefault(key, []).append(job)

    if not sku_to_jobs_on_page:
        return job_list

    unique_skus = list({(job.sku or '').strip() for job in job_list if (job.sku or '').strip()})
    cluster_qs = (
        PlanningJob.objects.filter(is_active=True)
        .exclude(status__iexact='completed')
        .filter(_sku_filter_q(unique_skus))
        .select_related('job_card')
    )
    clusters_by_key = {}
    for job in cluster_qs:
        key = _sku_key(job.sku)
        clusters_by_key.setdefault(key, []).append(job)

    dispatch_cutoff = timezone.localdate() - timedelta(days=RECENT_DISPATCH_DAYS)
    dispatch_sku_q = Q()
    for sku in unique_skus:
        dispatch_sku_q |= Q(job_card__SKU__iexact=sku)
    dispatch_qs = (
        Dispatch.objects.filter(is_active=True, dispatch_date__gte=dispatch_cutoff)
        .filter(dispatch_sku_q)
        .select_related('job_card', 'job_card__planning_job')
        .order_by('-dispatch_date', '-id')
    )
    # Case-insensitive dispatch grouping
    dispatches_by_key = {}
    for dispatch in dispatch_qs:
        sku = (dispatch.job_card.SKU or '').strip()
        key = _sku_key(sku)
        if not key:
            continue
        dispatches_by_key.setdefault(key, [])
        if len(dispatches_by_key[key]) < 8:
            today = timezone.localdate()
            planning_job = getattr(dispatch.job_card, 'planning_job', None)
            dispatches_by_key[key].append({
                'jc_number': dispatch.job_card.job_card_no,
                'po_number': (planning_job.po_number if planning_job else dispatch.job_card.PO_No) or '-',
                'dispatch_date': dispatch.dispatch_date,
                'days_ago': (today - dispatch.dispatch_date).days,
                'dispatch_qty': dispatch.dispatch_qty,
            })

    for key, page_jobs in sku_to_jobs_on_page.items():
        cluster_jobs = clusters_by_key.get(key, [])
        recent_dispatches = dispatches_by_key.get(key, [])
        if len(cluster_jobs) < 2 and not recent_dispatches:
            continue

        base = _compose_sku_alert_payload(
            sku=page_jobs[0].sku,
            cluster_jobs=cluster_jobs,
            recent_dispatches=recent_dispatches,
        )
        if not base:
            continue

        for job in page_jobs:
            alert = {
                **base,
                'current_job_id': job.pk,
                'current_jc_number': job.jc_number or '-',
                'members': [
                    {**member, 'is_current': member['job_id'] == job.pk}
                    for member in base['members']
                ],
            }
            job.sku_duplicate_alert = enrich_sku_duplicate_alert(alert)

    return job_list


def attach_sku_duplicate_alerts_to_job_cards(job_cards):
    """Attach planning_job.sku_duplicate_alert for approval queue job cards."""
    cards = list(job_cards)
    planning_jobs = []
    for card in cards:
        planning_job = getattr(card, 'planning_job', None)
        if planning_job:
            planning_jobs.append(planning_job)
    attach_sku_duplicate_alerts_to_jobs(planning_jobs)
    for card in cards:
        planning_job = getattr(card, 'planning_job', None)
        card.sku_duplicate_alert = getattr(planning_job, 'sku_duplicate_alert', None) if planning_job else None
    return cards


def _active_planning_jobs_queryset():
    return PlanningJob.objects.filter(is_active=True).exclude(status__iexact='completed')


def duplicate_sku_lower_values():
    """Lowercase SKUs that appear on two or more active planning jobs."""
    from django.db.models import Count
    from django.db.models.functions import Lower

    return list(
        _active_planning_jobs_queryset()
        .annotate(sku_lower=Lower('sku'))
        .exclude(sku_lower='')
        .values('sku_lower')
        .annotate(job_count=Count('id'))
        .filter(job_count__gte=2)
        .values_list('sku_lower', flat=True)
    )


def low_priority_sku_lower_values(days=RECENT_DISPATCH_DAYS):
    """Lowercase SKUs with dispatch activity in the recent window."""
    from django.db.models.functions import Lower

    cutoff = timezone.localdate() - timedelta(days=days)
    return list(
        Dispatch.objects.filter(is_active=True, dispatch_date__gte=cutoff)
        .annotate(sku_lower=Lower('job_card__SKU'))
        .exclude(sku_lower='')
        .values_list('sku_lower', flat=True)
        .distinct()
    )


def combine_sku_lower_values():
    """Lowercase SKUs where ≥2 active jobs have no printing logged (combine eligible)."""
    from django.db.models import Count, Exists, OuterRef
    from django.db.models.functions import Lower

    from core.models import Production

    printing_started = Production.objects.filter(
        job_card__planning_job_id=OuterRef('pk'),
        is_active=True,
        entry_type='printing',
    )
    return list(
        _active_planning_jobs_queryset()
        .annotate(
            sku_lower=Lower('sku'),
            has_printing=Exists(printing_started),
        )
        .filter(has_printing=False)
        .exclude(sku_lower='')
        .values('sku_lower')
        .annotate(job_count=Count('id'))
        .filter(job_count__gte=2)
        .values_list('sku_lower', flat=True)
    )


SKU_ALERT_FILTER_KEYS = frozenset({'duplicate', 'combine', 'low_priority', 'attention'})


def _sku_alert_target_values(alert_key):
    alert_key = (alert_key or '').strip().lower()
    if alert_key == 'duplicate':
        return set(duplicate_sku_lower_values())
    if alert_key == 'combine':
        return set(combine_sku_lower_values())
    if alert_key == 'low_priority':
        return set(low_priority_sku_lower_values())
    if alert_key == 'attention':
        return (
            set(duplicate_sku_lower_values())
            | set(combine_sku_lower_values())
            | set(low_priority_sku_lower_values())
        )
    return set()


def job_matches_sku_alert_filter(job, alert_key):
    target_values = _sku_alert_target_values(alert_key)
    if not target_values:
        return False
    sku_lower = (job.sku or '').strip().lower()
    return bool(sku_lower) and sku_lower in target_values


def filter_planning_jobs_by_sku_alert(queryset, alert_key):
    """Filter queryset to duplicate, combine, low-priority SKUs, or all alerts."""
    from django.db.models.functions import Lower

    target_values = _sku_alert_target_values(alert_key)
    if not target_values:
        return queryset.none()
    return queryset.annotate(sku_lower=Lower('sku')).filter(sku_lower__in=target_values)


def count_planning_jobs_by_sku_alert(queryset=None):
    qs = _active_planning_jobs_queryset() if queryset is None else queryset
    return {
        'duplicate': filter_planning_jobs_by_sku_alert(qs, 'duplicate').count(),
        'combine': filter_planning_jobs_by_sku_alert(qs, 'combine').count(),
        'low_priority': filter_planning_jobs_by_sku_alert(qs, 'low_priority').count(),
        'attention': filter_planning_jobs_by_sku_alert(qs, 'attention').count(),
    }


def resolve_sku_alert_kind(alert):
    if not alert:
        return ''
    if alert.get('combine_possible'):
        return 'combine'
    if alert.get('show_priority_hint'):
        return 'low_priority'
    return 'duplicate'


def enrich_sku_duplicate_alert(alert):
    if alert:
        alert['alert_kind'] = resolve_sku_alert_kind(alert)
    return alert


def build_planning_jobs_sku_alert_filter_urls(request, *, active_key=''):
    """Preserve current GET params while toggling sku_alert quick filters."""
    params = request.GET.copy()
    params.pop('page', None)

    def _url_for(key):
        query = params.copy()
        if key and active_key != key:
            query['sku_alert'] = key
        else:
            query.pop('sku_alert', None)
        encoded = query.urlencode()
        return f'?{encoded}' if encoded else '?'

    return {
        'duplicate': _url_for('duplicate'),
        'combine': _url_for('combine'),
        'low_priority': _url_for('low_priority'),
        'attention': _url_for('attention'),
        'clear': _url_for(''),
    }
