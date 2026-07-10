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


def build_sku_duplicate_alert(planning_job):
    """
    Return alert context when multiple active jobs share the SKU, else None.

    Shown on every member JC so planners see the full cluster from any search result.
    """
    if not planning_job or not _is_active_planning_job(planning_job):
        return None

    sku = (planning_job.sku or '').strip()
    if not sku:
        return None

    cluster_jobs = list(active_jobs_for_sku(sku))
    if len(cluster_jobs) < 2:
        return None

    cluster_jobs.sort(key=_job_sort_key)
    primary_job = cluster_jobs[0]
    members = []
    combine_eligible = []
    for job in cluster_jobs:
        is_primary = job.pk == primary_job.pk
        row = _build_member_row(job, current_job_id=planning_job.pk, is_primary=is_primary)
        members.append(row)
        if not row['printing_started']:
            combine_eligible.append(row['jc_number'])

    recent_dispatches = recent_dispatches_for_sku(sku)
    combine_possible = len(combine_eligible) >= 2

    return {
        'sku': sku,
        'active_count': len(members),
        'current_job_id': planning_job.pk,
        'current_jc_number': planning_job.jc_number or '-',
        'members': members,
        'primary_jc_number': primary_job.jc_number or '-',
        'primary_po_number': primary_job.po_number or '-',
        'combine_possible': combine_possible,
        'combine_eligible_jcs': combine_eligible,
        'recent_dispatches': recent_dispatches,
        'show_priority_hint': bool(recent_dispatches),
        'recent_dispatch_days': RECENT_DISPATCH_DAYS,
    }


def build_sku_duplicate_alert_for_sku(sku, *, current_job=None):
    """Alert by SKU when no planning job is loaded yet (e.g. pending master entry)."""
    sku_value = (sku or '').strip()
    if not sku_value:
        return None
    if current_job:
        return build_sku_duplicate_alert(current_job)

    cluster_jobs = list(active_jobs_for_sku(sku_value))
    if len(cluster_jobs) < 2:
        return None
    anchor = cluster_jobs[0]
    return build_sku_duplicate_alert(anchor)


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
        if len(cluster_jobs) < 2:
            continue
        cluster_jobs.sort(key=_job_sort_key)
        primary_job = cluster_jobs[0]
        members = []
        combine_eligible = []
        for cluster_job in cluster_jobs:
            is_primary = cluster_job.pk == primary_job.pk
            row = _build_member_row(cluster_job, is_primary=is_primary)
            members.append(row)
            if not row['printing_started']:
                combine_eligible.append(row['jc_number'])

        recent_dispatches = dispatches_by_key.get(key, [])
        combine_possible = len(combine_eligible) >= 2
        base = {
            'sku': page_jobs[0].sku,
            'active_count': len(members),
            'members': members,
            'primary_jc_number': primary_job.jc_number or '-',
            'primary_po_number': primary_job.po_number or '-',
            'combine_possible': combine_possible,
            'combine_eligible_jcs': combine_eligible,
            'recent_dispatches': recent_dispatches,
            'show_priority_hint': bool(recent_dispatches),
            'recent_dispatch_days': RECENT_DISPATCH_DAYS,
        }
        for job in page_jobs:
            alert = {
                **base,
                'current_job_id': job.pk,
                'current_jc_number': job.jc_number or '-',
                'members': [
                    {**member, 'is_current': member['job_id'] == job.pk}
                    for member in members
                ],
            }
            job.sku_duplicate_alert = alert

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
