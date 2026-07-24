from __future__ import annotations

import math

from django.db import transaction
from django.utils import timezone

from .models import RawMaterialSku, StockTransaction
from .raw_material_sku import resolve_raw_material_sku_for_job_card


def get_raw_material_sku_for_job_card(job_card):
    return resolve_raw_material_sku_for_job_card(job_card)


# Backward-compatible alias
get_supply_chain_item_for_job_card = get_raw_material_sku_for_job_card


def _month_label(value):
    return value.strftime('%B %Y') if value else None


def _is_released_to_production(job_card):
    """True only once the planner has released the JC to production.

    This mirrors the exact criterion the production module uses to decide which
    job cards flow from planning into production (see
    production/printing_entry_helpers.py and production/released_jobs.py, which
    filter on ``status__in=JOB_CARD_PRODUCTION_START_STATUSES``). Statuses before
    release (planning/QC/PM approval, production_approved) must NOT appear in
    issuance — material is issued only when the job actually goes to production.
    """
    from core.models import JOB_CARD_PRODUCTION_START_STATUSES

    return job_card.workflow_status in JOB_CARD_PRODUCTION_START_STATUSES


def _planned_sheet_qty(job_card):
    """Planned PRESS sheets drawn for the job (required + wastage)."""
    try:
        planned = int(job_card.total_sheets_planned or 0)
    except (TypeError, ValueError):
        planned = 0
    return planned


def _purchase_sheet_ups(job_card):
    """Press sheets obtained from one purchase sheet (for unit conversion)."""
    ups = job_card.purchase_sheet_ups
    if not ups:
        planning_job = getattr(job_card, 'planning_job', None)
        if planning_job is not None:
            ups = getattr(planning_job, 'purchase_sheet_ups_display', None)
    try:
        ups = int(ups or 0)
    except (TypeError, ValueError):
        ups = 0
    return ups if ups > 0 else None


def _planned_purchase_sheet_qty(job_card):
    """Planned issuance quantity expressed in PURCHASE sheets.

    Supply chain works entirely in purchase sheets (SKU purchase size,
    opening/receiving/on-hand, demand gap). Job cards plan in press sheets, so
    convert: purchase_sheets = ceil(press_sheets / purchase_sheet_ups). When the
    purchase-sheet UPS is unknown we cannot convert, so fall back to the raw
    planned quantity.
    """
    planned_press = _planned_sheet_qty(job_card)
    if planned_press <= 0:
        return 0
    ups = _purchase_sheet_ups(job_card)
    if ups:
        return math.ceil(planned_press / ups)
    return planned_press


@transaction.atomic
def sync_issuance_for_job_card_single(job_card):
    """Maintain a single issuance row per job card.

    Material is issued once for the job, based on the planned JC sheet quantity
    (required + wastage). Multiple print passes/runs reprint the same physical
    sheets, so we do NOT create one row per production run (that double-counts
    paper) — the single planned-quantity row stands regardless of how many runs
    are logged.
    """
    if not job_card:
        return None

    def _clear_pending():
        # Drop only unapproved rows. Approved issuance is a real stock deduction;
        # once a job is completed/closed (dropping out of the production queue) or
        # un-released, we must not silently reverse it.
        StockTransaction.objects.filter(
            job_card=job_card, source='JOB_CARD', is_approved=False
        ).delete()

    raw_sku = get_raw_material_sku_for_job_card(job_card)
    planned_qty = _planned_purchase_sheet_qty(job_card)

    # Only issue for JCs released to production — the same set of jobs the
    # production module shows (released / in_production).
    if not raw_sku or planned_qty <= 0 or not _is_released_to_production(job_card):
        _clear_pending()
        return None

    # Remove any legacy per-production rows for this JC; we keep exactly one row.
    StockTransaction.objects.filter(
        job_card=job_card, source='JOB_CARD', production__isnull=False
    ).delete()

    existing = StockTransaction.objects.filter(
        job_card=job_card, source='JOB_CARD', production__isnull=True
    ).first()
    is_approved = existing.is_approved if existing else False
    issue_date = job_card.po_date or timezone.now().date()

    txn, _created = StockTransaction.objects.update_or_create(
        job_card=job_card,
        production=None,
        source='JOB_CARD',
        transaction_type='ISSUANCE',
        defaults={
            'raw_material_sku': raw_sku,
            'month_str': _month_label(issue_date),
            'date': issue_date,
            'gin_jc': job_card.job_card_no,
            'sheet_qty_pcs': planned_qty,
            'pkt_rim_qty': 0,
            'is_approved': is_approved,
        },
    )
    return txn


def sync_issuance_from_production(production):
    """Signal entry point: re-sync the single issuance row for the production's JC."""
    job_card = getattr(production, 'job_card', None)
    if not job_card:
        return None
    return sync_issuance_for_job_card_single(job_card)


# Backward-compatible alias
sync_issuance_for_job_card = sync_issuance_for_job_card_single


@transaction.atomic
def sync_all_job_card_issuances():
    from core.models import JobCard

    synced = 0
    skipped = 0
    for job_card in JobCard.objects.filter(is_active=True).select_related('material', 'planning_job'):
        if sync_issuance_for_job_card_single(job_card):
            synced += 1
        else:
            skipped += 1

    return synced, skipped


def build_job_card_link_rows(limit=100):
    from core.models import JobCard

    rows = []
    job_cards = (
        JobCard.objects
        .filter(is_active=True, material__isnull=False)
        .select_related('material', 'planning_job')
        .order_by('-id')[:limit]
    )
    for job_card in job_cards:
        raw_sku = get_raw_material_sku_for_job_card(job_card)

        active_productions = job_card.productions.filter(is_active=True).count()
        synced_issuances = StockTransaction.objects.filter(
            job_card=job_card,
            source='JOB_CARD',
            transaction_type='ISSUANCE',
        ).count()

        rows.append({
            'job_card': job_card,
            'supply_chain_item': raw_sku,
            'raw_material_sku': raw_sku,
            'active_productions': active_productions,
            'synced_issuances': synced_issuances,
            'is_linked': raw_sku is not None,
        })
    return rows
