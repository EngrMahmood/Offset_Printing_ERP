from __future__ import annotations

from django.db import transaction

from .models import RawMaterialSku, StockTransaction
from .raw_material_sku import resolve_raw_material_sku_for_job_card


def get_raw_material_sku_for_job_card(job_card):
    return resolve_raw_material_sku_for_job_card(job_card)


# Backward-compatible alias
get_supply_chain_item_for_job_card = get_raw_material_sku_for_job_card


def _month_label(value):
    return value.strftime('%B %Y') if value else None


@transaction.atomic
def sync_issuance_from_production(production):
    """Create or update issuance linked to a production entry via job card."""
    if not production.is_active:
        StockTransaction.objects.filter(production=production, source='JOB_CARD').delete()
        return None

    job_card = production.job_card
    if not job_card:
        return None

    raw_sku = get_raw_material_sku_for_job_card(job_card)
    if not raw_sku:
        return None

    consumed_sheets = int(production.output_sheets or 0) + int(production.waste_sheets or 0)
    if consumed_sheets <= 0:
        StockTransaction.objects.filter(production=production, source='JOB_CARD').delete()
        return None

    existing = StockTransaction.objects.filter(production=production, source='JOB_CARD').first()
    is_approved = False
    if existing:
        is_approved = existing.is_approved

    txn, _created = StockTransaction.objects.update_or_create(
        production=production,
        defaults={
            'raw_material_sku': raw_sku,
            'transaction_type': 'ISSUANCE',
            'source': 'JOB_CARD',
            'job_card': job_card,
            'month_str': _month_label(production.date),
            'date': production.date,
            'gin_jc': job_card.job_card_no,
            'sheet_qty_pcs': consumed_sheets,
            'pkt_rim_qty': 0,
            'is_approved': is_approved,
        },
    )
    return txn


def sync_fallback_for_job_card(job_card):
    """If no production logs with sheets exist, generate a pending fallback transaction based on planned sheets."""
    has_real_txns = StockTransaction.objects.filter(
        job_card=job_card,
        production__isnull=False,
        is_active=True
    ).exists()

    if has_real_txns:
        # Delete fallback if actual production runs exist to prevent double counting
        StockTransaction.objects.filter(job_card=job_card, production__isnull=True).delete()
        return False

    raw_sku = get_raw_material_sku_for_job_card(job_card)
    if not raw_sku or not (job_card.total_sheet_quantity and int(job_card.total_sheet_quantity) > 0):
        # Delete fallback if SKU or quantity is invalid/zero
        StockTransaction.objects.filter(job_card=job_card, production__isnull=True).delete()
        return False

    existing = StockTransaction.objects.filter(job_card=job_card, production__isnull=True).first()
    is_approved = False
    if existing:
        is_approved = existing.is_approved

    StockTransaction.objects.update_or_create(
        job_card=job_card,
        production=None,
        defaults={
            'raw_material_sku': raw_sku,
            'transaction_type': 'ISSUANCE',
            'source': 'JOB_CARD',
            'month_str': _month_label(job_card.po_date or timezone.now().date()),
            'date': job_card.po_date or timezone.now().date(),
            'gin_jc': job_card.job_card_no,
            'sheet_qty_pcs': int(job_card.total_sheet_quantity),
            'pkt_rim_qty': 0,
            'is_approved': is_approved,
        }
    )
    return True


@transaction.atomic
def sync_issuance_for_job_card(job_card):
    """Sync issuance rows for all active production records on a job card."""
    synced = 0
    skipped = 0
    productions = job_card.productions.filter(is_active=True).order_by('date', 'id')
    for production in productions:
        txn = sync_issuance_from_production(production)
        if txn:
            synced += 1
        else:
            skipped += 1
            
    # Apply fallback sync if applicable
    if sync_fallback_for_job_card(job_card):
        synced += 1
        
    return synced, skipped


@transaction.atomic
def sync_all_job_card_issuances():
    from core.models import Production, JobCard

    synced = 0
    skipped = 0
    productions = (
        Production.objects
        .filter(is_active=True)
        .select_related('job_card', 'job_card__material', 'job_card__planning_job')
        .order_by('id')
    )
    for production in productions:
        txn = sync_issuance_from_production(production)
        if txn:
            synced += 1
        else:
            skipped += 1
            
    # Apply fallback sync for all active job cards
    for job_card in JobCard.objects.filter(is_active=True).select_related('material', 'planning_job'):
        if sync_fallback_for_job_card(job_card):
            synced += 1
            
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
