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
        },
    )
    return txn


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
    return synced, skipped


@transaction.atomic
def sync_all_job_card_issuances():
    from core.models import Production

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
