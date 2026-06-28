from __future__ import annotations

from django.db import transaction

from .models import StockTransaction, SupplyChainItem


def get_supply_chain_item_for_job_card(job_card):
    if not job_card or not job_card.material_id:
        return None
    return SupplyChainItem.objects.filter(material_id=job_card.material_id).first()


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

    item = get_supply_chain_item_for_job_card(job_card)
    if not item:
        return None

    consumed_sheets = int(production.output_sheets or 0) + int(production.waste_sheets or 0)
    if consumed_sheets <= 0:
        StockTransaction.objects.filter(production=production, source='JOB_CARD').delete()
        return None

    txn, _created = StockTransaction.objects.update_or_create(
        production=production,
        defaults={
            'item': item,
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
        .select_related('job_card', 'job_card__material')
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
        .select_related('material', 'material__supply_chain_details')
        .order_by('-id')[:limit]
    )
    for job_card in job_cards:
        try:
            sc_item = job_card.material.supply_chain_details
        except SupplyChainItem.DoesNotExist:
            sc_item = None

        active_productions = job_card.productions.filter(is_active=True).count()
        synced_issuances = StockTransaction.objects.filter(
            job_card=job_card,
            source='JOB_CARD',
            transaction_type='ISSUANCE',
        ).count()

        rows.append({
            'job_card': job_card,
            'supply_chain_item': sc_item,
            'active_productions': active_productions,
            'synced_issuances': synced_issuances,
            'is_linked': sc_item is not None,
        })
    return rows
