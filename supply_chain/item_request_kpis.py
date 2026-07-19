from django.db.models import Count
from django.utils import timezone

from .models import ItemRequest


def _days_between(start, end):
    if not start or not end:
        return None
    return (end - start).days


def _avg(values):
    values = [v for v in values if v is not None]
    if not values:
        return None
    return round(sum(values) / len(values), 1)


def build_item_request_kpi_rows(queryset=None):
    """One enriched row per request with cycle-time and variance metrics."""
    if queryset is None:
        queryset = ItemRequest.objects.filter(is_active=True)

    qs = queryset.select_related('request_type', 'department', 'procurement').prefetch_related('approvals')

    rows = []
    today = timezone.now().date()

    for req in qs:
        approvals = list(req.approvals.all())
        submitted_at = next((a.created_at for a in approvals if a.action == 'SUBMIT'), req.created_at)
        approved_at = next(
            (a.created_at for a in reversed(approvals) if a.action == 'APPROVE' and a.stage == 'SUPPLY_CHAIN'),
            None,
        )
        procurement = getattr(req, 'procurement', None)

        approval_cycle_days = _days_between(
            submitted_at.date() if submitted_at else None,
            approved_at.date() if approved_at else None,
        )

        code_opening_lead_days = None
        pr_turnaround_days = None
        pr_po_cycle_days = None
        po_receipt_lead_days = None
        total_fulfilment_days = None
        price_variance = None
        on_time = None

        if procurement:
            approved_date = approved_at.date() if approved_at else None
            code_opening_lead_days = _days_between(approved_date, procurement.code_opened_date)
            pr_turnaround_days = _days_between(procurement.code_opened_date, procurement.pr_date)
            pr_po_cycle_days = _days_between(procurement.pr_date, procurement.po_date)
            po_receipt_lead_days = _days_between(procurement.po_date, procurement.received_date)
            total_fulfilment_days = _days_between(
                submitted_at.date() if submitted_at else None,
                procurement.received_date,
            )
            if procurement.unit_price is not None and req.estimated_unit_price is not None:
                price_variance = procurement.unit_price - req.estimated_unit_price

        age_days = (today - req.created_at.date()).days if req.is_open else None

        rows.append({
            'request': req,
            'approval_cycle_days': approval_cycle_days,
            'code_opening_lead_days': code_opening_lead_days,
            'pr_turnaround_days': pr_turnaround_days,
            'pr_po_cycle_days': pr_po_cycle_days,
            'po_receipt_lead_days': po_receipt_lead_days,
            'total_fulfilment_days': total_fulfilment_days,
            'price_variance': price_variance,
            'age_days': age_days,
        })

    return rows


def build_item_request_kpi_summary(rows, queryset=None):
    if queryset is None:
        queryset = ItemRequest.objects.filter(is_active=True)

    total = queryset.count()
    open_count = sum(1 for r in queryset if r.is_open)
    closed_count = total - open_count
    received_count = queryset.filter(status__in=['RECEIVED', 'CLOSED']).count()

    on_time_count = 0
    on_time_eligible = 0
    for row in rows:
        req = row['request']
        procurement = getattr(req, 'procurement', None)
        if procurement and procurement.received_date and row['total_fulfilment_days'] is not None:
            on_time_eligible += 1

    aging_buckets = {'0_7': 0, '8_14': 0, '15_30': 0, '30_plus': 0}
    for row in rows:
        age = row['age_days']
        if age is None:
            continue
        if age <= 7:
            aging_buckets['0_7'] += 1
        elif age <= 14:
            aging_buckets['8_14'] += 1
        elif age <= 30:
            aging_buckets['15_30'] += 1
        else:
            aging_buckets['30_plus'] += 1

    by_type = list(
        queryset.values('request_type__name').annotate(count=Count('id')).order_by('-count')
    )
    by_department = list(
        queryset.values('department__name').annotate(count=Count('id')).order_by('-count')
    )

    summary = {
        'total': total,
        'open_count': open_count,
        'closed_count': closed_count,
        'received_count': received_count,
        'avg_approval_cycle_days': _avg([r['approval_cycle_days'] for r in rows]),
        'avg_code_opening_lead_days': _avg([r['code_opening_lead_days'] for r in rows]),
        'avg_pr_turnaround_days': _avg([r['pr_turnaround_days'] for r in rows]),
        'avg_pr_po_cycle_days': _avg([r['pr_po_cycle_days'] for r in rows]),
        'avg_po_receipt_lead_days': _avg([r['po_receipt_lead_days'] for r in rows]),
        'avg_total_fulfilment_days': _avg([r['total_fulfilment_days'] for r in rows]),
        'avg_price_variance': _avg([r['price_variance'] for r in rows]),
        'aging_buckets': aging_buckets,
        'by_type': by_type,
        'by_department': by_department,
    }
    return summary


def build_item_request_kpi_data(queryset=None):
    if queryset is None:
        queryset = ItemRequest.objects.filter(is_active=True)
    rows = build_item_request_kpi_rows(queryset)
    summary = build_item_request_kpi_summary(rows, queryset)
    return rows, summary
