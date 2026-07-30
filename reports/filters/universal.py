from __future__ import annotations


def parse_universal_filters(request) -> dict:
    return {
        'date_from': (request.GET.get('date_from') or '').strip(),
        'date_to': (request.GET.get('date_to') or '').strip(),
        'department': (request.GET.get('department') or '').strip(),
        'customer': (request.GET.get('customer') or '').strip(),
        'sku': (request.GET.get('sku') or '').strip(),
        'po': (request.GET.get('po') or '').strip(),
        'job_card': (request.GET.get('job_card') or '').strip(),
        'machine': (request.GET.get('machine') or '').strip(),
        'shift': (request.GET.get('shift') or '').strip(),
        'operator': (request.GET.get('operator') or '').strip(),
        'material': (request.GET.get('material') or '').strip(),
        'status': (request.GET.get('status') or '').strip(),
        'destination': (request.GET.get('destination') or '').strip(),
        'planner': (request.GET.get('planner') or '').strip(),
        'production_line': (request.GET.get('production_line') or '').strip(),
        'product_category': (request.GET.get('product_category') or '').strip(),
        'approval_status': (request.GET.get('approval_status') or '').strip(),
        'wastage_status': (request.GET.get('wastage_status') or '').strip(),
        'period': (request.GET.get('period') or '').strip(),
        # Selects which table a multi-table report renders/exports; part of the
        # cache key so per-tab exports don't collide.
        'tab': (request.GET.get('tab') or '').strip(),
        'page': (request.GET.get('page') or '').strip(),
        'high_wastage': (request.GET.get('high_wastage') or '').strip(),
        # KPI Scorecard's month/quarter selector — without these in the cache
        # key, every period selection collided on the same cached payload.
        'period_type': (request.GET.get('period_type') or '').strip(),
        'year': (request.GET.get('year') or '').strip(),
        'month': (request.GET.get('month') or '').strip(),
        'quarter': (request.GET.get('quarter') or '').strip(),
        # Pending Work's per-stage export selector.
        'stage': (request.GET.get('stage') or '').strip(),
        # KPI Scorecard drill-down export selector (which KPI's supporting data).
        'kpi': (request.GET.get('kpi') or '').strip(),
        # KPI Scorecard quarterly-detail export selector.
        'detail': (request.GET.get('detail') or '').strip(),
    }
