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
    }
