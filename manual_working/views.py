from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.shortcuts import render
from django.utils import timezone

from core.models import ProductionWipStatus
from planning.models import PLANNING_STAGE_CHOICES, PLANNING_STATUS_CHOICES
from reports.export.services import export_as_csv, export_as_xlsx

from .services import (
    EXPORT_COLUMNS,
    EXPORT_HEADER_LABELS,
    get_manual_working_export_rows,
    get_manual_working_page,
)


def _manual_working_filters(request):
    return {
        'jc_number': request.GET.get('jc_number', '').strip(),
        'plan_month': request.GET.get('plan_month', '').strip(),
        'customer': request.GET.get('customer', '').strip(),
        'sku': request.GET.get('sku', '').strip(),
        'job_name': request.GET.get('job_name', '').strip(),
        'machine_name': request.GET.get('machine_name', '').strip(),
        'status': request.GET.get('status', '').strip(),
        'wip_status': request.GET.get('wip_status', '').strip(),
        'planning_stage': request.GET.get('planning_stage', '').strip(),
        'date_from': request.GET.get('date_from', '').strip(),
        'date_to': request.GET.get('date_to', '').strip(),
        'release_date_from': request.GET.get('release_date_from', '').strip(),
        'release_date_to': request.GET.get('release_date_to', '').strip(),
    }


@login_required
def manual_working_list(request):
    filters = _manual_working_filters(request)

    export_type = (request.GET.get('export') or '').strip().lower()
    if export_type in {'xlsx', 'csv'}:
        rows = get_manual_working_export_rows(filters)
        payload = {
            'report': {'title': 'Manual Working Transition Board'},
            'generated_at': timezone.localtime().strftime('%Y-%m-%d %H:%M:%S'),
            'data': rows,
            'headers': EXPORT_COLUMNS,
            'header_labels': EXPORT_HEADER_LABELS,
        }
        if export_type == 'xlsx':
            content = export_as_xlsx(payload)
            response = HttpResponse(
                content,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            response['Content-Disposition'] = 'attachment; filename="manual-working.xlsx"'
            return response
        content = export_as_csv(payload)
        response = HttpResponse(content, content_type='text/csv; charset=utf-8')
        response['Content-Disposition'] = 'attachment; filename="manual-working.csv"'
        return response

    page_obj, rows = get_manual_working_page(filters, request.GET.get('page'))

    filter_params = request.GET.copy()
    filter_params.pop('page', None)
    filter_params.pop('export', None)
    filter_query = filter_params.urlencode()

    return render(request, 'manual_working/manual_working_list.html', {
        'rows': rows,
        'page_obj': page_obj,
        'filters': filters,
        'filter_query': filter_query,
        'export_query': filter_query,
        'status_choices': PLANNING_STATUS_CHOICES,
        'wip_statuses': ProductionWipStatus.objects.filter(is_active=True).order_by('name'),
        'stage_choices': PLANNING_STAGE_CHOICES,
    })
