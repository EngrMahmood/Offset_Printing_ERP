from django.contrib.auth.decorators import login_required
from django.shortcuts import render

from .services import get_manual_working_rows
from planning.models import PLANNING_STAGE_CHOICES, PLANNING_STATUS_CHOICES


@login_required
def manual_working_list(request):
    filters = {
        'jc_number': request.GET.get('jc_number', '').strip(),
        'plan_month': request.GET.get('plan_month', '').strip(),
        'customer': request.GET.get('customer', '').strip(),
        'sku': request.GET.get('sku', '').strip(),
        'job_name': request.GET.get('job_name', '').strip(),
        'machine_name': request.GET.get('machine_name', '').strip(),
        'status': request.GET.get('status', '').strip(),
        'planning_stage': request.GET.get('planning_stage', '').strip(),
        'date_from': request.GET.get('date_from', '').strip(),
        'date_to': request.GET.get('date_to', '').strip(),
    }

    rows = get_manual_working_rows(filters)

    return render(request, 'manual_working/manual_working_list.html', {
        'rows': rows,
        'filters': filters,
        'status_choices': PLANNING_STATUS_CHOICES,
        'stage_choices': PLANNING_STAGE_CHOICES,
    })
