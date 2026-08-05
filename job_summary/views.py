from django.db.models import Q
from django.http import HttpResponse
from django.shortcuts import render
from django.urls import reverse
from django.utils import timezone
from django.contrib.auth.decorators import login_required

from core.models import JobCard, JOB_CARD_STATUS_CHOICES, ProductionWipStatus
from core.views import permission_required
from planning.models import PLANNING_STAGE_CHOICES
from reports.export.services import export_as_pdf, export_as_xlsx

COLUMN_OPTIONS = [
    ('job_card_no', 'Job Card No'),
    ('planning_jc_number', 'Planning JC'),
    ('po_number', 'PO Number'),
    ('sku', 'SKU'),
    ('job_name', 'Job Name'),
    ('status', 'Status'),
    ('po_date', 'PO Date'),
    ('delivery_date', 'Delivery Date'),
    ('machine_name', 'Machine'),
    ('department', 'Department'),
    ('order_qty', 'Order Qty'),
    ('total_sheet_quantity', 'Total Sheet Qty'),
    ('total_colors', 'Total Colors'),
    ('plate_set_no', 'Plate Set'),
    ('job_process_type', 'Job Process'),
    ('application', 'Application'),
    ('material', 'Material'),
    ('print_sheet_size', 'Print Sheet'),
    ('purchase_sheet_size', 'Purchase Sheet'),
    ('total_production', 'Total Production Sheets'),
    ('total_packed_pcs', 'Total Packed Pcs'),
    ('total_sorting_waste_pcs', 'Sorting Waste Pcs'),
    ('total_waste', 'Total Waste Sheets'),
    ('total_dispatch', 'Total Dispatched Qty'),
    ('balance_qty', 'Balance Qty'),
    ('dispatch_completion_percent', 'Dispatch %'),
    ('wip_status', 'WIP Status'),
    ('planning_stage', 'Planning Stage'),
    ('created_at', 'Created At'),
]

DEFAULT_SUMMARY_COLUMNS = [
    'job_card_no',
    'planning_jc_number',
    'po_number',
    'sku',
    'job_name',
    'status',
    'machine_name',
    'department',
    'order_qty',
    'planning_stage',
    'wip_status',
    'total_dispatch',
    'balance_qty',
    'dispatch_completion_percent',
    'delivery_date',
]


def _format_date(value):
    if not value:
        return ''
    return value.strftime('%Y-%m-%d')


def _job_value(job, column):
    planning_job = getattr(job, 'planning_job', None)
    if column == 'job_card_no':
        return job.job_card_no or ''
    if column == 'planning_jc_number':
        return getattr(planning_job, 'jc_number', '') or ''
    if column == 'po_number':
        return job.PO_No or getattr(planning_job, 'po_number', '') or ''
    if column == 'sku':
        return job.SKU or getattr(planning_job, 'sku', '') or ''
    if column == 'job_name':
        return getattr(planning_job, 'job_name', '') or ''
    if column == 'status':
        return job.workflow_status or ''
    if column == 'po_date':
        return _format_date(job.po_date)
    if column == 'delivery_date':
        return _format_date(getattr(planning_job, 'delivery_date', None))
    if column == 'machine_name':
        if job.machine_name_id:
            return str(job.machine_name)
        return getattr(planning_job, 'machine_name', '') or ''
    if column == 'department':
        if job.department_id:
            return str(job.department)
        return getattr(planning_job, 'department', '') or ''
    if column == 'order_qty':
        return job.order_qty or ''
    if column == 'total_sheet_quantity':
        if job.total_sheet_quantity is not None:
            return job.total_sheet_quantity
        return getattr(planning_job, 'total_sheet_quantity', '') or ''
    if column == 'total_colors':
        if job.total_colors is not None:
            return job.total_colors
        return getattr(planning_job, 'total_colors', '') or ''
    if column == 'plate_set_no':
        return job.plate_set_no or ''
    if column == 'job_process_type':
        return getattr(planning_job, 'job_process_type', '') or ''
    if column == 'application':
        return job.application or getattr(planning_job, 'application', '') or ''
    if column == 'material':
        if job.material_id:
            return str(job.material)
        return getattr(planning_job, 'material', '') or ''
    if column == 'print_sheet_size':
        return job.print_sheet_size or getattr(planning_job, 'print_sheet_size', '') or ''
    if column == 'purchase_sheet_size':
        return job.purchase_sheet_size or getattr(planning_job, 'purchase_sheet_size', '') or ''
    if column == 'total_production':
        return job.total_production
    if column == 'total_packed_pcs':
        return job.total_packed_pcs
    if column == 'total_sorting_waste_pcs':
        return job.total_sorting_waste_pcs
    if column == 'total_waste':
        return job.total_waste
    if column == 'total_dispatch':
        return job.total_dispatch
    if column == 'balance_qty':
        return job.balance_qty
    if column == 'dispatch_completion_percent':
        return f"{job.dispatch_completion_percent}%"
    if column == 'wip_status':
        return job.wip_status_name
    if column == 'planning_stage':
        if planning_job is not None:
            return planning_job.get_planning_stage_display() if hasattr(planning_job, 'get_planning_stage_display') else (planning_job.planning_stage or '')
        return ''
    if column == 'created_at':
        return timezone.localtime(job.created_at).strftime('%Y-%m-%d %H:%M') if job.created_at else ''
    return getattr(job, column, '') or ''


def _detail_url(job):
    if getattr(job, 'planning_job_id', None):
        return reverse('planning:job_detail', kwargs={'job_id': job.planning_job_id})
    return reverse('job_card_records')


@login_required
@permission_required('can_view_job_summary')
def job_summary_home(request):
    available_columns = [key for key, _ in COLUMN_OPTIONS]
    selected_columns = request.GET.getlist('columns')
    if not selected_columns:
        raw_columns = (request.GET.get('columns') or '').strip()
        if raw_columns:
            selected_columns = [item.strip() for item in raw_columns.split(',') if item.strip()]
    selected_columns = [col for col in selected_columns if col in available_columns]
    if not selected_columns:
        selected_columns = DEFAULT_SUMMARY_COLUMNS.copy()

    job_card_no_filter = (request.GET.get('job_card_no') or '').strip()
    planning_jc_number_filter = (request.GET.get('planning_jc_number') or '').strip()
    po_number_filter = (request.GET.get('po_number') or '').strip()
    sku_filter = (request.GET.get('sku') or '').strip()
    job_name_filter = (request.GET.get('job_name') or '').strip()
    status_filter = [value.strip() for value in request.GET.getlist('status') if value.strip()]
    machine_filter = (request.GET.get('machine') or '').strip()
    department_filter = (request.GET.get('department') or '').strip()
    wip_status_filter = [value.strip() for value in request.GET.getlist('wip_status') if value.strip()]
    planning_stage_filter = [value.strip() for value in request.GET.getlist('planning_stage') if value.strip()]
    from_date = request.GET.get('from_date', '').strip()
    to_date = request.GET.get('to_date', '').strip()

    queryset = JobCard.objects.filter(is_active=True).select_related(
        'planning_job', 'material', 'machine_name', 'department'
    )
    if job_card_no_filter:
        queryset = queryset.filter(job_card_no__icontains=job_card_no_filter)
    if planning_jc_number_filter:
        queryset = queryset.filter(planning_job__jc_number__icontains=planning_jc_number_filter)
    if po_number_filter:
        queryset = queryset.filter(
            Q(PO_No__icontains=po_number_filter)
            | Q(planning_job__po_number__icontains=po_number_filter)
        )
    if sku_filter:
        queryset = queryset.filter(
            Q(SKU__icontains=sku_filter)
            | Q(planning_job__sku__icontains=sku_filter)
        )
    if job_name_filter:
        queryset = queryset.filter(planning_job__job_name__icontains=job_name_filter)
    if status_filter:
        queryset = queryset.filter(status__in=status_filter)
    if machine_filter:
        queryset = queryset.filter(
            Q(machine_name__name__icontains=machine_filter)
            | Q(planning_job__machine_name__icontains=machine_filter)
        )
    if department_filter:
        queryset = queryset.filter(
            Q(department__name__icontains=department_filter)
            | Q(planning_job__department__icontains=department_filter)
        )
    if wip_status_filter:
        queryset = queryset.filter(production_wip_status__status_id__in=wip_status_filter)
    if planning_stage_filter:
        planning_stage_q = Q()
        planning_stage_values = [value for value in planning_stage_filter if value != '__not_set__']
        if planning_stage_values:
            planning_stage_q |= Q(planning_job__planning_stage__in=planning_stage_values)
        if '__not_set__' in planning_stage_filter:
            planning_stage_q |= Q(planning_job__planning_stage='')
        queryset = queryset.filter(planning_stage_q)
    if from_date:
        queryset = queryset.filter(po_date__gte=from_date)
    if to_date:
        queryset = queryset.filter(po_date__lte=to_date)

    jobs = queryset.order_by('-created_at', '-id')[:1000]
    rows = []
    for job in jobs:
        row = {'id': job.id, 'detail_url': _detail_url(job)}
        for column in selected_columns:
            row[column] = _job_value(job, column)
        rows.append(row)

    export_type = (request.GET.get('export') or '').strip().lower()
    if export_type in {'xlsx', 'pdf'}:
        export_rows = []
        for index, row in enumerate(rows, start=1):
            export_row = {'row_number': index}
            export_row.update({col: row[col] for col in selected_columns})
            export_rows.append(export_row)

        payload = {
            'report': {'title': 'Jobs Summary'},
            'generated_at': timezone.localtime().strftime('%Y-%m-%d %H:%M:%S'),
            'data': export_rows,
            'headers': ['row_number', *selected_columns],
            'header_labels': {
                'row_number': '#',
                **{
                    key: label
                    for key, label in COLUMN_OPTIONS
                    if key in selected_columns
                },
            },
        }
        if export_type == 'xlsx':
            content = export_as_xlsx(payload)
            response = HttpResponse(
                content,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            response['Content-Disposition'] = 'attachment; filename="jobs-summary.xlsx"'
            return response
        content = export_as_pdf(payload)
        response = HttpResponse(content, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="jobs-summary.pdf"'
        return response

    params = request.GET.copy()
    params.pop('export', None)
    export_query = params.urlencode()

    return render(
        request,
        'job_summary/home.html',
        {
            'rows': rows,
            'selected_columns': selected_columns,
            'column_options': COLUMN_OPTIONS,
            'status_choices': JOB_CARD_STATUS_CHOICES,
            'wip_status_choices': [
                (str(status.id), status.name)
                for status in ProductionWipStatus.objects.filter(is_active=True).order_by('name')
            ],
            'planning_stage_choices': [
                ('__not_set__', 'Not Set') if value == '' else (value, label)
                for value, label in PLANNING_STAGE_CHOICES
            ],
            'filters': {
                'job_card_no': job_card_no_filter,
                'planning_jc_number': planning_jc_number_filter,
                'po_number': po_number_filter,
                'sku': sku_filter,
                'job_name': job_name_filter,
                'status': status_filter,
                'machine': machine_filter,
                'department': department_filter,
                'wip_status': wip_status_filter,
                'planning_stage': planning_stage_filter,
                'from_date': from_date,
                'to_date': to_date,
            },
            'export_query': export_query,
        },
    )
