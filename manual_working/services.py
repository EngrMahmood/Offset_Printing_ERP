from datetime import date, datetime

from django.core.exceptions import ObjectDoesNotExist
from django.core.paginator import Paginator
from django.db.models import Q
from django.utils import timezone
from django.utils.dateparse import parse_date

from core.models import ChangeLog
from planning.models import PLANNING_STAGE_CHOICES, PLANNING_STAGE_DONE, PlanningJob, SkuRecipe
from planning.services import _po_dates_from_payload

PAGE_SIZE = 50
EXPORT_LIMIT = 5000

# Display columns for full filtered Excel/CSV export (excludes internal *_obj fields).
EXPORT_COLUMNS = [
    'jc_number', 'month', 'date', 'po_number', 'sku', 'job_name', 'repeat_flag',
    'material', 'color_spec', 'application', 'size_w_mm', 'size_h_mm', 'size_w_inch', 'size_h_inch',
    'order_qty', 'print_pcs', 'ups', 'print_sheet_size', 'print_sheets', 'wastage_sheets',
    'actual_sheet_required', 'purchase_sheet_size', 'purchase_sheet_ups', 'purchase_sheet_required',
    'pkt_value', 'remarks', 'requirement', 'front_colors', 'back_colors', 'total_colors',
    'total_mr_time_minutes', 'front_pass', 'back_pass', 'total_impressions', 'mi_quantity', 'mi_balance',
    'print_date_1', 'print_qty_1', 'wastage_1', 'print_date_2', 'print_qty_2', 'wastage_2',
    'print_date_3', 'print_qty_3', 'wastage_3', 'print_date_4', 'print_qty_4', 'wastage_4',
    'print_date_5', 'print_qty_5', 'wastage_5', 'remaining_sheet', 'wip_status', 'pr_reference',
    'delivery_date_1', 'dc_1', 'delivered_qty_1', 'delivery_date_2', 'dc_2', 'delivered_qty_2',
    'delivery_date_3', 'dc_3', 'delivered_qty_3', 'delivery_date_4', 'dc_4', 'delivered_qty_4',
    'delivery_date_5', 'dc_5', 'delivered_qty_5', 'delivery_date_6', 'dc_6', 'delivered_qty_6',
    'rejected_qty', 'balance_qty', 'destination', 'unit_cost', 'stock_bag', 'machine_name',
    'purchase_material_origin', 'stock_qty', 'planning_stage', 'department', 'plate_set_no',
    'awc_no', 'aging_days', 'die_cutting', 'status',
    'released_to_production_date', 'released_to_production_time',
]

EXPORT_HEADER_LABELS = {
    'jc_number': 'JC No',
    'month': 'Month',
    'date': 'PO Approval Date',
    'po_number': 'PO',
    'sku': 'SKU',
    'job_name': 'Job Name',
    'repeat_flag': 'Repeat',
    'material': 'Material',
    'color_spec': 'Color',
    'application': 'Application',
    'size_w_mm': 'Size W mm',
    'size_h_mm': 'Size H mm',
    'size_w_inch': 'Size W Inch',
    'size_h_inch': 'Size H Inch',
    'order_qty': 'Order Qty',
    'print_pcs': 'Print Pcs',
    'ups': 'Ups',
    'print_sheet_size': 'Print Sheet Size',
    'print_sheets': 'Print Sheets',
    'wastage_sheets': 'Wastage',
    'actual_sheet_required': 'Actual Sheet require',
    'purchase_sheet_size': 'Purchase Sheet Size',
    'purchase_sheet_ups': 'Purchase Sheet ups',
    'purchase_sheet_required': 'Purchase Sheet require',
    'pkt_value': 'PKT',
    'remarks': 'Remarks',
    'requirement': 'Requirement',
    'front_colors': 'No. of Clrs Front',
    'back_colors': 'No. Of Clrs Back',
    'total_colors': 'Total Crls',
    'total_mr_time_minutes': 'Total M/R Time (15m/clr)',
    'front_pass': 'Front Pass',
    'back_pass': 'Back Pass',
    'total_impressions': 'Total Impressions',
    'mi_quantity': 'MI Quantity',
    'mi_balance': 'MI Balance',
    'print_date_1': 'Print Date 1',
    'print_qty_1': 'Print Qty 1',
    'wastage_1': 'Wastage 1',
    'print_date_2': 'Print Date 2',
    'print_qty_2': 'Print Qty 2',
    'wastage_2': 'Wastage 2',
    'print_date_3': 'Print Date 3',
    'print_qty_3': 'Print Qty 3',
    'wastage_3': 'Wastage 3',
    'print_date_4': 'Print Date 4',
    'print_qty_4': 'Print Qty 4',
    'wastage_4': 'Wastage 4',
    'print_date_5': 'Print Date 5',
    'print_qty_5': 'Print Qty 5',
    'wastage_5': 'Wastage 5',
    'remaining_sheet': 'Remaining Sheet',
    'wip_status': 'WIP Status',
    'pr_reference': 'PR',
    'delivery_date_1': 'Delivery Date 1',
    'dc_1': 'DC 1',
    'delivered_qty_1': 'Delivered Qty 1',
    'delivery_date_2': 'Delivery Date 2',
    'dc_2': 'DC 2',
    'delivered_qty_2': 'Delivered Qty 2',
    'delivery_date_3': 'Delivery Date 3',
    'dc_3': 'DC 3',
    'delivered_qty_3': 'Delivered Qty 3',
    'delivery_date_4': 'Delivery Date 4',
    'dc_4': 'DC 4',
    'delivered_qty_4': 'Delivered Qty 4',
    'delivery_date_5': 'Delivery Date 5',
    'dc_5': 'DC 5',
    'delivered_qty_5': 'Delivered Qty 5',
    'delivery_date_6': 'Delivery Date 6',
    'dc_6': 'DC 6',
    'delivered_qty_6': 'Delivered Qty 6',
    'rejected_qty': 'Rejected Qty',
    'balance_qty': 'Balance Qty',
    'destination': 'Destination',
    'unit_cost': 'Unit Cost',
    'stock_bag': 'Stock Bag',
    'machine_name': 'Machine Name',
    'purchase_material_origin': 'Purchase / Material',
    'stock_qty': 'Stock Qty',
    'planning_stage': 'Stage',
    'department': 'Department',
    'plate_set_no': 'Plate Set No',
    'awc_no': 'AWC No',
    'aging_days': 'Aging Days',
    'die_cutting': 'Die Cutting',
    'status': 'Status',
    'released_to_production_date': 'Released to Production Date',
    'released_to_production_time': 'Released to Production Time',
}


def get_manual_working_queryset(filters: dict[str, str]):
    queryset = PlanningJob.objects.filter(is_active=True).select_related(
        'job_card',
        'job_card__production_wip_status__status',
    ).prefetch_related('print_runs', 'dispatch_runs', 'po_documents')

    jc_number = filters.get('jc_number')
    if jc_number:
        queryset = queryset.filter(jc_number__icontains=jc_number)

    plan_month = filters.get('plan_month')
    if plan_month:
        queryset = queryset.filter(plan_month__icontains=plan_month)

    customer = filters.get('customer')
    if customer:
        queryset = queryset.filter(
            Q(po_number__icontains=customer)
            | Q(destination__icontains=customer)
        )

    sku = filters.get('sku')
    if sku:
        queryset = queryset.filter(sku__icontains=sku)

    job_name = filters.get('job_name')
    if job_name:
        queryset = queryset.filter(job_name__icontains=job_name)

    machine_name = filters.get('machine_name')
    if machine_name:
        queryset = queryset.filter(machine_name__icontains=machine_name)

    status = filters.get('status')
    if status:
        queryset = queryset.filter(status=status)

    wip_status = filters.get('wip_status')
    if wip_status:
        queryset = queryset.filter(job_card__production_wip_status__status_id=wip_status)

    planning_stage = filters.get('planning_stage')
    if planning_stage:
        queryset = queryset.filter(planning_stage=planning_stage)

    date_from = parse_date(filters.get('date_from') or '')
    date_to = parse_date(filters.get('date_to') or '')
    if date_from:
        queryset = queryset.filter(po_approval_date__gte=date_from)
    if date_to:
        queryset = queryset.filter(po_approval_date__lte=date_to)

    release_date_from = parse_date(filters.get('release_date_from') or '')
    release_date_to = parse_date(filters.get('release_date_to') or '')
    if release_date_from or release_date_to:
        queryset = _filter_by_release_date(queryset, release_date_from, release_date_to)

    return queryset.order_by('-created_at', '-jc_number')


def get_manual_working_page(filters: dict[str, str], page_number):
    queryset = get_manual_working_queryset(filters)
    paginator = Paginator(queryset, PAGE_SIZE)
    page_obj = paginator.get_page(page_number)
    rows = build_manual_working_rows(list(page_obj.object_list))
    return page_obj, rows


def get_manual_working_export_rows(filters: dict[str, str], *, limit: int = EXPORT_LIMIT) -> list[dict[str, str]]:
    """All filtered active rows for Excel/CSV (capped for safety)."""
    queryset = get_manual_working_queryset(filters)
    return build_manual_working_rows(list(queryset[:limit]))


def build_manual_working_rows(jobs) -> list[dict[str, str]]:
    recipe_map = _build_recipe_map(jobs)
    approval_map, release_map = _build_job_card_log_maps(jobs)
    return [
        _build_row(item, recipe_map, approval_map, release_map)
        for item in jobs
    ]


# Backwards-compatible helper used by older callers/tests.
def get_manual_working_rows(filters: dict[str, str]) -> list[dict[str, str]]:
    return get_manual_working_export_rows(filters, limit=1000)

def _filter_by_release_date(queryset, release_from, release_to):
    """Narrow by ChangeLog release events in range, plus planning_stage_changed_at fallback."""
    log_qs = ChangeLog.objects.filter(entity_type='job_card')
    if release_from:
        log_qs = log_qs.filter(created_at__date__gte=release_from)
    if release_to:
        log_qs = log_qs.filter(created_at__date__lte=release_to)

    released_card_ids = set()
    for log in log_qs.only('record_id', 'action', 'field_changes', 'created_at').iterator(chunk_size=500):
        if log.action == 'release' and log.created_at:
            released_card_ids.add(log.record_id)
            continue
        field_changes = log.field_changes if isinstance(log.field_changes, dict) else {}
        status_change = field_changes.get('status') if isinstance(field_changes, dict) else None
        if isinstance(status_change, dict):
            to_status = str(status_change.get('to') or '').strip().lower()
            if to_status == 'released' and log.created_at:
                released_card_ids.add(log.record_id)

    stage_q = Q(planning_stage=PLANNING_STAGE_DONE)
    if release_from:
        stage_q &= Q(planning_stage_changed_at__date__gte=release_from)
    if release_to:
        stage_q &= Q(planning_stage_changed_at__date__lte=release_to)

    release_q = Q(job_card__id__in=released_card_ids) if released_card_ids else Q(pk__in=[])
    return queryset.filter(release_q | stage_q)


def _build_job_card_log_maps(jobs):
    job_card_ids = [job.job_card.id for job in jobs if getattr(job, 'job_card', None)]
    if not job_card_ids:
        return {}, {}

    approval_map = {}
    release_map = {}

    logs = (
        ChangeLog.objects.filter(
            entity_type='job_card',
            record_id__in=job_card_ids,
        )
        .only('record_id', 'action', 'field_changes', 'created_at')
        .order_by('record_id', 'created_at')
    )

    for log in logs:
        if log.record_id not in release_map:
            field_changes = log.field_changes if isinstance(log.field_changes, dict) else {}
            status_change = field_changes.get('status') if isinstance(field_changes, dict) else None
            if isinstance(status_change, dict):
                to_status = str(status_change.get('to') or '').strip().lower()
                if to_status == 'released' and log.created_at:
                    release_map[log.record_id] = log.created_at
            elif log.action == 'release' and log.created_at:
                release_map[log.record_id] = log.created_at

        field_changes = log.field_changes if isinstance(log.field_changes, dict) else {}
        status_change = field_changes.get('status') if isinstance(field_changes, dict) else None
        if isinstance(status_change, dict):
            to_status = str(status_change.get('to') or '').strip().lower()
            if to_status in {'production_approved', 'qc_approved'} and log.created_at:
                approval_map[log.record_id] = log.created_at.date()

    return approval_map, release_map


def _build_recipe_map(items):
    skus = {(item.sku or '').strip().upper() for item in items if item.sku}
    if not skus:
        return {}
    recipes = SkuRecipe.objects.filter(sku__in=skus)
    return {(recipe.sku or '').strip().upper(): recipe for recipe in recipes}


def _format_text(value):
    if value is None:
        return ''
    return str(value)


def _format_date(value):
    if not value:
        return ''
    return value.strftime('%d-%m-%Y')


def _format_pkt_date(value):
    if not value:
        return ''
    local_dt = timezone.localtime(value)
    return local_dt.strftime('%d-%m-%Y')


def _format_pkt_time(value):
    if not value:
        return ''
    local_dt = timezone.localtime(value)
    return local_dt.strftime('%H:%M:%S')


def _pkt_date(value):
    if not value:
        return None
    return timezone.localtime(value).date()


def _released_to_production_datetime(job: PlanningJob, release_map: dict[int, object]):
    job_card = getattr(job, 'job_card', None)
    if job_card and job_card.id in release_map:
        return release_map[job_card.id]
    if (
        job_card
        and job_card.status in {'released', 'in_production', 'completed', 'closed'}
        and job.planning_stage == PLANNING_STAGE_DONE
        and job.planning_stage_changed_at
    ):
        return job.planning_stage_changed_at
    return None


def _format_month(value, fallback_date=None):
    if isinstance(value, (datetime, date)):
        return value.strftime('%B')

    month_text = _format_text(value).strip()
    if month_text:
        month_text_lower = month_text.lower()
        try:
            month_int = int(month_text)
            if 1 <= month_int <= 12:
                return datetime(2000, month_int, 1).strftime('%B')
        except Exception:
            pass
        month_names = {datetime(2000, m, 1).strftime('%B').lower(): datetime(2000, m, 1).strftime('%B') for m in range(1, 13)}
        month_abbr = {datetime(2000, m, 1).strftime('%b').lower(): datetime(2000, m, 1).strftime('%B') for m in range(1, 13)}
        if month_text_lower in month_names:
            return month_names[month_text_lower]
        if month_text_lower in month_abbr:
            return month_abbr[month_text_lower]
        return month_text.title()
    if fallback_date:
        return fallback_date.strftime('%B')
    return ''


def _resolve_job_or_recipe(job_value, recipe_value=None):
    job_text = _format_text(job_value).strip()
    if job_text:
        return job_text
    return _format_text(recipe_value)


def _resolve_repeat_flag(job_value, recipe):
    return _format_text(job_value).strip()


def _po_approval_date(job: PlanningJob, approval_map: dict[int, object]):
    """Resolve approval date from job fields / prefetched docs / batch log map — no extra queries."""
    if getattr(job, 'po_approval_date', None):
        return job.po_approval_date

    docs = list(job.po_documents.all())
    docs.sort(key=lambda doc: doc.created_at or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
    for po_document in docs:
        approval_date, po_date = _po_dates_from_payload(po_document.extracted_payload or {})
        if approval_date:
            return approval_date
        if po_date:
            return po_date

    job_card = getattr(job, 'job_card', None)
    if not job_card:
        return None
    return approval_map.get(job_card.id)


def _month_label_for_job(job: PlanningJob, po_approval_date):
    """Derive the month label from the system-entry date (created_at).

    plan_month and plan_date on PlanningJob can hold scheduled/future dates from
    CSV imports (e.g. the 'Month' and 'Date' columns in the planning sheet), so
    they must NOT be used as the primary source for this label.
    Priority: created_at → po_approval_date → plan_date (last resort).
    """
    # 1. Prefer the actual system-entry timestamp
    if job.created_at:
        entry_date = timezone.localtime(job.created_at).date()
        return _format_month('', fallback_date=entry_date)
    # 2. PO approval date
    if po_approval_date:
        return _format_month('', fallback_date=po_approval_date)
    # 3. Last resort: stored plan_date (may be a future CSV-imported date)
    fallback = job.plan_date
    if fallback:
        return _format_month('', fallback_date=fallback)
    return ''


def _wip_status_for_job(job: PlanningJob) -> str:
    job_card = getattr(job, 'job_card', None)
    if not job_card:
        return 'Not Set'
    try:
        status_name = job_card.wip_status_name
    except ObjectDoesNotExist:
        status_name = 'Not Set'
    if status_name == 'Not Set' and job_card.status == 'released':
        return 'Printing'
    return status_name or 'Not Set'


def _build_print_run_values(job: PlanningJob) -> list[dict[str, str]]:
    results = []
    runs = list(job.print_runs.all())[:5]
    for index in range(5):
        run = runs[index] if index < len(runs) else None
        results.append({
            'print_date': _format_date(getattr(run, 'print_date', None) if run else None),
            'print_qty': _format_text(getattr(run, 'print_qty', '') if run else ''),
            'wastage': _format_text(getattr(run, 'wastage_qty', '') if run else ''),
        })
    return results


def _build_dispatch_values(job: PlanningJob) -> list[dict[str, str]]:
    results = []
    runs = list(job.dispatch_runs.all())[:6]
    for index in range(6):
        run = runs[index] if index < len(runs) else None
        results.append({
            'delivery_date': _format_date(getattr(run, 'delivery_date', None) if run else None),
            'dc_no': _format_text(getattr(run, 'dc_no', '') if run else ''),
            'qty': _format_text(getattr(run, 'delivered_qty', '') if run else ''),
        })
    return results


def _build_row(job: PlanningJob, recipe_map: dict[str, SkuRecipe], approval_map: dict[int, object], release_map: dict[int, object]) -> dict[str, str]:
    recipe = recipe_map.get((job.sku or '').strip().upper())
    print_runs = _build_print_run_values(job)
    dispatch_values = _build_dispatch_values(job)

    po_approval_date = _po_approval_date(job, approval_map)
    released_at = _released_to_production_datetime(job, release_map)

    return {
        'jc_number': _format_text(job.jc_number),
        'month': _month_label_for_job(job, po_approval_date),
        'date': _format_date(po_approval_date),
        'po_approval_date_obj': po_approval_date,
        'po_number': _format_text(job.po_number),
        'sku': _format_text(job.sku),
        'job_name': _format_text(job.job_name),
        'repeat_flag': _resolve_repeat_flag(job.repeat_flag, recipe),
        'planning_stage': _format_text(dict(PLANNING_STAGE_CHOICES).get(job.planning_stage, job.planning_stage)),
        'material': _resolve_job_or_recipe(job.material, getattr(recipe, 'material', None)),
        'color_spec': _resolve_job_or_recipe(job.color_spec, getattr(recipe, 'color_spec', None)),
        'application': _resolve_job_or_recipe(job.application, getattr(recipe, 'application', None)),
        'size_w_mm': _resolve_job_or_recipe(job.size_w_mm, getattr(recipe, 'size_w_mm', None)),
        'size_h_mm': _resolve_job_or_recipe(job.size_h_mm, getattr(recipe, 'size_h_mm', None)),
        'size_w_inch': _format_text(job.size_w_inch),
        'size_h_inch': _format_text(job.size_h_inch),
        'order_qty': _format_text(job.order_qty),
        'print_pcs': _format_text(job.print_pcs),
        'ups': _resolve_job_or_recipe(job.ups, getattr(recipe, 'ups', None)),
        'print_sheet_size': _resolve_job_or_recipe(job.print_sheet_size, getattr(recipe, 'print_sheet_size', None)),
        'print_sheets': _format_text(job.print_sheets),
        'wastage_sheets': _format_text(job.wastage_sheets),
        'actual_sheet_required': _format_text(job.actual_sheet_required),
        'purchase_sheet_size': _resolve_job_or_recipe(job.purchase_sheet_size, getattr(recipe, 'purchase_sheet_size', None)),
        'purchase_sheet_ups': _resolve_job_or_recipe(job.purchase_sheet_ups, getattr(recipe, 'purchase_sheet_ups', None)),
        'purchase_sheet_required': _format_text(job.purchase_sheet_required),
        'pkt_value': _format_text(job.pkt_value),
        'remarks': _format_text(job.remarks),
        'requirement': _format_text(job.requirement),
        'front_colors': _format_text(job.front_colors),
        'back_colors': _format_text(job.back_colors),
        'total_colors': _format_text(job.total_colors),
        'total_mr_time_minutes': _format_text(job.total_mr_time_minutes),
        'front_pass': _format_text(job.front_pass),
        'back_pass': _format_text(job.back_pass),
        'total_impressions': _format_text(job.planned_total_impressions),
        'mi_quantity': _format_text(job.mi_quantity),
        'mi_balance': _format_text(job.mi_balance),
        'print_date_1': print_runs[0]['print_date'],
        'print_qty_1': print_runs[0]['print_qty'],
        'wastage_1': print_runs[0]['wastage'],
        'print_date_2': print_runs[1]['print_date'],
        'print_qty_2': print_runs[1]['print_qty'],
        'wastage_2': print_runs[1]['wastage'],
        'print_date_3': print_runs[2]['print_date'],
        'print_qty_3': print_runs[2]['print_qty'],
        'wastage_3': print_runs[2]['wastage'],
        'print_date_4': print_runs[3]['print_date'],
        'print_qty_4': print_runs[3]['print_qty'],
        'wastage_4': print_runs[3]['wastage'],
        'print_date_5': print_runs[4]['print_date'],
        'print_qty_5': print_runs[4]['print_qty'],
        'wastage_5': print_runs[4]['wastage'],
        'remaining_sheet': _format_text(job.remaining_sheet),
        'status': _format_text(job.status),
        'pr_reference': _format_text(job.pr_reference),
        'delivery_date_1': dispatch_values[0]['delivery_date'],
        'dc_1': dispatch_values[0]['dc_no'],
        'delivered_qty_1': dispatch_values[0]['qty'],
        'delivery_date_2': dispatch_values[1]['delivery_date'],
        'dc_2': dispatch_values[1]['dc_no'],
        'delivered_qty_2': dispatch_values[1]['qty'],
        'delivery_date_3': dispatch_values[2]['delivery_date'],
        'dc_3': dispatch_values[2]['dc_no'],
        'delivered_qty_3': dispatch_values[2]['qty'],
        'delivery_date_4': dispatch_values[3]['delivery_date'],
        'dc_4': dispatch_values[3]['dc_no'],
        'delivered_qty_4': dispatch_values[3]['qty'],
        'delivery_date_5': dispatch_values[4]['delivery_date'],
        'dc_5': dispatch_values[4]['dc_no'],
        'delivered_qty_5': dispatch_values[4]['qty'],
        'delivery_date_6': dispatch_values[5]['delivery_date'],
        'dc_6': dispatch_values[5]['dc_no'],
        'delivered_qty_6': dispatch_values[5]['qty'],
        'rejected_qty': _format_text(job.rejected_qty),
        'balance_qty': _format_text(job.balance_qty),
        'destination': _format_text(job.destination),
        'unit_cost': _format_text(job.unit_cost),
        'stock_bag': _format_text(job.stock_bag),
        'machine_name': _format_text(job.machine_name),
        'purchase_material_origin': _format_text(job.purchase_material_origin),
        'stock_qty': _format_text(job.stock_qty),
        'free': '',
        'department': _format_text(job.department),
        'plate_set_no': _format_text(job.plate_set_no),
        'awc_no': _resolve_job_or_recipe(getattr(job, 'awc_no', ''), getattr(recipe, 'awc_no', None)),
        'aging_days': _format_text(job.aging_days),
        'die_cutting': _resolve_job_or_recipe(getattr(job, 'die_cutting', ''), getattr(recipe, 'die_cutting', None)),
        'wip_status': _format_text(_wip_status_for_job(job)),
        'released_to_production_date': _format_pkt_date(released_at),
        'released_to_production_time': _format_pkt_time(released_at),
        'released_to_production_date_obj': _pkt_date(released_at),
    }
