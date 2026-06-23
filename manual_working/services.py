from datetime import datetime

from django.core.exceptions import ObjectDoesNotExist
from django.db.models import Q
from django.utils.dateparse import parse_date

from core.models import ChangeLog, JobCard
from planning.models import PLANNING_STAGE_CHOICES, PlanningJob, PoDocument, SkuRecipe


def get_manual_working_rows(filters: dict[str, str]) -> list[dict[str, str]]:
    queryset = PlanningJob.objects.all().select_related(
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

    queryset = queryset.order_by('-plan_date', '-jc_number')
    jobs = list(queryset)
    recipe_map = _build_recipe_map(jobs)
    approval_map = _build_job_card_approval_date_map(jobs)
    wip_status_map = _build_wip_status_map(jobs)

    rows = [_build_row(item, recipe_map, approval_map, wip_status_map) for item in jobs]

    date_from = filters.get('date_from')
    if date_from:
        parsed_from = parse_date(date_from)
        if parsed_from:
            rows = [row for row in rows if row.get('po_approval_date_obj') and row['po_approval_date_obj'] >= parsed_from]

    date_to = filters.get('date_to')
    if date_to:
        parsed_to = parse_date(date_to)
        if parsed_to:
            rows = [row for row in rows if row.get('po_approval_date_obj') and row['po_approval_date_obj'] <= parsed_to]

    return rows


def _build_job_card_approval_date_map(jobs):
    job_card_ids = [job.job_card.id for job in jobs if getattr(job, 'job_card', None)]
    if not job_card_ids:
        return {}

    approval_map = {}
    logs = ChangeLog.objects.filter(entity_type='job_card', record_id__in=job_card_ids).order_by('record_id', '-created_at')
    for log in logs:
        if log.record_id in approval_map:
            continue
        field_changes = log.field_changes if isinstance(log.field_changes, dict) else {}
        status_change = field_changes.get('status') if isinstance(field_changes, dict) else None
        if not isinstance(status_change, dict):
            continue
        to_status = str(status_change.get('to') or '').strip().lower()
        if to_status in {'production_approved', 'qc_approved'} and log.created_at:
            approval_map[log.record_id] = log.created_at.date()
    return approval_map


def _build_recipe_map(items):
    skus = {(item.sku or '').strip().upper() for item in items if item.sku}
    recipes = SkuRecipe.objects.filter(sku__in=skus)
    return {(recipe.sku or '').strip().upper(): recipe for recipe in recipes}


def _build_wip_status_map(jobs):
    jc_numbers = {(job.jc_number or '').strip() for job in jobs if job.jc_number}
    if not jc_numbers:
        return {}

    job_cards = JobCard.objects.filter(
        job_card_no__in=jc_numbers,
        is_active=True,
    ).select_related('production_wip_status__status')

    status_map = {}
    for job_card in job_cards:
        status_name = job_card.wip_status_name
        if status_name == 'Not Set' and job_card.status == 'released':
            status_name = 'Printing'
        status_map[job_card.job_card_no] = status_name
    return status_map


def _format_text(value):
    if value is None:
        return ''
    return str(value)


def _format_date(value):
    if not value:
        return ''
    return value.strftime('%d-%m-%Y')


def _format_month(value, fallback_date=None):
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
    job_text = _format_text(job_value).strip()
    if job_text:
        return job_text
    if recipe:
        return 'Repeat'
    return ''


def _po_approval_date(job: PlanningJob, approval_map: dict[int, object]):
    if getattr(job, 'po_approval_date', None):
        return job.po_approval_date

    if job.po_number:
        po_doc = PoDocument.objects.filter(
            extracted_payload__po_number__iexact=job.po_number,
            extraction_status='processed',
        ).order_by('-created_at').first()
        if po_doc:
            payload = po_doc.extracted_payload or {}
            approval_date = parse_date(payload.get('approval_date'))
            if approval_date:
                return approval_date

    job_card = getattr(job, 'job_card', None)
    if not job_card:
        return None
    return approval_map.get(job_card.id)


def _customer_from_po_documents(job: PlanningJob) -> str:
    for document in job.po_documents.all():
        payload = document.extracted_payload or {}
        if isinstance(payload, dict):
            for key in ('customer', 'Customer', 'customer_name', 'ship_to'):
                candidate = payload.get(key)
                if candidate:
                    return str(candidate)
    return ''


def _build_print_run_values(job: PlanningJob) -> list[dict[str, str]]:
    results = []
    runs = list(job.print_runs.all()[:5])
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
    runs = list(job.dispatch_runs.all()[:6])
    for index in range(6):
        run = runs[index] if index < len(runs) else None
        results.append({
            'delivery_date': _format_date(getattr(run, 'delivery_date', None) if run else None),
            'dc_no': _format_text(getattr(run, 'dc_no', '') if run else ''),
            'qty': _format_text(getattr(run, 'delivered_qty', '') if run else ''),
        })
    return results


def _build_row(job: PlanningJob, recipe_map: dict[str, SkuRecipe], approval_map: dict[int, object], wip_status_map: dict[str, str]) -> dict[str, str]:
    recipe = recipe_map.get((job.sku or '').strip().upper())
    print_runs = _build_print_run_values(job)
    dispatch_values = _build_dispatch_values(job)
    total_delivered_qty = sum((getattr(run, 'delivered_qty') or 0) for run in job.dispatch_runs.all())
    job_card = getattr(job, 'job_card', None)

    po_approval_date = _po_approval_date(job, approval_map)
    try:
        job_card_month = job_card.month if job_card else None
    except ObjectDoesNotExist:
        job_card_month = None

    display_month_source = None if po_approval_date else (job_card_month or job.plan_month)

    return {
        'jc_number': _format_text(job.jc_number),
        'month': _format_month(display_month_source, po_approval_date),
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
        'wip_status': _format_text(wip_status_map.get((job.jc_number or '').strip(), 'Not Set')),
    }
