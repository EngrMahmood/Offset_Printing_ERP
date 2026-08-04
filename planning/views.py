import csv
import logging
import base64
import io
import json
import re
from difflib import SequenceMatcher
from datetime import date, datetime


def _build_qr_image_base64(data):
    try:
        import qrcode
    except ImportError:
        return None

    buffer = io.BytesIO()
    qr = qrcode.QRCode(border=1, box_size=3)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white')
    img.save(buffer, format='PNG')
    return base64.b64encode(buffer.getvalue()).decode('utf-8')
from decimal import Decimal, InvalidOperation
from urllib.parse import urlencode

from django.contrib import messages
from django.contrib.auth.decorators import login_required

logger = logging.getLogger(__name__)
from django.core.exceptions import ValidationError
from django.core.files.base import ContentFile
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Count, Prefetch, Q, Sum
from django.db.models.deletion import ProtectedError
from django.http import Http404, HttpResponse
from django.utils import timezone
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

from core.jc_numbering import allocate_next_jc_number
from core.models import ChangeLog, JobCard, Machine
from core.jobcard_service import job_card_queue_queryset, execute_job_card_action
from core.views import permission_required
from reports.export.services import export_as_pdf, export_as_xlsx
from .merge_engine import (
    MergeConfig, allocate_ups, bucket_signature, build_suggestions, compute_savings,
    merge_blockers, normalise_material, size_key,
)
from .forms import PlanningJobEditForm, PlanningJobFinalizationForm, SkuRecipeForm
from .services import (
    _user_is_admin, _planning_status_filter_values, _parse_date_filter,
    cancel_planning_job, request_job_cancellation, approve_job_cancellation,
    _build_job_card_pdf_bytes, build_job_card_merge_context,
    _sku_key, _missing_required_master_fields,
    _sync_new_sku_requirement, _build_recipe_map, _to_optional_positive_int,
    _to_optional_decimal, _sanitize_po_payload_items, _po_payload_items,
    _annotate_items_with_recipe, _deduplicate_po_items_by_sku,
    _history_repeat_new_counts, _sync_repeat_jobs_from_po,
    _sync_new_jobs_for_approved_sku, _merge_po_items_for_existing_po,
    get_po_approval_date_for_job,
    _collect_pending_sku_rows, _normalize_po_number,
    trigger_plate_request_for_planning_job, document_type_label,
    get_plate_request_block_for_master_entry,
    ensure_draft_planning_job_for_po_sku,
    apply_master_data_sync,
    apply_sku_recipe_form_role_permissions,
    merge_preserved_sku_recipe_fields,
    restore_locked_designer_fields_on_recipe,
    prepare_sku_recipe_form_for_master_entry,
    get_plate_making_prerequisite_errors,
    get_sku_recipe_form_ui_context,
    get_best_sku_recipe_for_sku,
    ensure_sku_recipe_for_planning_job,
    sync_planning_job_fields_to_sku_recipe,
    build_sku_recipe_initial_from_planning_job,
    build_sku_recipe_initial_from_recipe,
    hydrate_sku_recipe_from_planning_jobs,
    get_plate_remake_warning_for_recipe_save,
    get_plate_remake_warning_for_job_sync,
    planner_can_edit_designer_fields,
    PLATE_REMAKE_IMPACT_FIELDS,
    SKU_RECIPE_STATUS_ORDER,
    can_request_master_data_sync,
    dismiss_master_data_sync_request,
    get_master_data_field_diffs,
    get_job_qc_submission_blockers,
    preview_job_qc_submission_blockers,
    preview_job_qc_submission_warnings,
    job_has_master_data_mismatch,
    job_requires_reopen_for_master_sync,
    preview_master_sync_calculations,
    reopen_and_apply_master_data_sync,
    request_master_data_sync,
)
from .models import (
    PLANNING_CANCEL_REASON_CHOICES,
    PLANNING_QC_GATE_STATUSES,
    PLANNING_STATUS_ALIASES,
    PLANNING_STATUS_CHOICES,
    PLANNING_STAGE_CHOICES,
    PURCHASE_MATERIAL_ORIGIN_CHOICES,
    MERGE_GROUP_OPEN_STATUSES,
    MergeGroup,
    MergeGroupItem,
    PlanningDispatchRun,
    PlanningJob,
    PlanningPrintRun,
    PoDocument,
    SkuRecipe,
    JobCardChangeRequest,
)
from printing_plates.models import PlateRequest
from workflow.services import (
    _annotate_items_with_recipe,
    _build_cost_mismatch_note,
    _build_recipe_map,
    _collect_pending_sku_rows,
    _format_decimal_string,
    _format_display_qty,
    _missing_required_master_fields,
    _warning_master_fields,
    _normalize_application_input,
    _normalize_color_spec_input,
    _normalize_status,
    _parse_iso_date,
    _po_payload_items,
    _sanitize_po_payload_items,
    _sku_key,
    sync_job_card_for_planning_status,
    _sync_new_jobs_for_approved_sku,
    _to_decimal,
    _to_int,
    _to_optional_decimal,
    _to_optional_positive_int,
    _user_is_admin,
)
from .po_extractor import extract_po_from_pdf


def _clear_ignored_sku_from_po_docs(po_number, sku):
    if not po_number or not sku:
        return 0
    normalized_sku = _sku_key(sku)
    updated_count = 0
    for doc in PoDocument.objects.filter(extracted_payload__po_number__iexact=po_number):
        payload = doc.extracted_payload or {}
        current_ignored = [s for s in (payload.get('new_skus_ignored') or []) if s]
        remaining_ignored = [s for s in current_ignored if _sku_key(s) != normalized_sku]
        if len(remaining_ignored) != len(current_ignored):
            payload['new_skus_ignored'] = sorted(remaining_ignored)
            doc.extracted_payload = payload
            doc.save(update_fields=['extracted_payload'])
            updated_count += 1
    return updated_count


def _display_user_identity(user):
    if not user:
        return ''
    full_name = (user.get_full_name() or '').strip()
    if full_name:
        return full_name
    return (user.username or '').strip()


def _po_split_requests(payload):
    requests = payload.get('split_requests') if isinstance(payload, dict) else []
    return requests if isinstance(requests, list) else []


def _pending_po_split_requests():
    pending = []
    for doc in PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('-created_at', '-id')[:500]:
        payload = doc.extracted_payload or {}
        for request_data in _po_split_requests(payload):
            if (request_data.get('status') or 'pending') != 'pending':
                continue
            pending.append({
                'po_doc': doc,
                'po_doc_id': doc.id,
                'po_number': payload.get('po_number') or '-',
                'request_id': request_data.get('id') or '',
                'sku': request_data.get('sku') or '',
                'job_name': request_data.get('job_name') or '',
                'quantity': request_data.get('quantity') or '',
                'unit': request_data.get('unit') or '',
                'requested_qty': request_data.get('requested_qty') or '',
                'reason': request_data.get('reason') or '',
                'requested_by': request_data.get('requested_by') or '',
                'requested_at': request_data.get('requested_at') or '',
            })
    return pending


def _update_po_split_request(po_doc, request_id, status, actor=None):
    payload = po_doc.extracted_payload or {}
    changed = False
    for request_data in _po_split_requests(payload):
        if str(request_data.get('id')) != str(request_id):
            continue
        request_data['status'] = status
        request_data['resolved_at'] = timezone.now().isoformat()
        if actor:
            request_data['resolved_by'] = _display_user_identity(actor)
        changed = True
        break
    if changed:
        po_doc.extracted_payload = payload
        po_doc.save(update_fields=['extracted_payload'])
    return changed


def _get_po_approval_date_for_job(job):
    return get_po_approval_date_for_job(job)


def _get_po_upload_date_for_job(job):
    if hasattr(job, 'po_documents'):
        po_document = job.po_documents.order_by('created_at').first()
        if po_document and po_document.created_at:
            return po_document.created_at.date()

    if job.po_number:
        po_document = PoDocument.objects.filter(
            extracted_payload__po_number__iexact=job.po_number,
        ).order_by('created_at').first()
        if po_document and po_document.created_at:
            return po_document.created_at.date()

    return None


PLANNING_STATUSES = PLANNING_STATUS_CHOICES
PLANNING_STATUS_SET = {value for value, _ in PLANNING_STATUSES}
PLANNING_ACTIVE_STATUS_SET = {'released', 'in_production', 'completed', 'closed'}
PLANNING_QUEUE_STATUS_SET = {'draft', 'pending_qc', 'qc_approved'}
PLANNING_STATUS_FILTER_ALIASES = {
    'draft': {'draft', 'open', 'pending'},
    'pending_qc': {'pending_qc', 'reviewed'},
    'qc_approved': {'qc_approved', 'approved'},
    'released': {'released'},
    'in_production': {'in_production'},
    'completed': {'completed', 'closed'},
}
NEW_SKU_REQUIREMENT_NOTE = 'NEW SKU: Shade matching and setup verification required before production run.'
COST_MISMATCH_NOTE_PREFIX = 'COST ALERT:'
SKU_MASTER_APPROVAL_REQUIRED_FIELDS = [
    ('job_name', 'Job Name'),
    ('material', 'Material'),
    ('job_process_type', 'Job Process'),
    ('color_spec', 'Print Color'),
    ('application', 'Application'),
    ('product_type', 'Product Type'),
    ('size_w_mm', 'Size Width (mm)'),
    ('size_h_mm', 'Size Height (mm)'),
    ('print_sheet_size', 'Print Sheet'),
    ('purchase_sheet_size', 'Purchase Sheet'),
    ('ups', 'UPS'),
    ('purchase_sheet_ups', 'Purchase Sheet Ups'),
    ('awc_no', 'AWC #'),
    ('die_cutting', 'Die Cutting'),
    ('plate_set_no', 'Plate Set No.'),
    ('print_passes', 'No. of Passes'),
]

_COLOR_PLUS_RE = re.compile(r'^(\d+)\s*\+\s*(\d+)$')
_COLOR_SINGLE_RE = re.compile(r'^(\d+)\s*(?:colou?r(?:s)?)?$', re.IGNORECASE)


def _normalize_purchase_material_origin(raw_value):
    value = (raw_value or '').strip().lower()
    if value in {'local'}:
        return 'local'
    if value in {'import', 'imported'}:
        return 'import'
    return ''


def _effective_planning_status(job):
    return job.effective_status


def _effective_planning_status_label(job, effective_status):
    return job.effective_status_label


def _repair_rejected_job_status(job):
    if _normalize_status(job.status) != 'pending_qc':
        return False
    job_card = getattr(job, 'job_card', None)
    if not job_card:
        return False
    try:
        card_status = (job_card.workflow_status or '').strip().lower()
    except Exception:
        return False
    if card_status in {'draft', 'pending_data', 'qc_rejected', 'pm_rejected'}:
        job.status = 'draft'
        job.save(update_fields=['status', 'updated_at'])
        return True
    return False


def _effective_planning_status_from_values(planning_status, job_card_status):
    status_rank = {
        'draft': 0,
        'pending_qc': 1,
        'qc_approved': 2,
        'released': 3,
        'in_production': 4,
        'completed': 5,
    }
    planning_status = _normalize_status(planning_status)
    card_status = (job_card_status or '').strip().lower()
    if planning_status == 'draft':
        return 'draft'
    if card_status in {'qc_rejected', 'pm_rejected'}:
        return 'draft'
    job_card_status_mapped = None
    if card_status in {'planning_approved', 'pending_qc'}:
        job_card_status_mapped = 'pending_qc'
    elif card_status == 'qc_approved':
        job_card_status_mapped = 'qc_approved'
    elif card_status in {'pending_pm_approval', 'production_approved', 'released', 'in_production', 'completed', 'closed'}:
        job_card_status_mapped = 'released'
    elif card_status in status_rank:
        job_card_status_mapped = card_status

    if planning_status not in status_rank:
        return job_card_status_mapped or planning_status or 'draft'
    if not job_card_status_mapped:
        return planning_status

    return planning_status if status_rank[planning_status] >= status_rank[job_card_status_mapped] else job_card_status_mapped


def build_planning_readme_text():
    return """Offset ERP - Planning Module Easy Guide

Last Updated: 2026-07-10

=============================
1) MASTER SKU (STEP 1)
=============================
Purpose:
- Keep approved SKU master data ready before routing PO jobs.

How to use:
- Create or bulk upload SKU master data.
- Save as Draft first.
- Move Draft -> Reviewed -> Approved.
- Only approved recipes are used as final master data.

Required fields for approval:
- Job Name, Material, Color, Application, Machine
- Print Sheet, Purchase Sheet, UPS, Purchase Material Origin

=============================
2) PO INTAKE (STEP 2)
=============================
Purpose:
- Upload PO and split lines into Repeat and New.

Routing rule:
- Repeat lines -> Planning Jobs
- New lines -> Pending SKU Master Data

Important notes:
- Duplicate SKU lines in one PO are merged.
- Qty display is normalized (trailing decimals removed).

=============================
3) PENDING NEW SKU (STEP 3)
=============================
Purpose:
- Complete missing master data for new SKU lines from PO.

Rules:
- Job Name comes from PO and is not manual.
- Department and unit cost can prefill from PO when available.
- Application must be one of: UV, Lamination Gloss, Lamination Matt, NO.
- Purchase Material Origin must be Local or Imported.

Approval path:
- Save Draft -> Send For Approval -> Approved
- Approved new SKU records refresh matching draft planning jobs.

=============================
4) PLANNING JOBS (STEP 4)
=============================
Purpose:
- Create/update, review, and manage production planning jobs.

Input source:
- Repeat jobs from PO intake
- Approved new SKUs after master approval

Operational controls:
- Filter by PO/SKU/status/department/machine/date.
- Bulk status update for selected rows.
- Open detail, edit, print A4 job card.

=============================
5) APPROVAL QUEUE (STEP 5)
=============================
Purpose:
- Release jobs through QC and Production Manager checkpoints.

Status transitions:
- Draft -> Pending Review
- Pending Review -> Pending Approval (Manager)

After approval:
- Print job card and run shop-floor execution flow.

=============================
6) SHOP FLOOR EXECUTION (STEP 6)
=============================
Purpose:
- Use QR scan and A4 card for execution traceability.

Use:
- Open job via scan.
- Track run/dispatch logs from planning-linked records.

=============================
7) CHANGE & MISMATCH ALERTS
=============================
Repeat route behavior:
- If PO cost differs from master cost, a COST ALERT note is attached.
- If PO department differs from master department, a DEPARTMENT ALERT note is attached.

New route behavior:
- New SKU records must complete master approval before planning sync.

=============================
8) DAILY DISCIPLINE
=============================
- Keep master records clean and approved.
- Route every PO through intake queue.
- Resolve pending SKUs same day.
- Complete approvals before production release.

=============================
9) DUPLICATE SKU ALERTS
=============================
Purpose:
- Warn planners when the same SKU is active on more than one PO / JC at the same time.

Where shown:
- Planning Jobs list (badge on JC)
- Job detail, Pending SKU master entry, Approval Queue (full banner)
- Every active JC in the cluster shows the same alert (not only the newest PO).

Combine run (save machine time):
- Shown when two or more active jobs for the SKU have NOT started printing yet.
- Primary reference is the FIRST PO / earliest JC.
- Expedite planning, release job cards, and tell production to run together.
- Keep documentation separate per PO (qty, dispatch, invoicing).

Priority hint (lower urgency):
- Shown when any JC for this SKU was dispatched within the last 14 days.
- Use for scheduling — new PO lines can wait unless customer expedite applies.
- Combine is not applicable once printing has started on a sibling job.

Plates:
- When combining or repeating the same artwork, reuse plates from the first PO.
- Do not send a second plate request unless plates are damaged or artwork changed.
"""


def _user_can_cancel_planning_job(user):
    profile = getattr(user, 'profile', None)
    if not profile:
        return False
    return profile.can_cancel_planning_job()


def _clear_cancellation(job):
    """Un-cancel a job being restored. Returns the extra update_fields touched."""
    if not job.is_cancelled:
        return []
    job.is_cancelled = False
    job.cancel_reason = ''
    job.cancel_reason_code = ''
    job.cancelled_by = None
    job.cancelled_at = None
    job.status = 'draft'
    return ['is_cancelled', 'cancel_reason', 'cancel_reason_code', 'cancelled_by', 'cancelled_at', 'status']


def _report_validation_error(request, exc, prefix=''):
    """Surface a ValidationError (dict or plain) as user-facing messages."""
    label = f'{prefix}: ' if prefix else ''
    error_dict = getattr(exc, 'message_dict', None)
    if error_dict:
        for field_errors in error_dict.values():
            for error_message in field_errors:
                messages.error(request, f'{label}{error_message}')
    else:
        messages.error(request, f'{label}{exc}')


@login_required
@permission_required('can_view_jobcard')
def planning_readme(request):
    return render(
        request,
        'planning/planning_readme.html',
        {'generated_on': timezone.now()},
    )


@login_required
@permission_required('can_view_jobcard')
def download_planning_readme(request):
    content = build_planning_readme_text()
    response = HttpResponse(content, content_type='text/plain; charset=utf-8')
    response['Content-Disposition'] = 'attachment; filename="planning_workflow_guide.txt"'
    return response


@login_required
def planning_welcome(request):
    profile = getattr(request.user, 'profile', None)
    user_role = 'unassigned'
    can_edit_jobcard = False
    can_manage_masters = False

    if profile is not None:
        user_role = (profile.role or 'unassigned').strip().lower()
        can_edit_jobcard = bool(profile.can_edit_jobcard())
        can_manage_masters = bool(profile.can_manage_masters())

    context = {
        'user_role': user_role,
        'can_edit_jobcard': can_edit_jobcard,
        'can_manage_masters': can_manage_masters,
    }
    return render(request, 'planning/planning_welcome.html', context)


@login_required
@permission_required('can_view_jobcard')
def planning_po_root(request):
    return redirect('planning:po_inbox')


@login_required
@permission_required('can_view_jobcard')
def planning_pending_actions(request):
    return redirect(f"{reverse('planning:jobs')}?status=draft")


@login_required
@permission_required('can_view_jobcard')
def planning_jobs_drafts(request):
    return redirect(f"{reverse('planning:jobs')}?status=draft")


@login_required
@permission_required('can_view_jobcard')
def planning_jobs_locked(request):
    return redirect(f"{reverse('planning:jobs')}?status=qc_approved")


SUMMARY_COLUMN_OPTIONS = [
    ('jc_number', 'Job Card No'),
    ('po_number', 'PO Number'),
    ('sku', 'SKU'),
    ('job_name', 'Job Name'),
    ('status', 'Status'),
    ('planning_stage', 'Planning Stage'),
    ('machine_name', 'Machine'),
    ('department', 'Department'),
    ('order_qty', 'Order Qty'),
    ('delivery_date', 'Delivery Date'),
    ('po_approval_date', 'PO Approval Date'),
    ('plan_date', 'Plan Date'),
]

DEFAULT_SUMMARY_COLUMNS = [
    'jc_number', 'po_number', 'sku', 'job_name', 'status',
    'planning_stage', 'machine_name', 'department', 'order_qty',
    'delivery_date',
]


def _summary_column_value(job, column):
    if column == 'po_approval_date':
        return job.po_approval_date.strftime('%Y-%m-%d') if job.po_approval_date else ''
    if column == 'plan_date':
        return job.plan_date.strftime('%Y-%m-%d') if job.plan_date else ''
    if column == 'delivery_date':
        return job.delivery_date.strftime('%Y-%m-%d') if job.delivery_date else ''
    if column == 'order_qty':
        return job.order_qty or ''
    if column == 'machine_name':
        return job.machine_name or ''
    if column == 'department':
        return job.department or ''
    if column == 'planning_stage':
        return job.planning_stage or ''
    if column == 'status':
        return job.status or ''
    return getattr(job, column, '') or ''


@login_required
@permission_required('can_view_planning_queue')
def planning_jobs_summary(request):
    column_keys = [key for key, _ in SUMMARY_COLUMN_OPTIONS]
    selected_columns = request.GET.getlist('columns')
    if not selected_columns:
        raw_columns = (request.GET.get('columns') or '').strip()
        if raw_columns:
            selected_columns = [item.strip() for item in raw_columns.split(',') if item.strip()]
    selected_columns = [col for col in selected_columns if col in column_keys]
    if not selected_columns:
        selected_columns = DEFAULT_SUMMARY_COLUMNS.copy()

    q = (request.GET.get('q') or '').strip()
    status_filter = (request.GET.get('status') or '').strip()
    department_filter = (request.GET.get('department') or '').strip()
    machine_filter = (request.GET.get('machine') or '').strip()
    from_date = _parse_date_filter(request.GET.get('from_date'))
    to_date = _parse_date_filter(request.GET.get('to_date'))

    queryset = PlanningJob.objects.filter(is_active=True)
    if q:
        queryset = queryset.filter(
            Q(jc_number__icontains=q)
            | Q(po_number__icontains=q)
            | Q(sku__icontains=q)
            | Q(job_name__icontains=q)
        )
    if status_filter:
        queryset = queryset.filter(status__in=_planning_status_filter_values(status_filter))
    if department_filter:
        queryset = queryset.filter(department__icontains=department_filter)
    if machine_filter:
        queryset = queryset.filter(machine_name__icontains=machine_filter)
    if from_date:
        queryset = queryset.filter(po_approval_date__gte=from_date)
    if to_date:
        queryset = queryset.filter(po_approval_date__lte=to_date)

    jobs = queryset.order_by('-plan_date', '-id').select_related('job_card')[:1000]
    rows = []
    for job in jobs:
        row = {'id': job.id}
        row.update({col: _summary_column_value(job, col) for col in selected_columns})
        rows.append(row)

    export_type = (request.GET.get('export') or '').strip().lower()
    if export_type in {'xlsx', 'pdf'}:
        payload = {
            'report': {'title': 'Job Summary'},
            'generated_at': timezone.localtime().strftime('%Y-%m-%d %H:%M:%S'),
            'data': rows,
        }
        if export_type == 'xlsx':
            content = export_as_xlsx(payload)
            response = HttpResponse(
                content,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            response['Content-Disposition'] = 'attachment; filename="job-summary.xlsx"'
            return response
        content = export_as_pdf(payload)
        response = HttpResponse(content, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="job-summary.pdf"'
        return response

    return render(
        request,
        'planning/planning_jobs_summary.html',
        {
            'rows': rows,
            'selected_columns': selected_columns,
            'column_options': SUMMARY_COLUMN_OPTIONS,
            'status_choices': PLANNING_STATUSES,
            'filters': {
                'q': q,
                'status': status_filter,
                'department': department_filter,
                'machine': machine_filter,
                'from_date': request.GET.get('from_date', ''),
                'to_date': request.GET.get('to_date', ''),
            },
            'export_query': request.GET.urlencode(),
        },
    )


@login_required
@permission_required('can_view_jobcard')
def planning_sku_queue(request):
    return redirect('planning:pending_skus')


@login_required
@permission_required('can_view_jobcard')
def planning_sku_recipes_list(request):
    return redirect('planning:sku_recipes')


@login_required
def planning_home(request):
    from printing_plates.models import PlateRequest

    _user_can_plan = getattr(getattr(request.user, 'profile', None), 'can_plan', lambda: False)()
    user_can_release = _user_can_plan
    queryset = PlanningJob.objects.select_related('job_card', 'planning_stage_changed_by').prefetch_related(
        Prefetch(
            'plate_requests',
            queryset=PlateRequest.objects.order_by('-requested_at', '-created_at', '-id'),
        )
    ).filter(
        is_active=True,
    )

    if request.method == 'POST':
        action = request.POST.get('action')
        if not _user_can_plan:
            messages.error(request, 'You do not have permission to modify planning jobs.')
            return redirect('planning:jobs')

        if action == 'bulk_update_status':
            selected_ids = []
            for raw_id in request.POST.getlist('selected_ids'):
                try:
                    selected_ids.append(int(raw_id))
                except (TypeError, ValueError):
                    continue

            target_status = _normalize_status(request.POST.get('target_status'), default='')
            if target_status not in {'released', 'in_production', 'completed'}:
                messages.error(request, 'Please select a valid target status for bulk update.')
                return redirect('planning:jobs')

            if not selected_ids:
                messages.error(request, 'Select at least one planning row for bulk update.')
                return redirect('planning:jobs')

            updated = 0
            skipped_locked = 0
            status_rank = {'released': 1, 'in_production': 2, 'completed': 3}
            for job in PlanningJob.objects.filter(id__in=selected_ids, is_active=True):
                current_status = _normalize_status(job.status)
                if status_rank.get(current_status, 0) > status_rank.get(target_status, 0):
                    skipped_locked += 1
                    continue
                if current_status == target_status:
                    continue

                job.status = target_status
                job.issued_to_production = True
                try:
                    job.save(update_fields=['status', 'issued_to_production', 'updated_at'])
                except ValidationError as exc:
                    skipped_locked += 1
                    error_dict = getattr(exc, 'message_dict', None)
                    if error_dict:
                        for field_errors in error_dict.values():
                            for error_message in field_errors:
                                messages.error(request, f'{job.jc_number}: {error_message}')
                    else:
                        messages.error(request, f'{job.jc_number}: {exc}')
                    continue
                updated += 1

            messages.success(
                request,
                f'Bulk status update complete. Updated {updated}, locked-skip {skipped_locked}.',
            )
            return redirect('planning:jobs')

        if action == 'bulk_archive':
            selected_ids = []
            for raw_id in request.POST.getlist('selected_ids'):
                try:
                    selected_ids.append(int(raw_id))
                except (TypeError, ValueError):
                    continue

            reason = (request.POST.get('archive_reason') or '').strip()
            if not selected_ids:
                messages.error(request, 'Select at least one planning row to archive.')
                return redirect('planning:jobs')

            archived_count = 0
            for job in PlanningJob.objects.filter(id__in=selected_ids, is_active=True):
                job.is_active = False
                job.archive_reason = reason
                job.archived_by = request.user
                job.archived_at = timezone.now()
                job.save(update_fields=['is_active', 'archive_reason', 'archived_by', 'archived_at', 'updated_at'])
                archived_count += 1

            messages.success(request, f'Bulk archive complete. Archived {archived_count} jobs.')
            return redirect('planning:jobs')

        if action == 'bulk_delete':
            if not _user_is_admin(request.user):
                messages.error(request, 'Only administrators can permanently delete planning jobs.')
                return redirect('planning:jobs')

            selected_ids = []
            for raw_id in request.POST.getlist('selected_ids'):
                try:
                    selected_ids.append(int(raw_id))
                except (TypeError, ValueError):
                    continue

            if not selected_ids:
                messages.error(request, 'Select at least one planning row to delete.')
                return redirect('planning:jobs')

            deleted_count = PlanningJob.objects.filter(id__in=selected_ids, is_active=True).delete()[0]
            messages.success(request, f'Bulk delete complete. Deleted {deleted_count} jobs.')
            return redirect('planning:jobs')

        if action in {'hold', 'release_hold', 'archive', 'cancel', 'delete', 'update_planning_stage'}:
            job_id = request.POST.get('job_id')
            try:
                job_id = int(job_id)
            except (TypeError, ValueError):
                messages.error(request, 'Invalid planning job selected.')
                return redirect('planning:jobs')

            job = get_object_or_404(PlanningJob, id=job_id)
            if action == 'delete' and not _user_is_admin(request.user):
                messages.error(request, 'Only administrators can delete planning jobs.')
                return redirect('planning:jobs')

            if action == 'delete':
                try:
                    job.delete()
                    messages.success(request, f'Planning job {job.jc_number} was permanently deleted.')
                except ProtectedError:
                    if _user_is_admin(request.user):
                        protected_plate_requests = PlateRequest.objects.filter(planning_job=job)
                        if protected_plate_requests.exists():
                            protected_count = protected_plate_requests.count()
                            protected_plate_requests.delete()
                            job.delete()
                            messages.success(
                                request,
                                f'Planning job {job.jc_number} and {protected_count} linked plate request(s) were permanently deleted.',
                            )
                        else:
                            messages.error(request, f'Cannot delete planning job {job.jc_number} because it is referenced by other records.')
                    else:
                        messages.error(request, 'Only administrators can delete planning jobs.')
                return redirect('planning:jobs')

            if action == 'update_planning_stage':
                planning_stage = (request.POST.get('planning_stage') or '').strip()

                # Enforce stage lock when plate making is in progress
                if job.planning_stage in ['new_plate_making', 'repeat_plate_making']:
                    from printing_plates.models import PlateRequest
                    active_request = PlateRequest.objects.filter(
                        planning_job=job,
                        status__in=[PlateRequest.STATUS_DRAFT, PlateRequest.STATUS_SENT, PlateRequest.STATUS_RECEIVED]
                    ).exists()
                    if active_request:
                        messages.error(request, f'Cannot change stage. Plate making is currently in progress for {job.jc_number}.')
                        return redirect('planning:jobs')

                # Resolve plate-making stage from repeat_flag (never trust raw new/repeat stage POST).
                if planning_stage in ['plate_making', 'new_plate_making', 'repeat_plate_making']:
                    plate_errors = get_plate_making_prerequisite_errors(job)
                    if plate_errors:
                        messages.error(request, ' '.join(plate_errors))
                        return redirect('planning:jobs')

                    from printing_plates.services import get_planning_plate_making_block_message

                    block_message = get_planning_plate_making_block_message(job)
                    if block_message:
                        messages.error(request, block_message)
                        return redirect('planning:jobs')

                    from planning.sku_classification import plate_making_stage_for_repeat_flag

                    planning_stage = plate_making_stage_for_repeat_flag(job.repeat_flag)

                valid_stages = [choice[0] for choice in PLANNING_STAGE_CHOICES]
                if planning_stage not in valid_stages:
                    messages.error(request, 'Please select a valid planning stage.')
                    return redirect('planning:jobs')

                job.planning_stage = planning_stage
                job.planning_stage_changed_at = timezone.now()
                job.planning_stage_changed_by = request.user
                job.save(update_fields=['planning_stage', 'planning_stage_changed_at', 'planning_stage_changed_by', 'updated_at'])
                trigger_plate_request_for_planning_job(job, request.user)
                stage_display = dict(PLANNING_STAGE_CHOICES).get(planning_stage, 'Not Set')
                messages.success(request, f'Planning job {job.jc_number} stage updated to: {stage_display}.')
                return redirect('planning:jobs')

            if action == 'hold':
                reason = (request.POST.get('reason') or '').strip()
                if not reason:
                    messages.error(request, 'A hold reason is required to place a job on hold.')
                    return redirect('planning:jobs')
                job.is_on_hold = True
                job.hold_reason = reason
                job.hold_by = request.user
                job.hold_at = timezone.now()
                job.save(update_fields=['is_on_hold', 'hold_reason', 'hold_by', 'hold_at', 'updated_at'])
                messages.success(request, f'Planning job {job.jc_number} was placed on hold.')
                return redirect('planning:jobs')

            if action == 'release_hold':
                job.is_on_hold = False
                job.hold_reason = ''
                job.hold_by = None
                job.hold_at = None
                job.save(update_fields=['is_on_hold', 'hold_reason', 'hold_by', 'hold_at', 'updated_at'])
                messages.success(request, f'Planning job {job.jc_number} hold was released.')
                return redirect('planning:jobs')

            if action == 'release_for_production':
                if not user_can_release:
                    messages.error(request, 'You do not have permission to release jobs for production.')
                    return redirect('planning:jobs')
                job_card = getattr(job, 'job_card', None)
                if not job_card:
                    messages.error(request, f'Job Card missing for planning job {job.jc_number}.')
                    return redirect('planning:jobs')
                plate_set_no = (request.POST.get('plate_set_no') or '').strip()
                if plate_set_no:
                    job.plate_set_no = plate_set_no
                    job.save(update_fields=['plate_set_no', 'updated_at'])
                try:
                    execute_job_card_action(job_card, 'release_for_production', actor=request.user, reason='Released from planning jobs list')
                    messages.success(request, f'Planning job {job.jc_number} was released for production.')
                except ValidationError as exc:
                    error_dict = getattr(exc, 'message_dict', None)
                    if error_dict:
                        for field_errors in error_dict.values():
                            for error_message in field_errors:
                                messages.error(request, f'{job.jc_number}: {error_message}')
                    else:
                        messages.error(request, f'{job.jc_number}: {exc}')
                return redirect('planning:jobs')

            if action == 'cancel':
                if not _user_can_cancel_planning_job(request.user):
                    messages.error(request, 'You do not have permission to cancel planning jobs.')
                    return redirect('planning:jobs')

                reason = (request.POST.get('reason') or '').strip()
                reason_code = (request.POST.get('cancel_reason_code') or '').strip()
                needs_approval = not job.can_cancel_directly
                try:
                    if needs_approval:
                        request_job_cancellation(job, actor=request.user, reason=reason, reason_code=reason_code)
                        messages.success(
                            request,
                            f'Cancellation request for {job.jc_number} was sent for approval.',
                        )
                    else:
                        cancel_planning_job(job, actor=request.user, reason=reason, reason_code=reason_code)
                        messages.success(request, f'Planning job {job.jc_number} was cancelled.')
                except ValidationError as exc:
                    _report_validation_error(request, exc, prefix=job.jc_number)
                return redirect('planning:jobs')

            if action == 'archive':
                reason = (request.POST.get('reason') or '').strip()
                job.is_active = False
                job.archive_reason = reason
                job.archived_by = request.user
                job.archived_at = timezone.now()
                job.save(update_fields=['is_active', 'archive_reason', 'archived_by', 'archived_at', 'updated_at'])
                messages.success(request, f'Planning job {job.jc_number} was archived.')
                return redirect('planning:jobs')

            if action == 'delete':
                job.delete()
                messages.success(request, f'Planning job {job.jc_number} was permanently deleted.')
                return redirect('planning:jobs')

    q = (request.GET.get('q') or '').strip()
    status_values = [v for v in request.GET.getlist('status') if v and v.strip()]
    stage_values = [v for v in request.GET.getlist('planning_stage') if v and v.strip()]
    department_filter = (request.GET.get('department') or '').strip()
    machine_filter = (request.GET.get('machine') or '').strip()
    from_date = _parse_date_filter(request.GET.get('from_date'))
    to_date = _parse_date_filter(request.GET.get('to_date'))
    sku_alert_filter = (request.GET.get('sku_alert') or '').strip().lower()

    if q:
        queryset = queryset.filter(
            Q(jc_number__icontains=q)
            | Q(po_number__icontains=q)
            | Q(sku__icontains=q)
            | Q(job_name__icontains=q)
        )
    if status_values:
        expanded_statuses = set()
        for value in status_values:
            expanded_statuses.update(_planning_status_filter_values(value))
        queryset = queryset.filter(status__in=sorted(expanded_statuses))
    elif not q:
        queryset = queryset.exclude(status='completed')
    # else: searching without an explicit status filter — leave completed jobs in
    # results so a specific job can never "vanish" from search (read-only in the template).
    if stage_values:
        stage_q = Q()
        for stage_filter in stage_values:
            if stage_filter == 'planning_done':
                stage_q |= Q(planning_stage='planning_done') | Q(planning_stage='in_production')
            elif stage_filter == 'not_set':
                stage_q |= Q(planning_stage='')
            else:
                stage_q |= Q(planning_stage=stage_filter)
        queryset = queryset.filter(stage_q)
    if department_filter:
        queryset = queryset.filter(department__icontains=department_filter)
    if machine_filter:
        queryset = queryset.filter(machine_name__icontains=machine_filter)

    # Merged job cards are managed under the Smart Merge tab, so they are hidden
    # from the regular queue by default to keep it clean. But they must never
    # "vanish": a search always surfaces them, and a "Merged" quick-filter chip
    # shows them on demand.
    show_merged = request.GET.get('merged') == '1'
    merged_member_ids = list(
        MergeGroupItem.objects.filter(merge_group__status__in=MERGE_GROUP_OPEN_STATUSES)
        .values_list('planning_job_id', flat=True)
    )
    merged_jobs_count = queryset.filter(id__in=merged_member_ids).count() if merged_member_ids else 0
    if merged_member_ids:
        if show_merged:
            queryset = queryset.filter(id__in=merged_member_ids)
        elif not q:
            queryset = queryset.exclude(id__in=merged_member_ids)
        # else: searching — leave merged jobs in the results so they are findable.

    merged_params = request.GET.copy()
    merged_params.pop('page', None)
    if show_merged:
        merged_params.pop('merged', None)
    else:
        merged_params['merged'] = '1'
    merged_filter_url = '?' + merged_params.urlencode() if merged_params.urlencode() else '?'

    from planning.sku_duplicate_alert import (
        SKU_ALERT_FILTER_KEYS,
        attach_sku_duplicate_alerts_to_jobs,
        build_planning_jobs_sku_alert_filter_urls,
        count_planning_jobs_by_sku_alert,
        filter_planning_jobs_by_sku_alert,
        job_matches_sku_alert_filter,
    )

    sku_alert_counts = count_planning_jobs_by_sku_alert(queryset)
    if sku_alert_filter in SKU_ALERT_FILTER_KEYS:
        queryset = filter_planning_jobs_by_sku_alert(queryset, sku_alert_filter)

    jobs_list = None
    if from_date or to_date:
        queryset = queryset.select_related('job_card')
        jobs_list = list(queryset)
        for job in jobs_list:
            job.po_approval_date_display = _get_po_approval_date_for_job(job)
            po_uploaded_date = _get_po_upload_date_for_job(job)
            if po_uploaded_date:
                job.plan_date_display = po_uploaded_date
            else:
                job.plan_date_display = job.plan_date or (job.created_at.date() if job.created_at else None)

        if from_date:
            jobs_list = [job for job in jobs_list if job.po_approval_date_display and job.po_approval_date_display >= from_date]
        if to_date:
            jobs_list = [job for job in jobs_list if job.po_approval_date_display and job.po_approval_date_display <= to_date]
        if sku_alert_filter in SKU_ALERT_FILTER_KEYS:
            jobs_list = [
                job for job in jobs_list
                if job_matches_sku_alert_filter(job, sku_alert_filter)
            ]

    export_type = (request.GET.get('export') or '').strip().lower()
    if export_type in {'xlsx', 'pdf'}:
        export_columns = [
            'jc_number', 'po_number', 'sku', 'job_name', 'sku_alert', 'active_jobs',
            'combine_possible', 'related_jcs', 'status', 'planning_stage',
            'machine_name', 'department', 'order_qty', 'delivery_date',
            'po_approval_date', 'plan_date',
        ]
        if jobs_list is not None:
            export_jobs = list(jobs_list[:1000])
        else:
            export_jobs = list(
                queryset.select_related('job_card').order_by('sku', '-plan_date', '-id')[:1000]
            )
        attach_sku_duplicate_alerts_to_jobs(export_jobs)
        # Group duplicate SKUs together so combine candidates are easy to scan.
        export_jobs.sort(
            key=lambda job: (
                (job.sku or '').strip().lower(),
                job.po_approval_date or job.plan_date or date.min,
                job.id or 0,
            )
        )
        rows = []
        for job in export_jobs:
            alert = getattr(job, 'sku_duplicate_alert', None) or {}
            alert_kind = alert.get('alert_kind') or ''
            if alert_kind == 'combine':
                sku_alert_label = 'Duplicate SKU · Combine'
            elif alert_kind == 'duplicate':
                sku_alert_label = 'Duplicate SKU'
            elif alert_kind == 'low_priority':
                sku_alert_label = 'Low priority'
            else:
                sku_alert_label = ''
            related_jcs = ''
            if alert.get('members'):
                related_jcs = ', '.join(
                    member.get('jc_number') or '-'
                    for member in alert['members']
                    if member.get('job_id') != job.id
                )
            row = {
                'id': job.id,
                'jc_number': _summary_column_value(job, 'jc_number'),
                'po_number': _summary_column_value(job, 'po_number'),
                'sku': _summary_column_value(job, 'sku'),
                'job_name': _summary_column_value(job, 'job_name'),
                'sku_alert': sku_alert_label,
                'active_jobs': alert.get('active_count') or '',
                'combine_possible': 'Yes' if alert.get('combine_possible') else ('No' if alert else ''),
                'related_jcs': related_jcs,
                'status': _summary_column_value(job, 'status'),
                'planning_stage': _summary_column_value(job, 'planning_stage'),
                'machine_name': _summary_column_value(job, 'machine_name'),
                'department': _summary_column_value(job, 'department'),
                'order_qty': _summary_column_value(job, 'order_qty'),
                'delivery_date': _summary_column_value(job, 'delivery_date'),
                'po_approval_date': _summary_column_value(job, 'po_approval_date'),
                'plan_date': _summary_column_value(job, 'plan_date'),
            }
            rows.append(row)
        header_labels = {
            **dict(SUMMARY_COLUMN_OPTIONS),
            'sku_alert': 'SKU Alert',
            'active_jobs': 'Active Jobs',
            'combine_possible': 'Can Combine',
            'related_jcs': 'Related JCs',
        }
        payload = {
            'report': {'title': 'Planning Jobs'},
            'generated_at': timezone.localtime().strftime('%Y-%m-%d %H:%M:%S'),
            'data': rows,
            'headers': export_columns,
            'header_labels': header_labels,
        }
        if export_type == 'xlsx':
            content = export_as_xlsx(payload)
            response = HttpResponse(
                content,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            response['Content-Disposition'] = 'attachment; filename="planning-jobs.xlsx"'
            return response
        content = export_as_pdf(payload)
        response = HttpResponse(content, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="planning-jobs.pdf"'
        return response

    if jobs_list is not None:
        status_counts = {}
        for job in jobs_list:
            normalized_status = _normalize_status(job.status)
            status_counts[normalized_status] = status_counts.get(normalized_status, 0) + 1

        paginator = Paginator(jobs_list, 50)
        page_number = request.GET.get('page')
        jobs = paginator.get_page(page_number)
    else:
        status_rows = queryset.values('status').annotate(total=Count('id')).order_by('status')
        status_counts = {}
        for row in status_rows:
            normalized_status = _normalize_status(row['status'])
            status_counts[normalized_status] = status_counts.get(normalized_status, 0) + (row['total'] or 0)

        queryset = queryset.select_related('job_card')
        paginator = Paginator(queryset, 50)
        page_number = request.GET.get('page')
        jobs = paginator.get_page(page_number)

    job_skus = {
        _sku_key(job.sku)
        for job in jobs
        if job.sku
    }
    approved_sku_keys = {
        _sku_key(sku)
        for sku in SkuRecipe.objects.filter(
            is_active=True,
            master_data_status='approved',
            sku__in=[sku for sku in job_skus if sku],
        ).values_list('sku', flat=True)
        if sku
    }

    for job in jobs:
        job.po_approval_date_display = _get_po_approval_date_for_job(job)
        po_uploaded_date = _get_po_upload_date_for_job(job)
        if po_uploaded_date:
            job.plan_date_display = po_uploaded_date
        else:
            job.plan_date_display = job.plan_date or (job.created_at.date() if job.created_at else None)

        job.can_submit_qc = True
        job.submit_qc_block_reason = ''
        if job.effective_status == 'draft':
            has_approved_recipe = _sku_key(job.sku) in approved_sku_keys
            if not has_approved_recipe:
                job.can_submit_qc = False
                job.submit_qc_block_reason = 'SKU master is pending review/approval in QC.'

    # Build a last-status-change audit map for linked job cards.
    job_card_ids = [job.job_card.id for job in jobs if getattr(job, 'job_card', None)]
    status_logs = {}
    if job_card_ids:
        logs = ChangeLog.objects.filter(
            entity_type='job_card',
            record_id__in=job_card_ids,
        ).select_related('changed_by').order_by('-created_at')
        for log in logs:
            if log.record_id in status_logs:
                continue
            field_changes = log.field_changes if isinstance(log.field_changes, dict) else {}
            status_change = field_changes.get('status') if isinstance(field_changes, dict) else None
            if not isinstance(status_change, dict):
                continue
            status_logs[log.record_id] = log

    def _display_user_identity(user):
        if not user:
            return ''
        full_name = (user.get_full_name() or '').strip()
        if full_name:
            return full_name
        return (user.username or '').strip()

    from printing_plates.services import get_open_planning_plate_requests_blocking_release_bulk

    release_blocking_map = get_open_planning_plate_requests_blocking_release_bulk(list(jobs))

    for job in jobs:
        job_card = getattr(job, 'job_card', None)
        job.job_card_display_machine_name = job.machine_name or (job_card.machine_name_display if job_card else None)
        job.job_card_workflow_status = job_card.workflow_status if job_card else ''
        job.has_job_card = bool(job_card)
        job.planning_stage_changed_by_display = _display_user_identity(getattr(job, 'planning_stage_changed_by', None))
        job.latest_cancelled_plate = job.latest_cancelled_plate_request
        job.release_blocked_open_plate = release_blocking_map.get(job.id)
        if job_card:
            job.job_card_status_label = job_card.workflow_status_label
            log = status_logs.get(job_card.id)
            if log:
                job.status_changed_by = _display_user_identity(log.changed_by)
                job.status_changed_at = log.created_at
            else:
                job.status_changed_by = ''
                job.status_changed_at = None
        else:
            job.job_card_status_label = ''
            job.status_changed_by = ''
            job.status_changed_at = None

    attach_sku_duplicate_alerts_to_jobs(jobs)

    sku_alert_filter_urls = build_planning_jobs_sku_alert_filter_urls(
        request,
        active_key=sku_alert_filter,
    )

    filter_params = request.GET.copy()
    filter_params.pop('page', None)
    filter_params.pop('export', None)
    filter_query = filter_params.urlencode()

    return render(
        request,
        'planning/planning_home.html',
        {
            'jobs': jobs,
            'status_counts': status_counts,
            'merged_jobs_count': merged_jobs_count,
            'status_choices': [
                (value, label)
                for value, label in PLANNING_STATUSES
            ],
            'stage_choices': PLANNING_STAGE_CHOICES,
            'can_admin_actions': _user_is_admin(request.user),
            'user_can_plan': _user_can_plan,
            'user_can_cancel': _user_can_cancel_planning_job(request.user),
            'cancel_reason_choices': PLANNING_CANCEL_REASON_CHOICES,
            'filters': {
                'q': q,
                'status': status_values,
                'planning_stage': stage_values,
                'department': department_filter,
                'machine': machine_filter,
                'from_date': request.GET.get('from_date', ''),
                'to_date': request.GET.get('to_date', ''),
                'sku_alert': sku_alert_filter,
                'merged': show_merged,
            },
            'merged_filter_url': merged_filter_url,
            'sku_alert_counts': sku_alert_counts,
            'sku_alert_filter_urls': sku_alert_filter_urls,
            'filter_query': filter_query,
            'export_query': filter_query,
        },
    )


@login_required
@permission_required('can_edit_jobcard')
def planning_jobs_archived(request):
    if request.method == 'POST':
        action = request.POST.get('action')
        if action in {'bulk_restore', 'bulk_delete'}:
            selected_ids = []
            for raw_id in request.POST.getlist('selected_ids'):
                try:
                    selected_ids.append(int(raw_id))
                except (TypeError, ValueError):
                    continue

            if not selected_ids:
                messages.error(request, 'Select at least one archived planning job.')
                return redirect('planning:jobs_archived')

            if action == 'bulk_restore':
                reason = (request.POST.get('reason') or '').strip()
                restored_count = 0
                restore_queryset = PlanningJob.objects.filter(id__in=selected_ids, is_active=False)
                if not _user_is_admin(request.user):
                    skipped = restore_queryset.filter(is_cancelled=True).count()
                    if skipped:
                        messages.warning(
                            request,
                            f'Skipped {skipped} cancelled job(s). Only administrators can restore a cancelled job.',
                        )
                    restore_queryset = restore_queryset.filter(is_cancelled=False)
                for job in restore_queryset:
                    job.is_active = True
                    job.restored_by = request.user
                    job.restored_at = timezone.now()
                    job.restore_reason = reason
                    job.save(update_fields=[
                        'is_active', 'restored_by', 'restored_at', 'restore_reason', 'updated_at',
                        *_clear_cancellation(job),
                    ])
                    _clear_ignored_sku_from_po_docs(job.po_number, job.sku)
                    restored_count += 1
                messages.success(request, f'Bulk restore complete. Restored {restored_count} jobs.')
                return redirect('planning:jobs_archived')

            if action == 'bulk_delete':
                if not _user_is_admin(request.user):
                    messages.error(request, 'Only administrators can permanently delete archived planning jobs.')
                    return redirect('planning:jobs_archived')
                deleted_count = PlanningJob.objects.filter(id__in=selected_ids, is_active=False).delete()[0]
                messages.success(request, f'Bulk delete complete. Deleted {deleted_count} jobs.')
                return redirect('planning:jobs_archived')

        job_id = request.POST.get('job_id')
        try:
            job_id = int(job_id)
        except (TypeError, ValueError):
            messages.error(request, 'Invalid archived planning job selected.')
            return redirect('planning:jobs_archived')

        job = get_object_or_404(PlanningJob, id=job_id, is_active=False)

        if action == 'restore':
            if job.is_cancelled and not _user_is_admin(request.user):
                messages.error(request, 'Only administrators can restore a cancelled planning job.')
                return redirect('planning:jobs_archived')
            reason = (request.POST.get('reason') or '').strip()
            job.is_active = True
            job.restored_by = request.user
            job.restored_at = timezone.now()
            job.restore_reason = reason
            job.save(update_fields=[
                'is_active', 'restored_by', 'restored_at', 'restore_reason', 'updated_at',
                *_clear_cancellation(job),
            ])
            _clear_ignored_sku_from_po_docs(job.po_number, job.sku)
            messages.success(request, f'Planning job {job.jc_number} was restored from archive.')
            return redirect('planning:jobs_archived')

        if action == 'delete':
            if not _user_is_admin(request.user):
                messages.error(request, 'Only administrators can permanently delete archived planning jobs.')
                return redirect('planning:jobs_archived')
            job.delete()
            messages.success(request, f'Planning job {job.jc_number} was permanently deleted.')
            return redirect('planning:jobs_archived')

        messages.error(request, 'Unknown action for archived planning jobs.')
        return redirect('planning:jobs_archived')

    queryset = PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs').filter(is_active=False)

    # view: '' = everything archived, 'cancelled' / 'archived' narrow it down.
    view_filter = (request.GET.get('view') or '').strip()
    if view_filter == 'cancelled':
        queryset = queryset.filter(is_cancelled=True)
    elif view_filter == 'archived':
        queryset = queryset.filter(is_cancelled=False)

    q = (request.GET.get('q') or '').strip()
    status_filter = _normalize_status(request.GET.get('status'), default='')
    department_filter = (request.GET.get('department') or '').strip()
    machine_filter = (request.GET.get('machine') or '').strip()
    from_date = _parse_date_filter(request.GET.get('from_date'))
    to_date = _parse_date_filter(request.GET.get('to_date'))

    if q:
        queryset = queryset.filter(
            Q(jc_number__icontains=q)
            | Q(po_number__icontains=q)
            | Q(sku__icontains=q)
            | Q(job_name__icontains=q)
        )
    if status_filter:
        queryset = queryset.filter(status__in=_planning_status_filter_values(status_filter))
    if department_filter:
        queryset = queryset.filter(department__icontains=department_filter)
    if machine_filter:
        queryset = queryset.filter(machine_name__icontains=machine_filter)
    if from_date:
        queryset = queryset.filter(po_approval_date__gte=from_date)
    if to_date:
        queryset = queryset.filter(po_approval_date__lte=to_date)

    status_rows = queryset.values('status').annotate(total=Count('id')).order_by('status')
    status_counts = {}
    for row in status_rows:
        normalized_status = _normalize_status(row['status'])
        status_counts[normalized_status] = status_counts.get(normalized_status, 0) + (row['total'] or 0)

    paginator = Paginator(queryset, 50)
    page_number = request.GET.get('page')
    jobs = paginator.get_page(page_number)
    return render(
        request,
        'planning/planning_archived_jobs.html',
        {
            'jobs': jobs,
            'status_counts': status_counts,
            'status_choices': PLANNING_STATUSES,
            'can_admin_actions': _user_is_admin(request.user),
            'view_filter': view_filter,
            'cancelled_count': PlanningJob.objects.filter(is_active=False, is_cancelled=True).count(),
            'filters': {
                'q': q,
                'status': status_filter,
                'department': department_filter,
                'machine': machine_filter,
                'view': view_filter,
                'from_date': request.GET.get('from_date', ''),
                'to_date': request.GET.get('to_date', ''),
            },
        },
    )


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def import_planning_sheet(request):
    if request.method == 'POST':
        upload = request.FILES.get('sheet_file')
        if not upload:
            messages.error(request, 'Please choose a CSV file first.')
            return redirect('planning:import_sheet')

        if not upload.name.lower().endswith('.csv'):
            messages.error(request, 'Only CSV file is supported in this first import phase.')
            return redirect('planning:import_sheet')

        decoded = upload.read().decode('utf-8-sig', errors='ignore')
        rows = list(csv.reader(io.StringIO(decoded)))
        header_index = None
        for idx, candidate in enumerate(rows[:8]):
            normalized = {str(col).strip().lower() for col in candidate}
            if 'jc' in normalized and 'job name' in normalized:
                header_index = idx
                break

        if header_index is None:
            messages.error(request, 'Could not detect a valid header row (expected JC and Job Name columns).')
            return redirect('planning:import_sheet')

        header = rows[header_index]
        data_rows = rows[header_index + 1 :]

        imported_count = 0
        updated_count = 0

        for raw_row in data_rows:
            row = {
                header[i]: raw_row[i] if i < len(raw_row) else ''
                for i in range(len(header))
            }
            jc_number = (row.get('JC') or '').strip()
            if not jc_number:
                continue

            defaults = {
                'plan_month': (row.get('Month') or '').strip(),
                'plan_date': _to_date(row.get('Date')),
                'po_approval_date': _to_date(row.get('Approval Date') or row.get('po_approval_date')),
                'delivery_date': _to_date(row.get('Delivery Date')),
                'po_number': (row.get('Po') or '').strip(),
                'sku': (row.get('SKU') or '').strip(),
                'job_name': (row.get('Job Name') or '').strip(),
                'repeat_flag': (row.get('Repeat') or '').strip(),
                'material': (row.get('Material') or '').strip(),
                'color_spec': (row.get('Color') or '').strip(),
                'application': (row.get('Application') or '').strip(),
                'size_w_mm': _to_decimal(row.get('Size W mm')),
                'size_h_mm': _to_decimal(row.get('Size H mm')),
                'size_w_inch': _to_decimal(row.get('Size W Inch')),
                'size_h_inch': _to_decimal(row.get('Size H Inch')),
                'order_qty': _to_int(row.get('Order Qty')),
                'print_pcs': _to_int(row.get('Print Pcs')),
                'ups': _to_decimal(row.get('Ups')),
                'print_sheet_size': (row.get('Print Sheet Size') or '').strip(),
                'print_sheets': _to_int(row.get('Print Sheets')),
                'wastage_sheets': _to_int(row.get('Wastage')),
                'actual_sheet_required': _to_int(row.get('Actual Sheet require')),
                'purchase_sheet_size': (row.get('Purchase Sheet Size') or '').strip(),
                'purchase_sheet_ups': _to_decimal(row.get('Purchase Sheet ups')),
                'purchase_sheet_required': _to_int(row.get('Purchase Sheet require')),
                'pkt_value': _to_decimal(row.get('PKT')),
                'remarks': (row.get('Remarks  ') or '').strip(),
                'requirement': (row.get('Requirement') or '').strip(),
                'front_colors': _to_int(row.get('No. of Clrs Front')),
                'back_colors': _to_int(row.get('No. Of Clrs Back')),
                'total_colors': _to_int(row.get('Total Crls')),
                'total_mr_time_minutes': _to_int(row.get('Total M/R Time (15m/clr)')),
                'front_pass': _to_int(row.get('Front Pass')),
                'back_pass': _to_int(row.get('Back Pass')),
                'planned_total_impressions': _to_int(row.get('Total Impressions')),
                'mi_quantity': _to_int(row.get('MI Quantity 5')),
                'mi_balance': _to_int(row.get('MI Balance')),
                'remaining_sheet': _to_int(row.get('Remaining sheet')),
                'status': (row.get('status') or '').strip(),
                'pr_reference': (row.get('PR') or '').strip(),
                'rejected_qty': _to_int(row.get('Rejected')),
                'balance_qty': _to_int(row.get('Balance')),
                'destination': (row.get('Destination') or '').strip(),
                'unit_cost': _to_decimal(row.get('Cost')),
                'stock_bag': _to_decimal(row.get('Stock Bag')),
                'machine_name': (row.get('Machine Name') or '').strip(),
                'purchase_material_origin': _normalize_purchase_material_origin((row.get('Purchase Material') or '').strip()),
                'stock_qty': _to_decimal(row.get('Stock')),
                'daily_demand': _to_decimal(row.get('Daily Demand')),
                'department': (row.get('Department') or '').strip(),
                'plate_set_no': (row.get('Plate Set No') or '').strip(),
                'aging_days': _to_int(row.get('Aging')),
            }

            job, created = PlanningJob.objects.update_or_create(
                jc_number=jc_number,
                defaults=defaults,
            )

            if created:
                imported_count += 1
            else:
                updated_count += 1

            job.print_runs.all().delete()
            print_rows = []
            for i in range(1, 6):
                print_date = _to_date(row.get(f'Print Date {i}'))
                print_qty = _to_int(row.get(f'Print Qty {i}'))
                wastage_qty = _to_int(row.get(f'Wastage {i}'))
                if print_date or print_qty or wastage_qty:
                    print_rows.append(
                        PlanningPrintRun(
                            planning_job=job,
                            run_index=i,
                            print_date=print_date,
                            print_qty=print_qty,
                            wastage_qty=wastage_qty,
                        )
                    )
            if print_rows:
                PlanningPrintRun.objects.bulk_create(print_rows)

            job.dispatch_runs.all().delete()
            dispatch_rows = []
            for i in range(1, 7):
                idx = f'{i:02d}'
                delivery_date = _to_date(row.get(f'Date Delivery {idx}'))
                dc_no = (row.get(f'DC {idx}') or '').strip()
                delivered_qty = _to_int(row.get(f'Delivered Quantity {idx}'))
                if delivery_date or dc_no or delivered_qty:
                    dispatch_rows.append(
                        PlanningDispatchRun(
                            planning_job=job,
                            dispatch_index=i,
                            delivery_date=delivery_date,
                            dc_no=dc_no,
                            delivered_qty=delivered_qty,
                        )
                    )
            if dispatch_rows:
                PlanningDispatchRun.objects.bulk_create(dispatch_rows)

        messages.success(
            request,
            f'Import completed. New jobs: {imported_count}, updated jobs: {updated_count}.',
        )
        return redirect('planning:jobs')

    return render(request, 'planning/planning_import.html')


@login_required
def planning_job_detail(request, job_id):
    job = get_object_or_404(
        PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs'),
        id=job_id,
    )
    is_repeat_with_changes = (
        (job.repeat_flag or '').lower() == 'repeat'
        and job.has_edits_since_creation
        and job.edited_fields_list
    )
    status_now = _normalize_status(job.status)
    job_card = getattr(job, 'job_card', None)
    can_print_from_job = status_now in {'released', 'in_production', 'completed'}
    can_print_from_card = job_card.can_print_job_card() if job_card else False

    po_approval_date_display = _get_po_approval_date_for_job(job)
    profile = getattr(request.user, 'profile', None)
    user_can_approve_planning = profile.can_approve_planning() if profile else False
    master_data_diffs = get_master_data_field_diffs(job)
    master_data_mismatch = bool(master_data_diffs)
    master_sync_preview = preview_master_sync_calculations(job) if master_data_mismatch else None
    requires_reopen_for_sync = job_requires_reopen_for_master_sync(job) if master_data_mismatch else False
    can_request_master_sync = can_request_master_data_sync(job) and not job.master_sync_requested
    can_apply_master_sync = (
        (user_can_approve_planning or _user_is_admin(request.user))
        and job.master_sync_requested
        and not job.master_data_sync_blocked()
        and not requires_reopen_for_sync
    )
    can_reopen_and_apply_master_sync = (
        (user_can_approve_planning or _user_is_admin(request.user))
        and master_data_mismatch
        and not job.master_data_sync_blocked()
        and requires_reopen_for_sync
    )
    can_dismiss_master_sync = (
        job.master_sync_requested
        and (
            _user_is_admin(request.user)
            or user_can_approve_planning
            or job.master_sync_requested_by_id == request.user.id
        )
    )

    pending_change_request = JobCardChangeRequest.objects.filter(planning_job=job, status='pending').first()
    active_machines = Machine.objects.filter(is_active=True).order_by('name')

    active_recipe = job.sku_recipe
    approved_recipe = job.approved_sku_recipe
    qc_recipe_warning = ''
    if active_recipe and not approved_recipe:
        qc_recipe_warning = (
            'Active SKU recipe exists but is not approved. '
            'QC submission is blocked until the recipe is approved.'
        )
    elif not active_recipe:
        qc_recipe_warning = (
            'No active SKU recipe exists for this job. '
            'QC submission is blocked until a recipe is created and approved.'
        )

    qc_submission_blockers = []
    qc_submission_warnings = []
    if status_now == 'draft' and not qc_recipe_warning:
        qc_submission_blockers = preview_job_qc_submission_blockers(job)
        qc_submission_warnings = preview_job_qc_submission_warnings(job)

    plate_requests = list(job.plate_requests.select_related('requested_by').order_by('-requested_at', '-created_at')[:12])
    latest_cancelled_plate = next((req for req in plate_requests if req.is_cancelled), None)
    from printing_plates.services import plate_request_is_stale_open

    stale_plate_requests = [req for req in plate_requests if plate_request_is_stale_open(req)]

    printing_entries = []
    packing_entries = []
    dispatch_entries = []
    if job_card:
        printing_entries = job_card.productions.filter(is_active=True, entry_type='printing').select_related('machine', 'operator', 'supervisor').order_by('date', 'id')
        packing_entries = job_card.productions.filter(is_active=True, entry_type='packing').select_related('sorter', 'created_by').order_by('date', 'id')
        dispatch_entries = job_card.dispatch_set.filter(is_active=True).select_related('created_by').order_by('dispatch_date', 'id')
    total_print_waste_sheets = sum(entry.waste_sheets or 0 for entry in printing_entries)
    total_sorting_waste_qty = sum(entry.sorting_waste_qty or 0 for entry in packing_entries)

    from core.services import compute_job_card_wastage_metrics
    wastage_metrics = compute_job_card_wastage_metrics(job_card) if job_card else None

    from planning.sku_duplicate_alert import build_sku_duplicate_alert

    sku_duplicate_alert = build_sku_duplicate_alert(job)

    # A plate set parked when this SKU joined an earlier combined layout can be
    # picked back up instead of paying for plate making again.
    retained_plate = None
    if not job.active_merge_item:
        from printing_plates.models import PlateRequest
        from printing_plates.services import get_retained_plate_for_sku

        has_live_plate = PlateRequest.objects.filter(planning_job=job).exclude(
            status=PlateRequest.STATUS_ARCHIVED
        ).exists()
        if not has_live_plate:
            retained_plate = get_retained_plate_for_sku(job.sku)

    return render(
        request,
        'planning/planning_job_detail.html',
        {
            'job': job,
            'retained_plate': retained_plate,
            'recipe': active_recipe,
            'status_now': status_now,
            'po_approval_date_display': po_approval_date_display,
            'can_edit_job': status_now in {'draft', 'pending_qc'},
            'is_repeat_with_changes': is_repeat_with_changes,
            'changed_fields': job.edited_fields_list or [],
            'last_edited_by': job.last_edited_by,
            'last_edited_at': job.last_edited_at,
            'qc_missing_fields': job.qc_missing_fields(),
            'qc_submission_blockers': qc_submission_blockers,
            'qc_submission_warnings': qc_submission_warnings,
            'qc_recipe_warning': qc_recipe_warning,
            'can_print_job_card': can_print_from_job or can_print_from_card,
            'can_admin_delete': _user_is_admin(request.user),
            'user_can_plan': getattr(getattr(request.user, 'profile', None), 'can_plan', lambda: False)(),
            'approved_recipe': job.approved_sku_recipe,
            'master_data_diffs': master_data_diffs,
            'master_data_mismatch': master_data_mismatch,
            'master_sync_preview': master_sync_preview,
            'requires_reopen_for_sync': requires_reopen_for_sync,
            'can_request_master_sync': can_request_master_sync,
            'can_apply_master_sync': can_apply_master_sync,
            'can_reopen_and_apply_master_sync': can_reopen_and_apply_master_sync,
            'can_dismiss_master_sync': can_dismiss_master_sync,
            'user_can_approve_planning': user_can_approve_planning,
            'pending_change_request': pending_change_request,
            'active_machines': active_machines,
            'plate_requests': plate_requests,
            'stale_plate_requests': stale_plate_requests,
            'latest_cancelled_plate': latest_cancelled_plate,
            'printing_entries': printing_entries,
            'packing_entries': packing_entries,
            'dispatch_entries': dispatch_entries,
            'total_print_waste_sheets': total_print_waste_sheets,
            'total_sorting_waste_qty': total_sorting_waste_qty,
            'wastage_metrics': wastage_metrics,
            'sku_duplicate_alert': sku_duplicate_alert,
        },
    )


@login_required
@permission_required('can_plan')
def planning_job_plate_request_cancel(request, job_id):
    if request.method != 'POST':
        return redirect('planning:job_detail', job_id=job_id)

    from django.core.exceptions import ValidationError
    from printing_plates.models import PlateRequest
    from printing_plates.services import cancel_plate_request

    job = get_object_or_404(PlanningJob, id=job_id)
    plate_request_id = (request.POST.get('plate_request_id') or '').strip()
    reason = (request.POST.get('cancel_reason') or '').strip()

    if not plate_request_id:
        messages.error(request, 'Plate request is required.')
        return redirect('planning:job_detail', job_id=job_id)

    plate_request = get_object_or_404(PlateRequest, pk=plate_request_id, planning_job=job)

    try:
        cancel_plate_request(plate_request, actor=request.user, reason=reason)
    except ValidationError as exc:
        message = exc.messages[0] if getattr(exc, 'messages', None) else str(exc)
        messages.error(request, message)
        return redirect('planning:job_detail', job_id=job_id)

    messages.success(
        request,
        f'Plate request #{plate_request.pk} cancelled for {job.jc_number} (plates not required).',
    )
    return redirect('planning:job_detail', job_id=job_id)


@login_required
@permission_required('can_plan')
@transaction.atomic
def planning_job_master_sync(request, job_id):
    if request.method != 'POST':
        return redirect('planning:job_detail', job_id=job_id)

    job = get_object_or_404(PlanningJob, id=job_id)
    action = (request.POST.get('action') or '').strip()
    reason = (request.POST.get('reason') or '').strip()
    next_url = (request.POST.get('next') or '').strip()
    profile = getattr(request.user, 'profile', None)
    user_can_approve_planning = profile.can_approve_planning() if profile else False

    def _redirect():
        if next_url:
            return redirect(next_url)
        return redirect('planning:job_detail', job_id=job.id)

    try:
        if action == 'request_master_sync':
            pre_sync_diffs = get_master_data_field_diffs(job)
            request_master_data_sync(job, actor=request.user, reason=reason)
            messages.success(
                request,
                f'Master data sync requested for {job.jc_number}. An approver can apply the update from this job or the Approval Queue.',
            )
            plate_warning = get_plate_remake_warning_for_job_sync(job, pre_sync_diffs)
            if plate_warning:
                messages.warning(request, plate_warning)
        elif action == 'apply_master_sync':
            if not (user_can_approve_planning or _user_is_admin(request.user)):
                messages.error(request, 'You do not have permission to apply master data sync.')
                return _redirect()
            if not job.master_sync_requested:
                messages.error(request, 'No pending master data sync request exists for this job.')
                return _redirect()
            if job_requires_reopen_for_master_sync(job):
                messages.error(
                    request,
                    'This job card is locked. Use Reopen & Apply Master Sync instead.',
                )
                return _redirect()
            pre_sync_diffs = get_master_data_field_diffs(job)
            job, result = apply_master_data_sync(job, actor=request.user)
            updated_fields = result.get('updated_fields') or []
            if updated_fields:
                messages.success(
                    request,
                    f'{job.jc_number} synced with approved SKU master ({len(updated_fields)} field(s) updated).',
                )
            else:
                messages.info(request, f'{job.jc_number} already matches approved SKU master data.')
            plate_warning = get_plate_remake_warning_for_job_sync(job, pre_sync_diffs)
            if plate_warning:
                messages.warning(request, plate_warning)
            if updated_fields and not result.get('job_card_refreshed'):
                messages.warning(
                    request,
                    'Planning job was updated but the linked job card is locked. Reopen the job card workflow if sheet values on the card must change.',
                )
        elif action == 'reopen_and_apply_master_sync':
            if not (user_can_approve_planning or _user_is_admin(request.user)):
                messages.error(request, 'You do not have permission to reopen and apply master data sync.')
                return _redirect()
            pre_sync_diffs = get_master_data_field_diffs(job)
            job, result = reopen_and_apply_master_data_sync(
                job,
                actor=request.user,
                reason=reason or 'Reopen and apply SKU master sync',
            )
            updated_fields = result.get('updated_fields') or []
            if updated_fields:
                messages.success(
                    request,
                    f'{job.jc_number} reopened and synced with approved SKU master ({len(updated_fields)} field(s) updated). Re-submit through QC when ready.',
                )
            else:
                messages.info(request, f'{job.jc_number} already matches approved SKU master data.')
            plate_warning = get_plate_remake_warning_for_job_sync(job, pre_sync_diffs)
            if plate_warning:
                messages.warning(request, plate_warning)
        elif action == 'dismiss_master_sync':
            if not (
                _user_is_admin(request.user)
                or user_can_approve_planning
                or job.master_sync_requested_by_id == request.user.id
            ):
                messages.error(request, 'You do not have permission to dismiss this master sync request.')
                return _redirect()
            dismiss_master_data_sync_request(job, actor=request.user)
            messages.info(request, f'Master data sync request for {job.jc_number} was dismissed.')
        else:
            messages.error(request, 'Unknown master data sync action.')
    except ValueError as exc:
        messages.error(request, str(exc))

    return _redirect()


@login_required
@permission_required('can_plan')
def planning_job_edit(request, job_id):
    job = get_object_or_404(PlanningJob, id=job_id)
    _repair_rejected_job_status(job)
    current_status = _normalize_status(job.status)

    locked_statuses = {'qc_approved', 'released', 'in_production', 'completed'}
    if current_status in locked_statuses and not job.change_request_pending:
        messages.error(request, 'QC approved and released records are locked. Reopen the job before editing.')
        return redirect('planning:job_detail', job_id=job.id)
    if current_status in locked_statuses and job.change_request_pending:
        messages.warning(request, 'This job has a pending change request. You may edit the job details and save changes to resolve the request.')

    if request.method == 'POST':
        form = PlanningJobEditForm(request.POST, instance=job)
        if form.is_valid():
            edited = form.save(commit=False)
            edited.status = _normalize_status(edited.status)
            edited.job_card_version = (job.job_card_version or 1) + 1
            
            # Detect changes for repeat jobs
            if (job.repeat_flag or '').lower() == 'repeat':
                changed_fields = []
                edit_fields = [
                    'delivery_date', 'wastage_sheets', 'plate_set_no', 'machine_name',
                    'planned_total_impressions', 'purchase_material_origin', 'destination',
                    'remarks', 'requirement', 'status',
                ]
                for field in edit_fields:
                    old_val = getattr(job, field, None)
                    new_val = getattr(edited, field, None)
                    if str(old_val) != str(new_val):
                        changed_fields.append(field)
                
                if changed_fields:
                    edited.has_edits_since_creation = True
                    edited.edited_fields_list = changed_fields
                    edited.last_edited_by = request.user
                    edited.last_edited_at = timezone.now()
            
            try:
                edited.save()
            except ValidationError as exc:
                error_dict = getattr(exc, 'message_dict', None)
                if error_dict:
                    for field_name, field_errors in error_dict.items():
                        for error_message in field_errors:
                            if field_name == '__all__' or field_name not in form.fields:
                                form.add_error(None, error_message)
                            else:
                                form.add_error(field_name, error_message)
                else:
                    form.add_error(None, str(exc))
            else:
                if request.POST.get('submit_to_qc'):
                    return _submit_job_to_qc(
                        edited,
                        request,
                        next_url=reverse('planning:job_detail', kwargs={'job_id': edited.id}),
                    )
                messages.success(request, f'Planning job {edited.jc_number} updated.')
                if edited.has_edits_since_creation and (edited.repeat_flag or '').lower() == 'repeat':
                    messages.info(request, f"Changes detected and flagged for production team: {', '.join(edited.edited_fields_list)}")
                return redirect('planning:job_detail', job_id=edited.id)
    else:
        form = PlanningJobEditForm(instance=job)

    approved_recipe = SkuRecipe.objects.filter(
        sku__iexact=job.sku, is_active=True, master_data_status='approved',
    ).first()
    active_recipe = SkuRecipe.objects.filter(
        sku__iexact=job.sku, is_active=True,
    ).first()
    qc_recipe_warning = ''
    if active_recipe and not approved_recipe:
        qc_recipe_warning = (
            'Active SKU recipe exists but is not approved. ' \
            'QC submission is blocked until the recipe is approved.'
        )
    elif not active_recipe:
        qc_recipe_warning = (
            'No active SKU recipe exists for this job. ' \
            'QC submission is blocked until a recipe is created and approved.'
        )

    po_approval_date_display = _get_po_approval_date_for_job(job)
    master_data_diffs = get_master_data_field_diffs(job)
    qc_submission_blockers = []
    qc_submission_warnings = []
    if job.status == 'draft' and not qc_recipe_warning:
        qc_submission_blockers = preview_job_qc_submission_blockers(job)
        qc_submission_warnings = preview_job_qc_submission_warnings(job)

    return render(
        request,
        'planning/planning_job_edit.html',
        {
            'job': job,
            'form': form,
            'recipe': active_recipe,
            'qc_recipe_warning': qc_recipe_warning,
            'qc_submission_blockers': qc_submission_blockers,
            'qc_submission_warnings': qc_submission_warnings,
            'can_admin_delete': _user_is_admin(request.user),
            'po_approval_date_display': po_approval_date_display,
            'master_data_diffs': master_data_diffs,
            'master_data_mismatch': bool(master_data_diffs),
        },
    )


def _submit_job_to_qc(job, request, next_url=None):
    current_status = _normalize_status(job.status)
    if current_status != 'draft':
        messages.error(request, 'Job must be in draft before QC submission.')
        return redirect(next_url or 'planning:job_detail', job_id=job.id)

    blockers = get_job_qc_submission_blockers(job)
    if blockers:
        for error_message in blockers:
            messages.error(request, error_message)
        return redirect(next_url or 'planning:job_edit', job_id=job.id)

    job.status = 'pending_qc'
    job.issued_to_production = False
    try:
        with transaction.atomic():
            job.save(update_fields=['status', 'issued_to_production', 'updated_at'])
            try:
                sync_job_card_for_planning_status(job, 'pending_qc', request.user)
            except ValidationError as exc:
                transaction.set_rollback(True)
                error_dict = getattr(exc, 'message_dict', None)
                if error_dict:
                    for field_errors in error_dict.values():
                        for error_message in field_errors:
                            messages.error(request, error_message)
                else:
                    messages.error(request, str(exc))
                return redirect(next_url or 'planning:job_edit', job_id=job.id)
    except ValidationError as exc:
        error_dict = getattr(exc, 'message_dict', None)
        if error_dict:
            for field_errors in error_dict.values():
                for error_message in field_errors:
                    messages.error(request, error_message)
        else:
            messages.error(request, str(exc))
        return redirect(next_url or 'planning:job_edit', job_id=job.id)

    messages.success(request, f'Job status updated: draft -> pending_qc.')
    return redirect(next_url or 'planning:job_detail', job_id=job.id)


@login_required
@permission_required('can_plan')
@transaction.atomic
def planning_job_status_update(request, job_id):
    if request.method != 'POST':
        return redirect('planning:job_detail', job_id=job_id)

    job = get_object_or_404(PlanningJob, id=job_id)
    current_status = _normalize_status(job.status)
    transition = (request.POST.get('transition') or '').strip()
    next_url = (request.POST.get('next') or '').strip()

    transitions = {
        'submit_review': ('draft', 'pending_qc'),
        'submit_qc': ('draft', 'pending_qc'),
        'reopen': ('completed', 'draft'),
    }
    if transition not in transitions:
        messages.error(request, 'Unknown status transition request.')
        return redirect(next_url) if next_url else redirect('planning:job_detail', job_id=job.id)

    required_from, target_status = transitions[transition]
    if current_status == target_status:
        messages.info(request, f'Job already in {target_status} status.')
        return redirect(next_url) if next_url else redirect('planning:job_detail', job_id=job.id)

    if required_from and current_status != required_from:
        messages.error(request, f'Transition not allowed from {current_status} to {target_status}.')
        return redirect(next_url) if next_url else redirect('planning:job_detail', job_id=job.id)

    if target_status == 'pending_qc':
        return _submit_job_to_qc(
            job,
            request,
            next_url=next_url or reverse('planning:job_detail', kwargs={'job_id': job.id}),
        )

    job.status = target_status
    if target_status == 'pending_qc':
        job.issued_to_production = False
    if target_status == 'draft':
        job.issued_to_production = False

    try:
        job.save(update_fields=['status', 'issued_to_production', 'updated_at'])
        if target_status == 'pending_qc':
            try:
                sync_job_card_for_planning_status(job, target_status, request.user)
            except ValidationError as exc:
                transaction.set_rollback(True)
                error_dict = getattr(exc, 'message_dict', None)
                if error_dict:
                    for field_errors in error_dict.values():
                        for error_message in field_errors:
                            messages.error(request, error_message)
                else:
                    messages.error(request, str(exc))
                return redirect(next_url) if next_url else redirect('planning:job_detail', job_id=job.id)
    except ValidationError as exc:
        error_dict = getattr(exc, 'message_dict', None)
        if error_dict:
            for field_errors in error_dict.values():
                for error_message in field_errors:
                    messages.error(request, error_message)
        else:
            messages.error(request, str(exc))
        return redirect(next_url) if next_url else redirect('planning:job_detail', job_id=job.id)

    messages.success(request, f'Job status updated: {current_status} -> {target_status}.')
    return redirect(next_url) if next_url else redirect('planning:job_detail', job_id=job.id)


@login_required
@permission_required('can_plan')
def planning_job_priority_update(request, job_id):
    from django.http import JsonResponse
    if request.method != 'POST':
        return JsonResponse({'ok': False, 'error': 'POST request required'}, status=405)

    job = get_object_or_404(PlanningJob, id=job_id)
    try:
        priority_val = int(request.POST.get('priority', 1))
        if priority_val in [choice[0] for choice in PlanningJob.PRIORITY_CHOICES]:
            job.priority = priority_val
            job.save(update_fields=['priority', 'updated_at'])
            return JsonResponse({
                'ok': True,
                'priority': priority_val,
                'priority_display': job.get_priority_display()
            })
        else:
            return JsonResponse({'ok': False, 'error': 'Invalid priority value'}, status=400)
    except Exception as e:
        return JsonResponse({'ok': False, 'error': str(e)}, status=400)


@login_required
@permission_required('can_view_jobcard')
def planning_job_card_print(request, job_id):
    job = get_object_or_404(
        PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs'),
        id=job_id,
    )
    status_now = _normalize_status(job.status)
    job_card = getattr(job, 'job_card', None)
    can_print_from_job = status_now in {'released', 'in_production', 'completed'}
    can_print_from_card = job_card.can_print_job_card() if job_card else False
    if not (can_print_from_job or can_print_from_card):
        messages.error(request, 'Job card print is available only after production approval.')
        return redirect('planning:job_detail', job_id=job.id)

    missing_qc_fields = job.qc_missing_fields()
    if missing_qc_fields:
        messages.error(request, f'Job card cannot be printed until QC fields are completed: {", ".join(missing_qc_fields)}.')
        return redirect('planning:job_detail', job_id=job.id)

    recipe = SkuRecipe.objects.filter(sku=job.sku).first()

    def _display_user_identity(user):
        if not user:
            return ''
        full_name = (user.get_full_name() or '').strip()
        if full_name:
            return full_name
        return (user.username or '').strip()

    def _workflow_actor_for_status(target_status):
        if not job_card:
            return ''
        logs = ChangeLog.objects.filter(
            entity_type='job_card',
            record_id=job_card.pk,
        ).select_related('changed_by').order_by('-created_at')
        for log in logs:
            field_changes = log.field_changes if isinstance(log.field_changes, dict) else {}
            status_change = field_changes.get('status') if isinstance(field_changes, dict) else None
            if not isinstance(status_change, dict):
                continue
            to_status = str(status_change.get('to') or '').strip().lower()
            if to_status == target_status and log.changed_by:
                return _display_user_identity(log.changed_by)
        return ''

    def _workflow_date_for_status(target_status):
        if not job_card:
            return None
        logs = ChangeLog.objects.filter(
            entity_type='job_card',
            record_id=job_card.pk,
        ).order_by('-created_at')
        for log in logs:
            field_changes = log.field_changes if isinstance(log.field_changes, dict) else {}
            status_change = field_changes.get('status') if isinstance(field_changes, dict) else None
            if not isinstance(status_change, dict):
                continue
            to_status = str(status_change.get('to') or '').strip().lower()
            if to_status == target_status and log.created_at:
                return log.created_at.date()
        return None

    def _mm_int_string(value):
        if value is None:
            return None
        try:
            return str(int(Decimal(str(value))))
        except Exception:
            return str(value)

    repeat_flag = (job.repeat_flag or '').strip().lower()
    is_repeat = 'repeat' in repeat_flag or repeat_flag in {'r', 'old', 'existing'}
    if is_repeat:
        production_type_tag = 'REPEAT'
    else:
        production_type_tag = 'NEW'

    recipe_material = (getattr(recipe, 'material', '') or '').strip() if recipe else ''
    material_type_clean = (job.material_display or '').strip() or recipe_material or '-'
    color_spec_clean = (job.color_spec_display or '').strip() or '-'

    special_notes = []
    if recipe and not is_repeat:
        recipe_special = (getattr(recipe, 'special_instructions', '') or '').strip()
        if recipe_special:
            special_notes.append(recipe_special)
    requirement_note = (job.requirement or '').strip()
    if requirement_note and requirement_note not in special_notes:
        special_notes.append(requirement_note)
    special_instructions_text = ' | '.join(special_notes) if special_notes else '-'

    po_approval_date = _get_po_approval_date_for_job(job)

    prepared_by_display = _display_user_identity(job.last_edited_by or job.created_by)
    checked_by_display = _workflow_actor_for_status('qc_approved')
    approved_by_display = _workflow_actor_for_status('production_approved')

    def _pdf_filename(jc_number):
        if not jc_number:
            return 'JOB-CARD'
        normalized = str(jc_number).strip().upper()
        parts = [part for part in normalized.split('-') if part]
        if 'UPP' in parts:
            return '-'.join(parts)
        if len(parts) == 4 and parts[0] == 'JC':
            return '-'.join([parts[0], parts[1], parts[2], 'UPP', parts[3]])
        if len(parts) >= 4 and parts[0] == 'JC' and parts[-2] != 'UPP':
            return '-'.join(parts[:-1] + ['UPP', parts[-1]])
        return normalized

    job_scan_url = request.build_absolute_uri(reverse('planning:job_card_print', kwargs={'job_id': job.id}))
    qr_base64 = _build_qr_image_base64(job_scan_url)
    job_qr_data_uri = f'data:image/png;base64,{qr_base64}' if qr_base64 else None
    pdf_filename = _pdf_filename(job.jc_number)

    return render(
        request,
        'Job Card.html',
        {
            'job': job,
            'recipe': recipe,
            'merge': build_job_card_merge_context(job),
            'job_scan_url': job_scan_url,
            'job_qr_data_uri': job_qr_data_uri,
            'production_type_tag': production_type_tag,
            'plan_date': job.plan_date,
            'po_approval_date_display': po_approval_date,
            'material_type_clean': material_type_clean,
            'color_spec_clean': color_spec_clean,
            'size_w_mm_int': _mm_int_string(job.size_w_mm_display),
            'size_h_mm_int': _mm_int_string(job.size_h_mm_display),
            'special_instructions_text': special_instructions_text,
            'prepared_by_display': prepared_by_display,
            'checked_by_display': checked_by_display,
            'approved_by_display': approved_by_display,
            'pdf_filename': pdf_filename,
        },
    )


@login_required
@permission_required('can_view_jobcard')
def planning_job_card_pdf(request, job_id):
    job = get_object_or_404(
        PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs'),
        id=job_id,
    )
    status_now = _normalize_status(job.status)
    job_card = getattr(job, 'job_card', None)
    can_print_from_job = status_now in {'released', 'in_production', 'completed'}
    can_print_from_card = job_card.can_print_job_card() if job_card else False
    if not (can_print_from_job or can_print_from_card):
        messages.error(request, 'Job card print is available only after production approval.')
        return redirect('planning:job_detail', job_id=job.id)

    missing_qc_fields = job.qc_missing_fields()
    if missing_qc_fields:
        messages.error(request, f'Job card cannot be downloaded until QC fields are completed: {", ".join(missing_qc_fields)}.')
        return redirect('planning:job_detail', job_id=job.id)

    def _pdf_filename(jc_number):
        if not jc_number:
            return 'JOB-CARD'
        normalized = str(jc_number).strip().upper()
        parts = [part for part in normalized.split('-') if part]
        if 'UPP' in parts:
            return '-'.join(parts)
        if len(parts) == 4 and parts[0] == 'JC':
            return '-'.join([parts[0], parts[1], parts[2], 'UPP', parts[3]])
        if len(parts) >= 4 and parts[0] == 'JC' and parts[-2] != 'UPP':
            return '-'.join(parts[:-1] + ['UPP', parts[-1]])
        return normalized

    pdf_filename = _pdf_filename(job.jc_number)
    job_scan_url = request.build_absolute_uri(reverse('planning:job_detail', kwargs={'job_id': job.id}))

    try:
        pdf_bytes = _build_job_card_pdf_bytes(job, job_scan_url)
    except RuntimeError as exc:
        messages.error(request, str(exc))
        return redirect('planning:job_card_print', job_id=job.id)

    response = HttpResponse(pdf_bytes, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{pdf_filename}.pdf"'
    return response


@login_required
def planning_job_history_report_pdf(request, job_id):
    """
    Full recorded history for this JC (planning/PO reference, SKU master data,
    plate requests, printing/packing/dispatch entries) as a downloadable PDF.
    Unlike the Job Card (a blank traveler to fill by hand), this reports what
    actually happened on the job — for review, audit, or sharing.
    """
    job = get_object_or_404(
        PlanningJob.objects.select_related('job_card').prefetch_related('plate_requests'),
        id=job_id,
    )

    from .services import build_job_history_report_pdf_bytes

    try:
        pdf_bytes = build_job_history_report_pdf_bytes(job)
    except RuntimeError as exc:
        messages.error(request, str(exc))
        return redirect('planning:job_detail', job_id=job.id)

    filename = f'{job.jc_number or "JOB"}-HISTORY-REPORT.pdf'
    response = HttpResponse(pdf_bytes, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response


@login_required
@permission_required('can_view_approval_queue')
def approval_queue(request):
    planning_jobs = job_card_queue_queryset('planning')
    qc_jobs = job_card_queue_queryset('qc')
    pm_jobs = job_card_queue_queryset('production_manager')
    release_jobs = job_card_queue_queryset('production')
    queue_q = (request.GET.get('q') or '').strip()

    profile = getattr(request.user, 'profile', None)
    user_can_approve_planning = profile.can_approve_planning() if profile else False
    user_can_approve_qc = profile.can_approve_qc() if profile else False
    user_can_approve_pm = profile.can_approve_pm() if profile else False
    user_can_release = user_can_approve_pm or (profile.can_plan() if profile else False)

    master_sync_requests = PlanningJob.objects.filter(
        is_active=True,
        master_sync_requested=True,
    ).select_related('master_sync_requested_by').order_by('-master_sync_requested_at')
    split_requests = _pending_po_split_requests()
    pending_wastage_machine_change_requests = JobCardChangeRequest.objects.filter(
        status='pending'
    ).select_related('planning_job', 'requested_by').order_by('-requested_at')

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        job_card_id = (request.POST.get('job_card_id') or '').strip()
        planning_job_id = (request.POST.get('planning_job_id') or '').strip()
        reason = (request.POST.get('reason') or request.POST.get('change_reason') or '').strip()

        def _get_job_card():
            if not job_card_id:
                raise ValueError('Job Card is required.')
            return get_object_or_404(JobCard, pk=job_card_id, is_active=True)

        def _get_planning_job():
            if not planning_job_id:
                raise ValueError('Planning job is required.')
            return get_object_or_404(PlanningJob, pk=planning_job_id, is_active=True)

        def _sync_status(job_card, transition_name):
            return execute_job_card_action(job_card, transition_name, actor=request.user, reason=reason)

        if action in {'approve_split_request', 'archive_split_request'}:
            if not user_can_approve_pm:
                messages.error(request, 'You do not have permission to manage PO split requests.')
                return redirect('planning:approval_queue')
            po_doc_id = (request.POST.get('po_doc_id') or '').strip()
            request_id = (request.POST.get('request_id') or '').strip()
            try:
                po_doc = PoDocument.objects.get(id=int(po_doc_id))
            except (TypeError, ValueError, PoDocument.DoesNotExist):
                messages.error(request, 'Invalid PO split request selected.')
                return redirect('planning:approval_queue')
            status = 'approved' if action == 'approve_split_request' else 'archived'
            if _update_po_split_request(po_doc, request_id, status, request.user):
                messages.success(request, f'PO split request {status}.')
            else:
                messages.error(request, 'PO split request was not found or already resolved.')
            return redirect('planning:approval_queue')



        if action == 'apply_master_sync':
            if not user_can_approve_planning and not _user_is_admin(request.user):
                messages.error(request, 'You do not have permission to apply master data sync.')
                return redirect('planning:approval_queue')
            try:
                planning_job = _get_planning_job()
                if not planning_job.master_sync_requested:
                    messages.error(request, f'No pending master sync request for {planning_job.jc_number}.')
                    return redirect('planning:approval_queue')
                if job_requires_reopen_for_master_sync(planning_job):
                    messages.error(
                        request,
                        f'{planning_job.jc_number} is locked. Use Reopen & Apply from the job detail page.',
                    )
                    return redirect('planning:approval_queue')
                planning_job, result = apply_master_data_sync(planning_job, actor=request.user)
                updated_fields = result.get('updated_fields') or []
                if updated_fields:
                    messages.success(
                        request,
                        f'{planning_job.jc_number} synced with approved SKU master ({len(updated_fields)} field(s) updated).',
                    )
                else:
                    messages.info(request, f'{planning_job.jc_number} already matches approved SKU master data.')
                if updated_fields and not result.get('job_card_refreshed'):
                    messages.warning(
                        request,
                        f'{planning_job.jc_number}: job card is locked; reopen workflow if the printed card must change.',
                    )
            except ValueError as exc:
                messages.error(request, str(exc))
            except (ValueError, Http404):
                messages.error(request, 'Invalid planning job selected.')
            return redirect('planning:approval_queue')

        if action == 'dismiss_master_sync':
            try:
                planning_job = _get_planning_job()
                if not (
                    _user_is_admin(request.user)
                    or user_can_approve_planning
                    or planning_job.master_sync_requested_by_id == request.user.id
                ):
                    messages.error(request, 'You do not have permission to dismiss this master sync request.')
                    return redirect('planning:approval_queue')
                dismiss_master_data_sync_request(planning_job, actor=request.user)
                messages.info(request, f'Master data sync request for {planning_job.jc_number} was dismissed.')
            except (ValueError, Http404):
                messages.error(request, 'Invalid planning job selected.')
            return redirect('planning:approval_queue')

        if action in {
            'approve_planning', 'reject_planning',
            'approve_qc', 'reject_qc',
            'approve_pm', 'reject_pm', 'release_for_production',
        }:
            if action in {'reject_planning', 'reject_qc', 'reject_pm'} and not reason:
                messages.error(request, 'Rejection reason is required.')
                return redirect('planning:approval_queue')

            if action in {'approve_planning', 'reject_planning'} and not user_can_approve_planning:
                messages.error(request, 'You do not have permission to approve planning jobs.')
                return redirect('planning:approval_queue')
            if action in {'approve_qc', 'reject_qc'} and not user_can_approve_qc:
                messages.error(request, 'You do not have permission to approve QC jobs.')
                return redirect('planning:approval_queue')
            if action in {'approve_pm', 'reject_pm'} and not user_can_approve_pm:
                messages.error(request, 'You do not have permission to approve production manager jobs.')
                return redirect('planning:approval_queue')
            if action == 'release_for_production' and not user_can_release:
                messages.error(request, 'You do not have permission to release jobs for production.')
                return redirect('planning:approval_queue')

            job_card = _get_job_card()
            if action == 'release_for_production':
                plate_set_no = (request.POST.get('plate_set_no') or '').strip()
                planning_job = getattr(job_card, 'planning_job', None)
                if plate_set_no and planning_job:
                    planning_job.plate_set_no = plate_set_no
                    planning_job.save(update_fields=['plate_set_no', 'updated_at'])

            try:
                _sync_status(job_card, action)
                messages.success(request, f'Job Card {job_card.job_card_no} moved successfully.')
            except ValidationError as exc:
                error_dict = getattr(exc, 'message_dict', None)
                if error_dict:
                    for field_errors in error_dict.values():
                        for error_message in field_errors:
                            messages.error(request, error_message)
                else:
                    messages.error(request, str(exc))
            return redirect('planning:approval_queue')

    if queue_q:
        queue_filter = (
            Q(job_card_no__icontains=queue_q)
            | Q(PO_No__icontains=queue_q)
            | Q(SKU__icontains=queue_q)
            | Q(planning_job__po_number__icontains=queue_q)
            | Q(planning_job__sku__icontains=queue_q)
        )
        planning_jobs = planning_jobs.filter(queue_filter)
        qc_jobs = qc_jobs.filter(queue_filter)
        pm_jobs = pm_jobs.filter(queue_filter)
        release_jobs = release_jobs.filter(queue_filter)

    planning_jobs = planning_jobs.order_by('-updated_at', '-id')[:300]
    qc_jobs = qc_jobs.order_by('-updated_at', '-id')[:300]
    pm_jobs = pm_jobs.order_by('-updated_at', '-id')[:300]
    release_jobs = release_jobs.order_by('-updated_at', '-id')[:300]

    job_card_ids = set()
    for qs in [planning_jobs, qc_jobs, pm_jobs, release_jobs]:
        job_card_ids.update(qs.values_list('id', flat=True))

    status_logs = {}
    if job_card_ids:
        logs = ChangeLog.objects.filter(
            entity_type='job_card',
            record_id__in=job_card_ids,
        ).select_related('changed_by').order_by('-created_at')
        for log in logs:
            if log.record_id in status_logs:
                continue
            field_changes = log.field_changes if isinstance(log.field_changes, dict) else {}
            status_change = field_changes.get('status') if isinstance(field_changes, dict) else None
            if not isinstance(status_change, dict):
                continue
            status_logs[log.record_id] = log

    def _display_user_identity(user):
        if not user:
            return ''
        full_name = (user.get_full_name() or '').strip()
        if full_name:
            return full_name
        return (user.username or '').strip()

    def _attach_status_audit(qs):
        for job in qs:
            log = status_logs.get(job.id)
            if log:
                job.status_changed_by = _display_user_identity(log.changed_by)
                job.status_changed_at = log.created_at
            else:
                job.status_changed_by = ''
                job.status_changed_at = None

    _attach_status_audit(planning_jobs)
    _attach_status_audit(qc_jobs)
    _attach_status_audit(pm_jobs)
    _attach_status_audit(release_jobs)

    from printing_plates.services import get_open_planning_plate_requests_blocking_release_bulk

    release_blocking_map = get_open_planning_plate_requests_blocking_release_bulk(
        [getattr(jc, 'planning_job', None) for jc in release_jobs]
    )
    for job_card in release_jobs:
        planning_job = getattr(job_card, 'planning_job', None)
        job_card.release_blocked_open_plate = release_blocking_map.get(planning_job.id) if planning_job else None

    pending_qc_jobs_count = qc_jobs.count()
    approved_qc_jobs_count = JobCard.objects.filter(is_active=True, status='qc_approved').count()
    pending_pm_jobs_count = pm_jobs.count()
    released_jobs_count = JobCard.objects.filter(is_active=True, status='released').count()

    # Collecting the set of SKU keys with any PO/WO activity doesn't need the full
    # OCR-near-duplicate fuzzy merge that _po_payload_items() does (SequenceMatcher
    # over every item pair, per document) — that pass exists to avoid double-counting
    # near-identical extracted rows on review screens, not to determine "does this SKU
    # exist anywhere". Exact-key dedup + ignored-SKU filtering is enough here and avoids
    # an O(n^2) fuzzy-match scan across every PO document on every approval queue load.
    active_sku_keys = set()
    for payload in PoDocument.objects.exclude(extracted_payload__isnull=True).values_list('extracted_payload', flat=True):
        payload = payload or {}
        ignored = {_sku_key(s) for s in (payload.get('new_skus_ignored') or []) if s}
        deduped_items, _ = _deduplicate_po_items_by_sku(payload.get('items') or [])
        for item in deduped_items:
            key = _sku_key(item.get('sku'))
            if key and key not in ignored:
                active_sku_keys.add(key)
    active_sku_keys.update(
        _sku_key(sku)
        for sku in PlanningJob.objects.filter(is_active=True).values_list('sku', flat=True)
        if sku
    )

    if active_sku_keys:
        sku_query = Q()
        for sku in active_sku_keys:
            sku_query |= Q(sku__iexact=sku)
        pending_sku_approval_count = SkuRecipe.objects.filter(is_active=True, master_data_status='pending_review').filter(sku_query).count()
        sku_reviewed_count = SkuRecipe.objects.filter(is_active=True, master_data_status='reviewed').filter(sku_query).count()
        sku_approved_count = SkuRecipe.objects.filter(is_active=True, master_data_status='approved').filter(sku_query).count()
    else:
        pending_sku_approval_count = 0
        sku_reviewed_count = 0
        sku_approved_count = 0

    for job in master_sync_requests:
        job.master_data_diffs = get_master_data_field_diffs(job)
        job.requires_reopen_for_sync = job_requires_reopen_for_master_sync(job)

    from planning.sku_duplicate_alert import attach_sku_duplicate_alerts_to_job_cards, attach_sku_duplicate_alerts_to_jobs

    attach_sku_duplicate_alerts_to_job_cards(planning_jobs)
    attach_sku_duplicate_alerts_to_job_cards(qc_jobs)
    attach_sku_duplicate_alerts_to_job_cards(pm_jobs)
    attach_sku_duplicate_alerts_to_job_cards(release_jobs)
    attach_sku_duplicate_alerts_to_jobs(master_sync_requests)

    context = {
        'planning_jobs': planning_jobs,
        'qc_jobs': qc_jobs,
        'pm_jobs': pm_jobs,
        'release_jobs': release_jobs,
        'master_sync_requests': master_sync_requests,
        'split_requests': split_requests,
        'pending_wastage_machine_change_requests': pending_wastage_machine_change_requests,
        'queue_q': queue_q,
        'pending_qc_jobs_count': pending_qc_jobs_count,
        'approved_qc_jobs_count': approved_qc_jobs_count,
        'pending_pm_jobs_count': pending_pm_jobs_count,
        'released_jobs_count': released_jobs_count,
        'pending_sku_approval_count': pending_sku_approval_count,
        'sku_reviewed_count': sku_reviewed_count,
        'sku_approved_count': sku_approved_count,
        'user_can_approve_planning': user_can_approve_planning,
        'user_can_approve_qc': user_can_approve_qc,
        'user_can_approve_pm': user_can_approve_pm,
        'user_can_release': user_can_release,
    }
    context['can_admin_actions'] = _user_is_admin(request.user)
    return render(request, 'planning/approval_queue.html', context)


@login_required
@permission_required('can_edit_jobcard')
def planning_scan(request):
    if request.method == 'POST':
        raw_code = (request.POST.get('scan_code') or '').strip()
        if not raw_code:
            messages.error(request, 'Scan code cannot be empty.')
            return redirect('planning:scan')

        # QR may contain full URL, plain JC number, or prefixed JC field.
        parsed = raw_code
        if '/scan/open/' in parsed:
            parsed = parsed.rsplit('/scan/open/', 1)[-1]
        if '?' in parsed:
            parsed = parsed.split('?', 1)[0]
        parsed = parsed.replace('JC:', '').strip().strip('/')

        job = PlanningJob.objects.filter(jc_number__iexact=parsed).order_by('-id').first()
        if not job:
            messages.error(request, f'No planning job found for code: {parsed}')
            return redirect('planning:scan')

        return redirect('planning:job_detail', job_id=job.id)

    return render(request, 'planning/planning_scan.html')


@login_required
@permission_required('can_edit_jobcard')
def planning_scan_open(request, jc_number):
    job = PlanningJob.objects.filter(jc_number__iexact=(jc_number or '').strip()).order_by('-id').first()
    if not job:
        messages.error(request, f'No planning job found for code: {jc_number}')
        return redirect('planning:scan')
    return redirect('planning:job_detail', job_id=job.id)


@login_required
def po_debug_extract(request):
    """Debug view: upload PDF and see raw text + table rows + per-strategy parse results."""
    import json as _json
    context = {}
    if request.method == 'POST':
        pdf_file = request.FILES.get('po_pdf')
        if pdf_file:
            try:
                import pdfplumber
                full_text = ''
                table_blobs = []
                table_rows = []
                pdf_file.seek(0)
                with pdfplumber.open(pdf_file) as pdf:
                    for page in pdf.pages:
                        page_text = page.extract_text(x_tolerance=3, y_tolerance=3) or ''
                        full_text += page_text + '\n'
                        for table in (page.extract_tables() or []):
                            for row in table or []:
                                parts = [str(col).strip() for col in (row or []) if str(col).strip()]
                                if parts:
                                    table_blobs.append(' '.join(parts))
                                    table_rows.append(parts)

                from .po_extractor import (
                    _build_sku_jobname_map,
                    _detect_expected_line_count,
                    _extract_items_flexible,
                    _extract_items_from_table_blobs,
                    _extract_items_from_table_rows,
                    _extract_items_from_text_windows,
                    _extract_items_strict,
                )
                sku_map = _build_sku_jobname_map(full_text, table_blobs)
                expected = _detect_expected_line_count(full_text, table_rows)
                strict = _extract_items_strict(full_text, sku_map)
                flexible = _extract_items_flexible(full_text, sku_map)
                from_rows = _extract_items_from_table_rows(table_rows, sku_map)
                from_blobs = _extract_items_from_table_blobs(table_blobs, sku_map)
                from_windows = _extract_items_from_text_windows(full_text, sku_map)

                context = {
                    'full_text': full_text,
                    'table_rows': _json.dumps(table_rows, indent=2),
                    'table_blobs': _json.dumps(table_blobs, indent=2),
                    'expected': expected,
                    'strict': _json.dumps(strict, indent=2),
                    'flexible': _json.dumps(flexible, indent=2),
                    'from_rows': _json.dumps(from_rows, indent=2),
                    'from_blobs': _json.dumps(from_blobs, indent=2),
                    'from_windows': _json.dumps(from_windows, indent=2),
                    'strict_count': len(strict),
                    'flexible_count': len(flexible),
                    'from_rows_count': len(from_rows),
                    'from_blobs_count': len(from_blobs),
                    'from_windows_count': len(from_windows),
                }
            except Exception as exc:
                context = {'error': str(exc)}
    return render(request, 'planning/po_debug.html', context)


@login_required
@permission_required('can_edit_jobcard')
def sku_recipes_list(request):
    """List all SKU recipes with search; handles delete via POST."""
    is_admin_user = _user_is_admin(request.user)
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        recipe_id = request.POST.get('recipe_id')
        redirect_url = request.path
        if request.GET:
            redirect_url += '?' + request.GET.urlencode()

        if action == 'delete':
            try:
                recipe_obj = SkuRecipe.objects.get(id=int(recipe_id))
                if recipe_obj.master_data_status == 'approved' and not is_admin_user:
                    messages.error(request, 'Approved records can only be deleted by admin users.')
                    return redirect(redirect_url)
                recipe_obj.delete()
                messages.success(request, 'SKU Recipe deleted.')
            except (TypeError, ValueError, SkuRecipe.DoesNotExist):
                messages.error(request, 'Invalid recipe ID.')
            return redirect(redirect_url)

        if action == 'archive':
            try:
                recipe_obj = SkuRecipe.objects.get(id=int(recipe_id))
                if recipe_obj.master_data_status == 'approved' and not is_admin_user:
                    messages.error(request, 'Approved records can only be archived by admin users.')
                    return redirect(redirect_url)
                recipe_obj.is_active = False
                recipe_obj.archived_by = request.user
                recipe_obj.archived_at = timezone.now()
                recipe_obj.archive_reason = (request.POST.get('archive_reason') or '').strip()
                recipe_obj.save(update_fields=['is_active', 'archived_by', 'archived_at', 'archive_reason', 'updated_at'])
                messages.success(request, 'SKU Recipe archived.')
            except (TypeError, ValueError, SkuRecipe.DoesNotExist):
                messages.error(request, 'Invalid recipe ID.')
            return redirect(redirect_url)

        if action in {'bulk_archive', 'bulk_delete'}:
            selected_ids = []
            for raw_id in request.POST.getlist('selected_ids'):
                try:
                    selected_ids.append(int(raw_id))
                except (TypeError, ValueError):
                    continue

            if not selected_ids:
                messages.error(request, 'Select at least one SKU recipe first.')
                return redirect(redirect_url)

            processed = 0
            skipped_locked = 0
            failures = []
            for recipe_obj in SkuRecipe.objects.filter(id__in=selected_ids, is_active=True):
                if recipe_obj.master_data_status == 'approved' and not is_admin_user:
                    skipped_locked += 1
                    continue

                if action == 'bulk_archive':
                    recipe_obj.is_active = False
                    recipe_obj.archived_by = request.user
                    recipe_obj.archived_at = timezone.now()
                    recipe_obj.archive_reason = ''
                    recipe_obj.save(update_fields=['is_active', 'archived_by', 'archived_at', 'archive_reason', 'updated_at'])
                    processed += 1
                else:
                    try:
                        recipe_obj.delete()
                        processed += 1
                    except Exception as exc:
                        failures.append(f'{recipe_obj.sku}: {str(exc)}')

            if action == 'bulk_archive':
                messages.success(request, f'Bulk archive complete. Archived {processed}, skipped {skipped_locked}.')
            else:
                messages.success(request, f'Bulk delete complete. Deleted {processed}, skipped {skipped_locked}.')
            if failures:
                messages.error(request, 'Some items could not be processed: ' + '; '.join(failures))
            return redirect(redirect_url)

        if action in {'submit_review', 'review', 'approve', 'back_to_draft'}:
            try:
                recipe = SkuRecipe.objects.get(id=int(recipe_id))
            except (TypeError, ValueError, SkuRecipe.DoesNotExist):
                messages.error(request, 'Invalid recipe ID.')
                return redirect(redirect_url)

            current_status = (recipe.master_data_status or 'draft').lower()

            if action == 'submit_review':
                if current_status != 'draft':
                    messages.error(request, f'SKU {recipe.sku} can only be submitted for review from Draft.')
                    return redirect(redirect_url)
                recipe.master_data_status = 'pending_review'
                recipe.reviewed_by = None
                recipe.reviewed_at = None
                recipe.approved_by = None
                recipe.approved_at = None
                recipe.save(update_fields=['master_data_status', 'reviewed_by', 'reviewed_at', 'approved_by', 'approved_at', 'updated_at'])
                messages.success(request, f'SKU {recipe.sku} submitted for review.')
                return redirect(redirect_url)

            if action == 'review':
                if current_status != 'pending_review':
                    messages.error(request, f'SKU {recipe.sku} can only move to Reviewed from Pending Review.')
                    return redirect(redirect_url)
                recipe.master_data_status = 'reviewed'
                recipe.reviewed_by = request.user
                recipe.reviewed_at = timezone.now()
                recipe.approved_by = None
                recipe.approved_at = None
                recipe.save(update_fields=['master_data_status', 'reviewed_by', 'reviewed_at', 'approved_by', 'approved_at', 'updated_at'])
                messages.success(request, f'SKU {recipe.sku} moved to Reviewed.')
                return redirect(redirect_url)

            if action == 'approve':
                if current_status != 'reviewed':
                    messages.error(request, f'SKU {recipe.sku} can only be Approved from Reviewed status.')
                    return redirect(redirect_url)
                missing_required = _missing_required_master_fields(recipe, recipe.job_name, allow_missing_plate_set_no=True)
                if missing_required:
                    messages.error(
                        request,
                        f'SKU {recipe.sku} cannot be approved. Missing required master data: {", ".join(missing_required)}.',
                    )
                    return redirect(redirect_url)
                recipe.master_data_status = 'approved'
                recipe.approved_by = request.user
                recipe.approved_at = timezone.now()
                recipe.save(update_fields=['master_data_status', 'approved_by', 'approved_at', 'updated_at'])
                approval_warnings = _warning_master_fields(recipe, recipe.job_name)
                if approval_warnings:
                    messages.warning(
                        request,
                        f'SKU {recipe.sku} approved. Notice: {", ".join(approval_warnings)} not set yet '
                        '(usually assigned at plate making).',
                    )
                else:
                    messages.success(request, f'SKU {recipe.sku} approved.')
                return redirect(redirect_url)

            if action == 'back_to_draft':
                if current_status == 'draft':
                    messages.info(request, f'SKU {recipe.sku} is already in Draft.')
                    return redirect(redirect_url)
                if current_status == 'approved' and not is_admin_user:
                    messages.error(request, 'Approved records can only be reverted by admin users.')
                    return redirect(redirect_url)
                comment = (request.POST.get('rejection_comment') or '').strip()
                if not comment:
                    messages.error(request, 'Please provide a reason when sending a record back to Draft.')
                    return redirect(redirect_url)
                recipe.master_data_status = 'draft'
                recipe.reviewed_by = None
                recipe.reviewed_at = None
                recipe.approved_by = None
                recipe.approved_at = None
                recipe.rejection_comment = comment
                recipe.last_rejected_by = request.user
                recipe.last_rejected_at = timezone.now()
                recipe.save(update_fields=[
                    'master_data_status', 'reviewed_by', 'reviewed_at',
                    'approved_by', 'approved_at', 'rejection_comment',
                    'last_rejected_by', 'last_rejected_at', 'updated_at',
                ])
                hydrate_sku_recipe_from_planning_jobs(recipe)
                if comment:
                    messages.warning(request, f'SKU {recipe.sku} sent back to Draft. Reason: {comment}')
                else:
                    messages.success(request, f'SKU {recipe.sku} moved back to Draft.')
                return redirect(redirect_url)

    q = (request.GET.get('q') or '').strip()
    status_filter = (request.GET.get('status') or '').strip()
    qs = SkuRecipe.objects.filter(is_active=True)
    if q:
        qs = qs.filter(
            Q(sku__icontains=q)
            | Q(job_name__icontains=q)
            | Q(material__icontains=q)
        )
    if status_filter in ('draft', 'pending_review', 'reviewed', 'approved'):
        qs = qs.filter(master_data_status=status_filter)
    paginator = Paginator(qs, 50)
    recipes = paginator.get_page(request.GET.get('page'))

    draft_count = SkuRecipe.objects.filter(is_active=True, master_data_status='draft').count()
    pending_review_count = SkuRecipe.objects.filter(is_active=True, master_data_status='pending_review').count()
    reviewed_count = SkuRecipe.objects.filter(is_active=True, master_data_status='reviewed').count()
    approved_count = SkuRecipe.objects.filter(is_active=True, master_data_status='approved').count()

    bulk_highlights = request.session.pop('sku_recipe_bulk_highlights', {})
    for recipe in recipes:
        meta = bulk_highlights.get(str(recipe.id), {})
        recipe.bulk_highlight_type = meta.get('type', '')
        recipe.bulk_highlight_fields = meta.get('fields', [])

    return render(request, 'planning/sku_recipes.html', {
        'recipes': recipes,
        'q': q,
        'status_filter': status_filter,
        'draft_count': draft_count,
        'pending_review_count': pending_review_count,
        'reviewed_count': reviewed_count,
        'approved_count': approved_count,
        'can_edit_approved': True,
        'can_admin_actions': is_admin_user,
    })


@login_required
def sku_recipes_status(request, status=None):
    """List SKU recipes filtered by a fixed status for role-specific views."""
    if status not in {'draft', 'pending_review', 'reviewed', 'approved'}:
        raise Http404('Unknown SKU recipe status view.')
    request.GET = request.GET.copy()
    request.GET['status'] = status
    return sku_recipes_list(request)


@login_required
@permission_required('can_edit_jobcard')
def sku_recipes_archived(request):
    """List archived SKU recipes."""
    is_admin_user = _user_is_admin(request.user)
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        redirect_url = request.path
        if request.GET:
            redirect_url += '?' + request.GET.urlencode()

        if action == 'restore':
            recipe_id = request.POST.get('recipe_id')
            try:
                recipe_obj = SkuRecipe.objects.get(id=int(recipe_id), is_active=False)
                if recipe_obj.master_data_status == 'approved' and not is_admin_user:
                    messages.error(request, 'Approved recipes can only be restored by admin users.')
                    return redirect(redirect_url)
                recipe_obj.is_active = True
                recipe_obj.save(update_fields=['is_active', 'updated_at'])
                messages.success(request, 'SKU Recipe restored to active master list.')
            except (TypeError, ValueError, SkuRecipe.DoesNotExist):
                messages.error(request, 'Invalid recipe ID.')
            return redirect(redirect_url)

        if action == 'delete':
            if not is_admin_user:
                messages.error(request, 'Only admin users can permanently delete archived recipes.')
                return redirect(redirect_url)

            recipe_id = request.POST.get('recipe_id')
            try:
                recipe_obj = SkuRecipe.objects.get(id=int(recipe_id), is_active=False)
                recipe_obj.delete()
                messages.success(request, 'Archived SKU Recipe permanently deleted.')
            except (TypeError, ValueError, SkuRecipe.DoesNotExist):
                messages.error(request, 'Invalid recipe ID.')
            return redirect(redirect_url)

        if action == 'bulk_restore':
            selected_ids = []
            for raw_id in request.POST.getlist('selected_ids'):
                try:
                    selected_ids.append(int(raw_id))
                except (TypeError, ValueError):
                    continue

            if not selected_ids:
                messages.error(request, 'Select at least one archived SKU recipe first.')
                return redirect(redirect_url)

            restored = 0
            skipped_locked = 0
            for recipe_obj in SkuRecipe.objects.filter(id__in=selected_ids, is_active=False):
                if recipe_obj.master_data_status == 'approved' and not is_admin_user:
                    skipped_locked += 1
                    continue
                recipe_obj.is_active = True
                recipe_obj.save(update_fields=['is_active', 'updated_at'])
                restored += 1

            messages.success(request, f'Bulk restore complete. Restored {restored}, skipped {skipped_locked}.')
            return redirect(redirect_url)

        if action == 'bulk_delete':
            if not is_admin_user:
                messages.error(request, 'Only admin users can permanently delete archived recipes.')
                return redirect(redirect_url)

            selected_ids = []
            for raw_id in request.POST.getlist('selected_ids'):
                try:
                    selected_ids.append(int(raw_id))
                except (TypeError, ValueError):
                    continue

            if not selected_ids:
                messages.error(request, 'Select at least one archived SKU recipe first.')
                return redirect(redirect_url)

            deleted = 0
            failures = []
            for recipe_obj in SkuRecipe.objects.filter(id__in=selected_ids, is_active=False):
                try:
                    recipe_obj.delete()
                    deleted += 1
                except Exception as exc:
                    failures.append(f'{recipe_obj.sku}: {str(exc)}')

            messages.success(request, f'Bulk delete complete. Deleted {deleted}.')
            if failures:
                messages.error(request, 'Some items could not be deleted: ' + '; '.join(failures))
            return redirect(redirect_url)

    q = (request.GET.get('q') or '').strip()
    status_filter = (request.GET.get('status') or '').strip()
    qs = SkuRecipe.objects.filter(is_active=False)
    if q:
        qs = qs.filter(
            Q(sku__icontains=q)
            | Q(job_name__icontains=q)
            | Q(material__icontains=q)
        )
    if status_filter in ('draft', 'pending_review', 'reviewed', 'approved'):
        qs = qs.filter(master_data_status=status_filter)
    paginator = Paginator(qs, 50)
    recipes = paginator.get_page(request.GET.get('page'))
    return render(request, 'planning/sku_recipes_archived.html', {
        'recipes': recipes,
        'q': q,
        'status_filter': status_filter,
        'can_restore_approved': is_admin_user,
        'can_delete_archived': is_admin_user,
    })


@login_required
@permission_required('can_edit_jobcard')
def sku_recipe_edit(request, recipe_id=None):
    """Create or edit a single SKU recipe."""
    if recipe_id:
        recipe = get_object_or_404(SkuRecipe, id=recipe_id, is_active=True)
        page_title = f'Edit SKU Recipe — {recipe.sku}'
    else:
        recipe = None
        page_title = 'Add New SKU Recipe'

    is_admin_user = _user_is_admin(request.user)
    can_edit_approved = not (recipe and recipe.master_data_status == 'approved')
    can_admin_actions = is_admin_user

    def _sku_recipe_edit_context(form, recipe_obj=None):
        current_recipe = recipe_obj if recipe_obj is not None else recipe
        context = {
            'form': form,
            'recipe': current_recipe,
            'page_title': page_title,
            'can_edit_approved': not (current_recipe and current_recipe.master_data_status == 'approved'),
            'can_admin_actions': can_admin_actions,
            'missing_required_fields': _missing_required_master_fields(
                current_recipe, allow_missing_plate_set_no=True,
            ) if current_recipe else [],
            'warning_master_fields': _warning_master_fields(current_recipe) if current_recipe else [],
            'is_readonly': False,
            'planner_can_edit_layout': planner_can_edit_designer_fields(current_recipe),
        }
        context.update(get_sku_recipe_form_ui_context(request.user, is_readonly=False))
        return context

    def _apply_form_permissions(form_obj, recipe_obj=None):
        apply_sku_recipe_form_role_permissions(
            form_obj,
            request.user,
            recipe=recipe_obj if recipe_obj is not None else recipe,
        )

    if request.method == 'POST':
        action = request.POST.get('action', '').strip()
        _do_sync_on_approve = False
        _notify_sku_event = None
        _hydrate_after_save = False
        if recipe and action == 'delete':
            if recipe.master_data_status == 'approved' and not is_admin_user:
                messages.error(request, 'Approved records can only be deleted by admin users.')
            else:
                recipe.delete()
                messages.success(request, f'SKU Recipe "{recipe.sku}" deleted.')
            return redirect('planning:sku_recipes')

        if recipe and action == 'archive':
            if recipe.master_data_status == 'approved' and not is_admin_user:
                messages.error(request, 'Approved records can only be archived by admin users.')
                return redirect('planning:sku_recipes')
            recipe.is_active = False
            recipe.archived_by = request.user
            recipe.archived_at = timezone.now()
            recipe.archive_reason = (request.POST.get('archive_reason') or '').strip()
            recipe.save(update_fields=['is_active', 'archived_by', 'archived_at', 'archive_reason', 'updated_at'])
            messages.success(request, f'SKU Recipe "{recipe.sku}" archived.')
            return redirect('planning:sku_recipes')

        if recipe and action == 'reopen_sku' and recipe.master_data_status == 'approved':
            # Reopen unlocks the SKU without validating master fields.
            # Only the reopen reason is required (audit trail).
            comment = (request.POST.get('rejection_comment') or '').strip()
            if not comment:
                messages.error(request, 'Enter a reopen reason, then click Reopen SKU.')
                locked_form = SkuRecipeForm(instance=recipe)
                _apply_form_permissions(locked_form)
                return render(request, 'planning/sku_recipe_edit.html', _sku_recipe_edit_context(locked_form))

            recipe.master_data_status = 'draft'
            recipe.reviewed_by = None
            recipe.reviewed_at = None
            recipe.approved_by = None
            recipe.approved_at = None
            recipe.rejection_comment = comment
            recipe.last_rejected_by = request.user
            recipe.last_rejected_at = timezone.now()
            recipe.save(update_fields=[
                'master_data_status',
                'reviewed_by',
                'reviewed_at',
                'approved_by',
                'approved_at',
                'rejection_comment',
                'last_rejected_by',
                'last_rejected_at',
                'updated_at',
            ])
            # Pull layout/planner blanks from planning jobs so reopen does not look empty.
            if hydrate_sku_recipe_from_planning_jobs(recipe):
                recipe.refresh_from_db()
            messages.success(
                request,
                f'SKU Recipe "{recipe.sku}" reopened to Draft. Known job layout data was restored where blank. '
                f'Update fields, then Submit for Review when complete.',
            )
            return redirect('planning:sku_recipe_edit', recipe_id=recipe.id)

        if recipe and recipe.master_data_status == 'approved' and action != 'reopen_sku':
            messages.error(request, 'Approved SKU is locked. Use Reopen SKU before making edits.')
            locked_form = SkuRecipeForm(instance=recipe)
            _apply_form_permissions(locked_form)
            return render(request, 'planning/sku_recipe_edit.html', _sku_recipe_edit_context(locked_form))

        previous_plate_values = {
            field: getattr(recipe, field, None) for field in PLATE_REMAKE_IMPACT_FIELDS
        } if recipe else {}

        posted = request.POST.copy()
        posted = merge_preserved_sku_recipe_fields(posted, recipe, request.user)
        form = SkuRecipeForm(posted, instance=recipe)
        if form.is_valid():
            obj = form.save(commit=False)
            obj = restore_locked_designer_fields_on_recipe(obj, recipe, request.user)
            if not recipe_id:
                obj.created_by = request.user

            current_status = (recipe.master_data_status if recipe else 'draft')
            obj.master_data_status = current_status
            obj.reviewed_by = recipe.reviewed_by if recipe else None
            obj.reviewed_at = recipe.reviewed_at if recipe else None
            obj.approved_by = recipe.approved_by if recipe else None
            obj.approved_at = recipe.approved_at if recipe else None

            if recipe_id and action:
                if action == 'submit_review' and current_status == 'draft':
                    missing = _missing_required_master_fields(obj, allow_missing_plate_set_no=True)
                    if missing:
                        messages.error(request, f'Cannot submit for review. Missing required fields: {", ".join(missing)}.')
                        _apply_form_permissions(form, obj)
                        return render(request, 'planning/sku_recipe_edit.html', _sku_recipe_edit_context(form, obj))
                    obj.master_data_status = 'pending_review'
                    obj.reviewed_by = None
                    obj.reviewed_at = None
                    obj.approved_by = None
                    obj.approved_at = None
                    messages.success(request, f'SKU Recipe "{obj.sku}" submitted for review. Status: Pending Review.')
                    _notify_sku_event = 'pending_review'
                elif action == 'review' and current_status == 'pending_review':
                    missing = _missing_required_master_fields(obj, allow_missing_plate_set_no=True)
                    if missing:
                        messages.error(request, f'Cannot submit for approval. Missing required fields: {", ".join(missing)}.')
                        _apply_form_permissions(form, obj)
                        return render(request, 'planning/sku_recipe_edit.html', _sku_recipe_edit_context(form, obj))
                    obj.master_data_status = 'reviewed'
                    obj.reviewed_by = request.user
                    obj.reviewed_at = timezone.now()
                    obj.approved_by = None
                    obj.approved_at = None
                    review_warnings = _warning_master_fields(obj)
                    if review_warnings:
                        messages.warning(
                            request,
                            f'SKU Recipe "{obj.sku}" reviewed. Notice: {", ".join(review_warnings)} not set yet '
                            '(usually assigned at plate making).',
                        )
                    else:
                        messages.success(request, f'SKU Recipe "{obj.sku}" reviewed and submitted for approval.')
                elif action == 'approve' and current_status == 'reviewed':
                    missing = _missing_required_master_fields(obj, allow_missing_plate_set_no=True)
                    if missing:
                        messages.error(request, f'Cannot approve. Missing required fields: {", ".join(missing)}.')
                        _apply_form_permissions(form, obj)
                        return render(request, 'planning/sku_recipe_edit.html', _sku_recipe_edit_context(form, obj))
                    obj.master_data_status = 'approved'
                    obj.approved_by = request.user
                    obj.approved_at = timezone.now()
                    # Will sync to planning after save
                    _do_sync_on_approve = True
                    approval_warnings = _warning_master_fields(obj)
                    if approval_warnings:
                        messages.warning(
                            request,
                            f'SKU Recipe "{obj.sku}" approved. Notice: {", ".join(approval_warnings)} not set yet '
                            '(usually assigned at plate making).',
                        )
                    else:
                        messages.success(request, f'SKU Recipe "{obj.sku}" approved for master data usage.')
                elif action == 'back_to_draft' and current_status in ('pending_review', 'reviewed'):
                    comment = (request.POST.get('rejection_comment') or '').strip()
                    if not comment:
                        messages.error(request, 'Please provide a reason when sending a record back to Draft.')
                        _apply_form_permissions(form, obj)
                        return render(request, 'planning/sku_recipe_edit.html', _sku_recipe_edit_context(form, obj))
                    obj.master_data_status = 'draft'
                    obj.reviewed_by = None
                    obj.reviewed_at = None
                    obj.approved_by = None
                    obj.approved_at = None
                    obj.rejection_comment = comment
                    obj.last_rejected_by = request.user
                    obj.last_rejected_at = timezone.now()
                    messages.success(request, f'SKU Recipe "{obj.sku}" moved back to Draft.')
                    _notify_sku_event = 'sent_back'
                    _hydrate_after_save = True
                else:
                    messages.info(request, f'SKU Recipe "{obj.sku}" saved without changing workflow status.')
            else:
                if recipe_id:
                    messages.success(request, f'SKU Recipe "{obj.sku}" saved.')
                else:
                    obj.master_data_status = 'draft'
                    obj.reviewed_by = None
                    obj.reviewed_at = None
                    obj.approved_by = None
                    obj.approved_at = None
                    messages.success(request, f'SKU Recipe "{obj.sku}" saved as Draft. Submit for approval from SKU Recipe Master.')

            obj.save()
            if _hydrate_after_save:
                hydrate_sku_recipe_from_planning_jobs(obj)
                obj.refresh_from_db()
            if _notify_sku_event == 'pending_review':
                try:
                    from core.notifications import notify_sku_pending_review
                    notify_sku_pending_review(obj, actor=request.user)
                except Exception:
                    logger.exception('SKU pending-review notification failed for %s', obj.sku)
            elif _notify_sku_event == 'sent_back':
                try:
                    from core.notifications import notify_sku_sent_back
                    notify_sku_sent_back(obj, actor=request.user)
                except Exception:
                    logger.exception('SKU sent-back notification failed for %s', obj.sku)
            plate_warning = get_plate_remake_warning_for_recipe_save(obj, previous_plate_values)
            if plate_warning and action != 'reopen_sku':
                messages.warning(request, plate_warning)
            if _do_sync_on_approve:
                try:
                    sync_result = _sync_new_jobs_for_approved_sku(obj.sku, actor=request.user)
                    messages.success(
                        request,
                        f'Planning jobs refreshed for approved SKU: updated {sync_result["updated"]}, locked {sync_result["locked"]}, missing draft jobs {sync_result.get("missing_jobs", 0)}.',
                    )
                except Exception:
                    logger.exception('Approved SKU sync to planning failed for %s', obj.sku)
                    messages.error(request, 'Error while sending approved SKU to Planning; check logs.')
            return redirect('planning:sku_recipes')
        else:
            # Surface a clear top-level message so users notice validation errors
            messages.error(request, 'There are errors in the form. Please correct the highlighted fields and try again.')
    else:
        if recipe and (recipe.master_data_status or '') != 'approved':
            if hydrate_sku_recipe_from_planning_jobs(recipe):
                recipe.refresh_from_db()
        form = SkuRecipeForm(
            instance=recipe,
            initial=build_sku_recipe_initial_from_recipe(recipe) if recipe else None,
        )

    _apply_form_permissions(form)

    return render(request, 'planning/sku_recipe_edit.html', _sku_recipe_edit_context(form))


@login_required
@permission_required('can_edit_jobcard')
def sku_recipe_bulk_upload(request):
    """Bulk upload SKU recipes from Google Sheet CSV/XLSX.

    Existing recipes: fill blank fields only (never wipe, never demote approved).
    New SKUs: created as draft.
    """
    if request.method == 'POST':
        upload_file = request.FILES.get('upload_file')
        if not upload_file:
            messages.error(request, 'Please choose a CSV, XLSX, or XLSB file to upload.')
            return redirect('planning:sku_recipe_bulk_upload')

        from planning.sku_sheet_import import (
            apply_sheet_values_to_recipe,
            parse_sheet_rows,
            row_to_field_values,
            _sheet_row_get,
        )

        try:
            rows = parse_sheet_rows(upload_file)
        except Exception as exc:
            messages.error(request, f'Could not read upload file: {exc}')
            return redirect('planning:sku_recipe_bulk_upload')

        if not rows:
            messages.error(request, 'No rows found in upload file.')
            return redirect('planning:sku_recipe_bulk_upload')

        created = 0
        updated = 0
        unchanged = 0
        bulk_highlights = {}
        highlight_fields = {
            'sku', 'job_name', 'material', 'color_spec', 'application', 'product_type',
            'job_process_type', 'print_passes', 'machine_name', 'plate_set_no',
            'size_w_mm', 'size_h_mm', 'ups', 'print_sheet_size',
            'purchase_sheet_size', 'purchase_sheet_ups',
            'default_unit_cost', 'daily_demand',
            'awc_no', 'die_cutting', 'notes',
        }

        for source in rows:
            sku = str(_sheet_row_get(source, 'SKU') or '').strip()
            if not sku:
                continue
            values = row_to_field_values(source)
            job_name = values.get('job_name') or str(_sheet_row_get(source, 'JOB NAME') or '').strip()
            if job_name:
                values['job_name'] = job_name

            existing = SkuRecipe.objects.filter(sku__iexact=sku).first()
            if existing:
                changed = apply_sheet_values_to_recipe(existing, values, fill_blanks_only=True)
                if not existing.legacy_produced:
                    existing.legacy_produced = True
                    changed = list(changed) + ['legacy_produced']
                if not changed:
                    unchanged += 1
                    continue
                existing.save(update_fields=list(dict.fromkeys(changed + ['updated_at'])))
                bulk_highlights[str(existing.id)] = {
                    'type': 'updated',
                    'fields': [f for f in changed if f in highlight_fields] or ['sku'],
                }
                updated += 1
                continue

            recipe = SkuRecipe(
                sku=sku,
                created_by=request.user,
                master_data_status='draft',
                legacy_produced=True,
            )
            if job_name:
                recipe.job_name = job_name
            changed = apply_sheet_values_to_recipe(recipe, values, fill_blanks_only=False)
            if 'f+b' in str(_sheet_row_get(source, 'Application') or '').lower():
                recipe.lamination_front_and_back = True
            recipe.save()
            bulk_highlights[str(recipe.id)] = {
                'type': 'created',
                'fields': [f for f in changed if f in highlight_fields] or ['sku'],
            }
            created += 1

        if bulk_highlights:
            request.session['sku_recipe_bulk_highlights'] = bulk_highlights

        if created or updated:
            messages.success(
                request,
                f'Bulk upload complete. Created {created}, filled blanks on {updated}, '
                f'unchanged {unchanged}. Existing recipes are never wiped or demoted.',
            )
        else:
            messages.info(
                request,
                f'No changes. Unchanged {unchanged}. Sheet cells only fill blank master fields.',
            )

        return redirect('planning:sku_recipes')

    return render(request, 'planning/sku_recipe_bulk_upload.html')


@login_required
@permission_required('can_view_jobcard')
def sku_recipe_template_download(request):
    """Return a CSV template for bulk SKU recipe upload."""
    headers = [
        'Sno.', 'SKU', 'JOB NAME', 'Order Status', 'Material', 'Color', 'Application', 'Product Type',
        'Size W mm', 'Size H mm', 'Size W Inch', 'Size H Inch', 'Ups', 'Print Sheet Size',
        'Purchase Sheet Size', 'Purchase Sheet ups', 'Remarks', 'Default Unit Cost',
        'Machine', 'Purchase Material', 'Daily Demand', 'AWC No', 'Plate Set No', 'Die', 'Notes',
    ]
    sample_row = [
        '1', 'SKU-001', 'Sample Job Name', 'Repeat', 'Art Card 300gsm', '4', 'UV', 'Labels',
        '100', '150', '3.94', '5.91', '4', '720x1020', '720x1020', '2', 'Sample remarks',
        '5.00', 'SM 74', 'Imported', '500', 'AWC-001', '1499', 'YES', 'Sample notes',
    ]
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(headers)
    writer.writerow(sample_row)
    response = HttpResponse(output.getvalue(), content_type='text/csv; charset=utf-8')
    response['Content-Disposition'] = 'attachment; filename="sku_recipe_upload_template.csv"'
    return response


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def pending_skus(request):
    """Central queue of SKUs that are still missing in SKU Recipe master data."""
    is_admin_user = _user_is_admin(request.user)
    if request.method == 'POST':
        action = (request.POST.get('action') or 'save').strip()
        sku = (request.POST.get('sku') or '').strip()
        po_doc_id = request.POST.get('po_doc_id')
        return_po = (request.POST.get('return_po') or '').strip()
        return_q = (request.POST.get('return_q') or '').strip()

        def _redirect_pending():
            params = {}
            if return_po:
                params['po'] = return_po
            if return_q:
                params['q'] = return_q
            url = reverse('qc:pending_skus')
            return redirect(f'{url}?{urlencode(params)}' if params else url)

        if action in {'delete', 'archive'}:
            recipe_id = request.POST.get('recipe_id')
            try:
                recipe_obj = SkuRecipe.objects.get(id=int(recipe_id))
            except (TypeError, ValueError, SkuRecipe.DoesNotExist):
                messages.error(request, 'Invalid recipe ID.')
                return _redirect_pending()

            if recipe_obj.master_data_status == 'approved' and not is_admin_user:
                messages.error(request, 'Approved records can only be managed by admin users.')
                return _redirect_pending()

            if action == 'delete':
                recipe_obj.delete()
                messages.success(request, f'SKU recipe {recipe_obj.sku} deleted.')
            else:
                recipe_obj.is_active = False
                recipe_obj.archived_by = request.user
                recipe_obj.archived_at = timezone.now()
                recipe_obj.archive_reason = (request.POST.get('archive_reason') or '').strip()
                recipe_obj.save(update_fields=['is_active', 'archived_by', 'archived_at', 'archive_reason', 'updated_at'])
                messages.success(request, f'SKU recipe {recipe_obj.sku} archived.')

            return _redirect_pending()

        if action == 'ignore':
            try:
                po_doc = PoDocument.objects.get(id=int(po_doc_id)) if po_doc_id else None
            except (TypeError, ValueError, PoDocument.DoesNotExist):
                po_doc = None

            if not po_doc or not sku:
                messages.error(request, 'Invalid SKU or PO reference for ignore action.')
                return _redirect_pending()

            payload = po_doc.extracted_payload or {}
            ignored = { _sku_key(s) for s in (payload.get('new_skus_ignored') or []) if s }
            ignored.add(_sku_key(sku))
            payload['new_skus_ignored'] = sorted(ignored)
            po_doc.extracted_payload = payload
            po_doc.save(update_fields=['extracted_payload'])

            po_number = (payload.get('po_number') or '').strip()
            if po_number:
                PlanningJob.objects.filter(
                    po_number__iexact=po_number,
                    sku__iexact=sku,
                    status__iexact='draft',
                    is_active=True,
                ).update(is_active=False, updated_at=timezone.now())

            messages.success(request, f'SKU {sku} will be ignored and removed from pending processing.')
            return _redirect_pending()

        if not sku:
            messages.error(request, 'SKU is required.')
            return _redirect_pending()

        if action in {'submit_review', 'approve', 'back_to_draft'}:
            recipe = SkuRecipe.objects.filter(sku__iexact=sku).first()
            if not recipe:
                messages.error(request, f'No SKU recipe found for {sku}. Save recipe data first.')
                return _redirect_pending()

            current_status = (recipe.master_data_status or 'draft').lower()

            if action == 'submit_review':
                if current_status != 'draft':
                    messages.error(request, f'SKU {sku} can only move to Pending Review from Draft.')
                    return _redirect_pending()
                recipe.master_data_status = 'pending_review'
                recipe.reviewed_by = None
                recipe.reviewed_at = None
                recipe.approved_by = None
                recipe.approved_at = None
                recipe.save(update_fields=['master_data_status', 'reviewed_by', 'reviewed_at', 'approved_by', 'approved_at', 'updated_at'])
                messages.success(request, f'SKU {sku} moved to Pending Review.')
                return _redirect_pending()

            if action == 'approve':
                if current_status != 'reviewed':
                    messages.error(request, f'SKU {sku} can only be Approved from Reviewed status.')
                    return _redirect_pending()
                missing_required = _missing_required_master_fields(recipe, allow_missing_plate_set_no=True)
                if missing_required:
                    messages.error(
                        request,
                        f'SKU {sku} cannot be approved. Missing required master data: {", ".join(missing_required)}.',
                    )
                    return _redirect_pending()
                recipe.master_data_status = 'approved'
                recipe.approved_by = request.user
                recipe.approved_at = timezone.now()
                recipe.save(update_fields=['master_data_status', 'approved_by', 'approved_at', 'updated_at'])
                sync_result = _sync_new_jobs_for_approved_sku(sku, actor=request.user)
                approval_warnings = _warning_master_fields(recipe)
                notice = ''
                if approval_warnings:
                    notice = (
                        f' Notice: {", ".join(approval_warnings)} not set yet '
                        '(usually assigned at plate making).'
                    )
                messages.success(
                    request,
                    f'SKU {sku} approved for master data usage. Planning jobs refreshed: updated {sync_result["updated"]}, locked {sync_result["locked"]}, missing draft jobs {sync_result.get("missing_jobs", 0)}.{notice}',
                )
                return _redirect_pending()

            if action == 'back_to_draft':
                if current_status == 'draft':
                    messages.info(request, f'SKU {sku} is already in Draft.')
                    return _redirect_pending()
                recipe.master_data_status = 'draft'
                recipe.reviewed_by = None
                recipe.reviewed_at = None
                recipe.approved_by = None
                recipe.approved_at = None
                recipe.save(update_fields=['master_data_status', 'reviewed_by', 'reviewed_at', 'approved_by', 'approved_at', 'updated_at'])
                messages.success(request, f'SKU {sku} moved back to Draft.')
                return _redirect_pending()

        job_name = (request.POST.get('job_name') or '').strip()
        material = (request.POST.get('material') or '').strip()
        color_spec = (request.POST.get('color_spec') or '').strip()
        application = (request.POST.get('application') or '').strip()
        department = (request.POST.get('department') or '').strip()
        print_sheet_size = (request.POST.get('print_sheet_size') or '').strip()
        purchase_sheet_size = (request.POST.get('purchase_sheet_size') or '').strip()
        purchase_sheet_ups = _to_optional_decimal(request.POST.get('purchase_sheet_ups'))
        ups = _to_optional_decimal(request.POST.get('ups'))
        daily_demand = _to_optional_decimal(request.POST.get('daily_demand'))
        awc_no = (request.POST.get('awc_no') or '').strip()
        die_cutting = (request.POST.get('die_cutting') or '').strip()

        unit_cost_raw = (request.POST.get('default_unit_cost') or '').strip()
        unit_cost = None
        if unit_cost_raw:
            try:
                unit_cost = Decimal(unit_cost_raw)
            except InvalidOperation:
                unit_cost = None

        if not job_name and not material:
            messages.error(request, 'Please enter at least Job Name or Material before saving.')
            return _redirect_pending()

        SkuRecipe.objects.update_or_create(
            sku=sku,
            defaults={
                'job_name': job_name,
                'material': material,
                'color_spec': color_spec,
                'application': application,
                'print_sheet_size': print_sheet_size,
                'purchase_sheet_size': purchase_sheet_size,
                'purchase_sheet_ups': purchase_sheet_ups,
                'ups': ups,
                'default_unit_cost': unit_cost,
                'daily_demand': daily_demand,
                'awc_no': awc_no,
                'die_cutting': die_cutting,
                'created_by': request.user,
                'master_data_status': 'draft',
                'reviewed_by': None,
                'reviewed_at': None,
                'approved_by': None,
                'approved_at': None,
            },
        )

        if po_doc_id:
            try:
                po_doc = PoDocument.objects.filter(id=int(po_doc_id)).first()
            except (TypeError, ValueError):
                po_doc = None
            if po_doc:
                payload = po_doc.extracted_payload or {}
                configured = set(payload.get('new_skus_configured') or [])
                configured.add(sku)
                payload['new_skus_configured'] = sorted(configured)
                po_doc.extracted_payload = payload
                po_doc.save(update_fields=['extracted_payload'])

        messages.success(request, f'SKU recipe saved for {sku}.')
        return _redirect_pending()

    po_filter = (request.GET.get('po') or '').strip()
    q = (request.GET.get('q') or '').strip()

    po_docs = PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('-created_at')[:400]
    grouped_docs = {}
    for doc in po_docs:
        payload = doc.extracted_payload or {}
        po_number = (payload.get('po_number') or '').strip()
        po_key = _normalize_po_number(po_number) or f'__doc_{doc.id}'
        grouped_docs.setdefault(po_key, []).append(doc)

    all_pending_rows = []
    for po_key, docs in grouped_docs.items():
        rows = _collect_pending_sku_rows(docs)
        display_po_number = (docs[0].extracted_payload or {}).get('po_number') or '-'
        seen_skus = set()
        for row in rows:
            sku_key = _sku_key(row.get('sku'))
            if sku_key and sku_key in seen_skus:
                continue
            if sku_key:
                seen_skus.add(sku_key)
            row['po_number'] = display_po_number
            all_pending_rows.append(row)

    po_summary_map = {}
    for row in all_pending_rows:
        po_key = row.get('po_number') or '-'
        current = po_summary_map.get(po_key)
        if not current:
            po_summary_map[po_key] = {
                'po_number': po_key,
                'count': 1,
                'po_doc_id': row.get('po_doc_id'),
            }
        else:
            current['count'] += 1

    pending_rows = all_pending_rows
    if po_filter:
        po_filter_key = _normalize_po_number(po_filter)
        pending_rows = [
            row
            for row in pending_rows
            if _normalize_po_number(row.get('po_number')) == po_filter_key
        ]

    # Bulk-fetch JC numbers early so text search can also filter by JC No
    job_key_map = {}
    if pending_rows:
        po_numbers = {r['po_number'] for r in pending_rows if r.get('po_number')}
        jobs_qs = PlanningJob.objects.filter(po_number__in=po_numbers).values('po_number', 'sku', 'jc_number')
        for j in jobs_qs:
            key = (_normalize_po_number(j['po_number']), _sku_key(j['sku']))
            if key not in job_key_map and j.get('jc_number'):
                job_key_map[key] = j['jc_number']
    for row in pending_rows:
        row['jc_number'] = job_key_map.get(
            (_normalize_po_number(row.get('po_number', '')), _sku_key(row.get('sku', ''))), ''
        )

    if q:
        q_upper = q.upper()
        pending_rows = [
            row
            for row in pending_rows
            if q_upper in (row.get('sku') or '').upper()
            or q_upper in (row.get('po_number') or '').upper()
            or q_upper in (row.get('job_name') or '').upper()
            or q_upper in (row.get('jc_number') or '').upper()
        ]

    sku_values = sorted({row['sku'] for row in pending_rows if row.get('sku')})
    recipes_by_sku = {}
    if sku_values:
        recipe_query = Q()
        for sku in sku_values:
            recipe_query |= Q(sku__iexact=sku)
        recipes = SkuRecipe.objects.filter(recipe_query)
        recipes_by_sku = {recipe.sku.upper(): recipe for recipe in recipes}

    for row in pending_rows:
        recipe = recipes_by_sku.get(_sku_key(row.get('sku')))
        row['recipe'] = recipe
        row['recipe_status'] = recipe.master_data_status if recipe else 'missing'
        row['missing_required_fields'] = _missing_required_master_fields(
            recipe, row.get('job_name') or '', allow_missing_plate_set_no=True,
        )
        row['warning_master_fields'] = _warning_master_fields(recipe, row.get('job_name') or '')

    pending_rows.sort(key=lambda row: (row['po_number'], row['sku']))

    context = {
        'pending_rows': pending_rows,
        'pending_count': len(pending_rows),
        'po_summary': sorted(po_summary_map.values(), key=lambda x: x['po_number']),
        'po_filter': po_filter,
        'q': q,
        'can_admin_actions': is_admin_user,
    }
    return render(request, 'planning/pending_skus.html', context)


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def pending_skus_ignored(request):
    """Display pending SKUs that were marked ignored and no longer appear in the active pending queue."""
    is_admin_user = _user_is_admin(request.user)
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        sku = (request.POST.get('sku') or '').strip()
        po_doc_id = request.POST.get('po_doc_id')
        if action == 'unignore':
            try:
                po_doc = PoDocument.objects.get(id=int(po_doc_id)) if po_doc_id else None
            except (TypeError, ValueError, PoDocument.DoesNotExist):
                po_doc = None

            if not po_doc or not sku:
                messages.error(request, 'Invalid PO or SKU for unignore action.')
                return redirect('planning:pending_skus_ignored')

            payload = po_doc.extracted_payload or {}
            ignored = [s for s in (payload.get('new_skus_ignored') or []) if s]
            normalized = _sku_key(sku)
            kept = [s for s in ignored if _sku_key(s) != normalized]
            payload['new_skus_ignored'] = sorted(kept)
            po_doc.extracted_payload = payload
            po_doc.save(update_fields=['extracted_payload'])
            
            # Reactivate archived planning jobs for this SKU and PO
            po_number = payload.get('po_number') or ''
            reactivated = PlanningJob.objects.filter(
                po_number__iexact=po_number,
                sku__iexact=sku,
                is_active=False,
            ).update(is_active=True, updated_at=timezone.now())
            
            messages.success(request, f'SKU {sku} restored to the pending queue. Reactivated {reactivated} planning job(s).')
            return redirect('planning:pending_skus_ignored')

    po_filter = (request.GET.get('po') or '').strip()
    q = (request.GET.get('q') or '').strip()

    docs = PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('-created_at')[:400]
    rows = []
    for doc in docs:
        payload = doc.extracted_payload or {}
        ignored_skus = [s for s in (payload.get('new_skus_ignored') or []) if s]
        if not ignored_skus:
            continue

        po_number = payload.get('po_number') or '-'
        for sku in ignored_skus:
            if not sku:
                continue
            recipe = SkuRecipe.objects.filter(sku__iexact=sku).first()
            rows.append({
                'po_doc_id': doc.id,
                'po_number': po_number,
                'supplier': payload.get('supplier_name') or '-',
                'sku': sku,
                'job_name': recipe.job_name if recipe else '-',
                'recipe_status': recipe.master_data_status if recipe else 'missing',
                'recipe': recipe,
                'uploaded_at': doc.created_at,
            })

    if po_filter:
        po_filter_key = _normalize_po_number(po_filter)
        rows = [
            row
            for row in rows
            if _normalize_po_number(row['po_number']) == po_filter_key
        ]
    if q:
        q_upper = q.upper()
        rows = [
            row
            for row in rows
            if q_upper in (row['sku'] or '').upper()
            or q_upper in (row['po_number'] or '').upper()
            or q_upper in (row['job_name'] or '').upper()
        ]

    po_summary = {}
    for row in rows:
        po_summary.setdefault(row['po_number'], {'po_number': row['po_number'], 'count': 0})
        po_summary[row['po_number']]['count'] += 1

    rows.sort(key=lambda row: (row['po_number'], row['sku']))

    return render(request, 'planning/pending_skus_ignored.html', {
        'rows': rows,
        'po_summary': sorted(po_summary.values(), key=lambda x: x['po_number']),
        'po_filter': po_filter,
        'q': q,
        'can_admin_actions': is_admin_user,
    })


@login_required
@permission_required('can_view_sku_master_review_queue')
@transaction.atomic
def pending_sku_master_entry(request):
    """Open a focused form for one pending SKU and send it through master-data approval flow."""
    sku = (request.GET.get('sku') or request.POST.get('sku') or '').strip()
    po_doc_id_raw = request.GET.get('po_doc_id') or request.POST.get('po_doc_id')
    return_po = (request.GET.get('return_po') or request.POST.get('return_po') or '').strip()
    return_q = (request.GET.get('return_q') or request.POST.get('return_q') or '').strip()
    is_readonly = (request.GET.get('readonly') or request.POST.get('readonly') or '').strip() in {'1', 'true', 'yes'}

    profile = getattr(request.user, 'profile', None)
    if not is_readonly and not (profile and profile.can_edit_jobcard()):
        messages.error(request, 'You do not have permission to edit this SKU master entry.')
        params = {}
        if return_po:
            params['po'] = return_po
        if return_q:
            params['q'] = return_q
        target = reverse('qc:pending_skus')
        return redirect(f'{target}?{urlencode(params)}' if params else target)

    def _redirect_pending():
        params = {}
        if return_po:
            params['po'] = return_po
        if return_q:
            params['q'] = return_q
        url = reverse('qc:pending_skus')
        return redirect(f'{url}?{urlencode(params)}' if params else url)

    def _redirect_qc_review():
        params = {}
        if return_po:
            params['po'] = return_po
        if return_q:
            params['q'] = return_q
        url = reverse('qc:master_review')
        return redirect(f'{url}?{urlencode(params)}' if params else url)

    try:
        po_doc_id = int(po_doc_id_raw)
    except (TypeError, ValueError):
        po_doc_id = None

    if not sku or not po_doc_id:
        messages.error(request, 'Missing SKU or PO reference for master-data entry.')
        return _redirect_pending()

    po_doc = PoDocument.objects.filter(id=po_doc_id).first()
    if not po_doc:
        messages.error(request, 'PO document was not found.')
        return _redirect_pending()

    payload = po_doc.extracted_payload or {}
    po_number = payload.get('po_number') or '-'
    items = _sanitize_po_payload_items(payload)
    is_admin_user = _user_is_admin(request.user)
    user_can_approve_qc = bool(profile and profile.can_approve_sku_master_review())

    suggested_item = None
    sku_key = _sku_key(sku)
    for item in items:
        if _sku_key(item.get('sku')) == sku_key:
            suggested_item = item
            break
    suggested_item = suggested_item or {}
    po_job_name = (suggested_item.get('job_name') or '').strip() or sku
    po_department = (payload.get('department') or '').strip()
    po_unit_cost = _to_decimal(suggested_item.get('unit_cost'))
    po_color_spec = _normalize_color_spec_input(
        suggested_item.get('color_spec') or suggested_item.get('colour') or suggested_item.get('color') or ''
    )
    po_application = _normalize_application_input(suggested_item.get('application') or payload.get('application') or '')
    po_remarks = (suggested_item.get('remarks') or '').strip()

    # Fetch the best available recipe using priority: approved > reviewed > pending_review > draft
    _all_recipes = list(SkuRecipe.objects.filter(sku__iexact=sku))
    recipe = None
    if _all_recipes:
        recipe = min(_all_recipes, key=lambda r: SKU_RECIPE_STATUS_ORDER.get(r.master_data_status or '', 99))

    po_number_val = payload.get('po_number') or ''
    job_obj = PlanningJob.objects.filter(po_number=po_number_val, sku__iexact=sku).first()
    if job_obj and not recipe:
        recipe = ensure_sku_recipe_for_planning_job(job_obj, actor=request.user)
        sync_planning_job_fields_to_sku_recipe(job_obj, recipe)


    if request.method == 'POST':
        action = (request.POST.get('action') or 'save_draft').strip()

        # QC actions from readonly review view (approve / reject).
        if is_readonly and action in {'approve', 'back_to_draft', 'reject'}:
            if not user_can_approve_qc:
                messages.error(request, 'You do not have permission to review SKU master records.')
                return _redirect_qc_review()
            if not recipe:
                messages.error(request, f'SKU recipe for {sku} was not found.')
                return _redirect_qc_review()

            current_status = (recipe.master_data_status or 'draft').lower()
            if action == 'approve':
                if current_status not in {'pending_review', 'reviewed'}:
                    messages.error(request, f'SKU {sku} is not in QC review queue.')
                    return _redirect_qc_review()
                missing_required = _missing_required_master_fields(
                    recipe, recipe.job_name or po_job_name, allow_missing_plate_set_no=True,
                )
                if missing_required:
                    messages.error(
                        request,
                        f'SKU {sku} cannot be approved. Missing required fields: {", ".join(missing_required)}.',
                    )
                    return _redirect_qc_review()
                if current_status == 'pending_review':
                    recipe.reviewed_by = request.user
                    recipe.reviewed_at = timezone.now()
                recipe.master_data_status = 'approved'
                recipe.approved_by = request.user
                recipe.approved_at = timezone.now()
                recipe.rejection_comment = ''
                recipe.last_rejected_by = None
                recipe.last_rejected_at = None
                recipe.save(update_fields=[
                    'master_data_status', 'reviewed_by', 'reviewed_at',
                    'approved_by', 'approved_at', 'rejection_comment',
                    'last_rejected_by', 'last_rejected_at', 'updated_at',
                ])
                sync_result = _sync_new_jobs_for_approved_sku(sku, actor=request.user)
                approval_warnings = _warning_master_fields(recipe, recipe.job_name or po_job_name)
                notice = ''
                if approval_warnings:
                    notice = (
                        f' Notice: {", ".join(approval_warnings)} not set yet '
                        '(usually assigned at plate making).'
                    )
                messages.success(
                    request,
                    f'SKU {sku} approved. Planning jobs refreshed: updated {sync_result["updated"]}, '
                    f'locked {sync_result["locked"]}, missing draft jobs {sync_result.get("missing_jobs", 0)}.{notice}',
                )
                return _redirect_qc_review()

            # reject / back_to_draft
            rejection_comment = (request.POST.get('rejection_comment') or '').strip()
            if current_status == 'draft':
                messages.info(request, f'SKU {sku} is already in Draft.')
                return _redirect_qc_review()
            if not rejection_comment:
                messages.error(request, 'Reason is required to reject / send back this SKU.')
                return _redirect_qc_review()
            recipe.master_data_status = 'draft'
            recipe.reviewed_by = None
            recipe.reviewed_at = None
            recipe.approved_by = None
            recipe.approved_at = None
            recipe.rejection_comment = rejection_comment
            recipe.last_rejected_by = request.user
            recipe.last_rejected_at = timezone.now()
            recipe.save(update_fields=[
                'master_data_status', 'reviewed_by', 'reviewed_at',
                'approved_by', 'approved_at', 'rejection_comment',
                'last_rejected_by', 'last_rejected_at', 'updated_at',
            ])
            hydrate_sku_recipe_from_planning_jobs(recipe)
            try:
                from core.notifications import notify_sku_sent_back
                notify_sku_sent_back(recipe, actor=request.user)
            except Exception:
                logger.exception('SKU sent-back notification failed for %s', sku)
            messages.warning(request, f'SKU {sku} sent back to Draft. Reason: {rejection_comment}')
            return _redirect_qc_review()

        if is_readonly:
            messages.error(request, 'Read-only mode does not allow edits.')
            return _redirect_pending()

        if recipe and action == 'delete':
            if recipe.master_data_status == 'approved' and not is_admin_user:
                messages.error(request, 'Approved records can only be deleted by admin users.')
            else:
                recipe.delete()
                messages.success(request, f'SKU Recipe "{recipe.sku}" deleted.')
            return _redirect_pending()

        if recipe and action == 'archive':
            if recipe.master_data_status == 'approved' and not is_admin_user:
                messages.error(request, 'Approved records can only be archived by admin users.')
                return _redirect_pending()
            recipe.is_active = False
            recipe.archived_by = request.user
            recipe.archived_at = timezone.now()
            recipe.archive_reason = (request.POST.get('archive_reason') or '').strip()
            recipe.save(update_fields=['is_active', 'archived_by', 'archived_at', 'archive_reason', 'updated_at'])
            messages.success(request, f'SKU Recipe "{recipe.sku}" archived.')
            return _redirect_pending()

        posted = request.POST.copy()
        # Job name is sourced from PO parsing; keep it authoritative and non-editable.
        posted['job_name'] = po_job_name
        posted['sku'] = sku
        if not (posted.get('job_process_type') or '').strip():
            posted['job_process_type'] = (recipe.job_process_type if recipe else '') or 'print_and_pack'
        if not (posted.get('default_unit_cost') or '').strip() and po_unit_cost is not None:
            posted['default_unit_cost'] = str(po_unit_cost)
        if not (posted.get('color_spec') or '').strip() and po_color_spec:
            posted['color_spec'] = po_color_spec
        if not (posted.get('application') or '').strip() and po_application:
            posted['application'] = po_application
        posted = merge_preserved_sku_recipe_fields(posted, recipe, request.user)
        form = SkuRecipeForm(posted, instance=recipe)
        apply_sku_recipe_form_role_permissions(form, request.user, is_readonly=is_readonly, recipe=recipe)
        prepare_sku_recipe_form_for_master_entry(form, action=action)
        if form.is_valid():
            action = (request.POST.get('action') or 'save_draft').strip()
            obj = form.save(commit=False)
            obj = restore_locked_designer_fields_on_recipe(obj, recipe, request.user)
            obj.sku = sku
            obj.job_name = po_job_name
            if not recipe:
                obj.created_by = request.user

            if action == 'send_to_plate_making':
                # Find the draft PlanningJob first so we can block duplicate plate requests.
                po_number = payload.get('po_number') or ''
                job = (
                    PlanningJob.objects.filter(po_number=po_number, sku__iexact=sku)
                    .order_by('-updated_at', '-id')
                    .first()
                )
                if job and _normalize_status(job.status) != 'draft':
                    messages.error(
                        request,
                        f'Planning job {job.jc_number} for SKU {sku} is no longer in draft '
                        f'({job.get_status_display()}). Open the job detail page instead of sending from pending master entry.',
                    )
                    return redirect(
                        f"{reverse('planning:pending_sku_master_entry')}?po_doc_id={po_doc_id}&sku={sku}"
                        f"{'&return_q=' + return_q if return_q else ''}{'&return_po=' + return_po if return_po else ''}"
                    )

                existing_plate_req, block_reason = get_plate_request_block_for_master_entry(job)
                if existing_plate_req:
                    from django.urls import reverse as _reverse
                    plate_url = _reverse('printing_plates:request_detail', args=[existing_plate_req.pk])
                    if block_reason == 'open':
                        messages.error(
                            request,
                            f'Plate request already active for {job.jc_number if job else sku} '
                            f'(status: {existing_plate_req.get_status_display()}). '
                            f'Do not create another request. Open the existing request, or use '
                            f'Send to QC Review when master data is complete. '
                            f'Plate request: {plate_url}',
                        )
                    else:
                        messages.error(
                            request,
                            f'Plates were already issued for {job.jc_number if job else sku}. '
                            f'Use Released Jobs for plate replacement, or Send to QC Review when master data is complete. '
                            f'Plate request: {plate_url}',
                        )
                    return redirect(
                        f"{reverse('planning:pending_sku_master_entry')}?po_doc_id={po_doc_id}&sku={sku}"
                        f"{'&return_q=' + return_q if return_q else ''}{'&return_po=' + return_po if return_po else ''}"
                    )

                obj.master_data_status = 'draft'
                obj.reviewed_by = None
                obj.reviewed_at = None
                obj.approved_by = None
                obj.approved_at = None
                obj.save()

                configured = set(payload.get('new_skus_configured') or [])
                configured.add(sku)
                payload['new_skus_configured'] = sorted(configured)
                po_doc.extracted_payload = payload
                po_doc.save(update_fields=['extracted_payload'])

                if not job:
                    job = ensure_draft_planning_job_for_po_sku(
                        po_doc,
                        sku,
                        actor=request.user,
                        recipe=obj,
                    )
                    if not job:
                        messages.error(
                            request,
                            f'Could not create a draft planning job for SKU {sku} on '
                            f'{document_type_label(po_number)} {po_number}. '
                            f'Check that this SKU exists on the {document_type_label(po_number)} document.',
                        )
                        return redirect(
                            f"{reverse('planning:pending_sku_master_entry')}?po_doc_id={po_doc_id}&sku={sku}"
                            f"{'&return_q=' + return_q if return_q else ''}{'&return_po=' + return_po if return_po else ''}"
                        )

                if job:
                    if (job.repeat_flag or '').strip() != 'New' and not job.approved_sku_recipe:
                        messages.error(
                            request,
                            f'Repeat SKU {sku} on {job.jc_number} needs an approved master recipe before plate making. '
                            f'Complete and approve master data on the New PO for this SKU first.',
                        )
                        return redirect(
                            f"{reverse('planning:pending_sku_master_entry')}?po_doc_id={po_doc_id}&sku={sku}"
                            f"{'&return_q=' + return_q if return_q else ''}{'&return_po=' + return_po if return_po else ''}"
                        )

                    job.material = obj.material
                    job.application = obj.application
                    job.machine_name = obj.machine_name
                    job.plate_set_no = obj.plate_set_no
                    if (obj.job_process_type or 'print_and_pack') == 'cut_and_pack':
                        job.print_passes = None
                    elif obj.print_passes:
                        job.print_passes = int(obj.print_passes)
                    job.save(update_fields=[
                        'material', 'application', 'machine_name', 'plate_set_no',
                        'print_passes', 'updated_at',
                    ])

                    plate_errors = get_plate_making_prerequisite_errors(job)
                    if plate_errors:
                        messages.error(request, ' '.join(plate_errors))
                        return redirect(
                            f"{reverse('planning:pending_sku_master_entry')}?po_doc_id={po_doc_id}&sku={sku}"
                            f"{'&return_q=' + return_q if return_q else ''}{'&return_po=' + return_po if return_po else ''}"
                        )

                    from planning.sku_classification import plate_making_stage_for_repeat_flag

                    planning_stage = plate_making_stage_for_repeat_flag(job.repeat_flag)

                    job.planning_stage = planning_stage
                    job.planning_stage_changed_at = timezone.now()
                    job.planning_stage_changed_by = request.user
                    job.save(update_fields=['planning_stage', 'planning_stage_changed_at', 'planning_stage_changed_by', 'updated_at'])
                    
                    plate_req = trigger_plate_request_for_planning_job(job, request.user)
                    if plate_req:
                        from django.urls import reverse as _reverse
                        plate_url = _reverse('printing_plates:request_detail', args=[plate_req.pk])
                        messages.success(
                            request,
                            f'SKU {sku} saved, and Job Card {job.jc_number} sent to Plate Making. '
                            f'Open plate request: {plate_url}',
                        )
                    else:
                        from printing_plates.services import get_planning_plate_making_block_message

                        block_message = get_planning_plate_making_block_message(job)
                        if block_message:
                            messages.error(request, block_message)
                        else:
                            messages.warning(
                                request,
                                f'SKU {sku} saved and {job.jc_number} moved to plate making, '
                                f'but no plate request was created. Check Printing Plates or contact admin.',
                            )

                return _redirect_pending()
            elif action == 'submit_review':
                missing_required = _missing_required_master_fields(obj, allow_missing_plate_set_no=True)
                if missing_required:
                    messages.error(
                        request,
                        f'SKU {sku} cannot be sent for QC review. Missing required data: {", ".join(missing_required)}.',
                    )
                else:
                    obj.master_data_status = 'pending_review'
                    obj.reviewed_by = None
                    obj.reviewed_at = None
                    obj.approved_by = None
                    obj.approved_at = None
                    obj.save()

                    configured = set(payload.get('new_skus_configured') or [])
                    configured.add(sku)
                    payload['new_skus_configured'] = sorted(configured)
                    po_doc.extracted_payload = payload
                    po_doc.save(update_fields=['extracted_payload'])

                    messages.success(request, f'SKU {sku} submitted for QC review.')
                    return _redirect_pending()
            else:
                obj.master_data_status = 'draft'
                obj.reviewed_by = None
                obj.reviewed_at = None
                obj.approved_by = None
                obj.approved_at = None
                obj.save()

                configured = set(payload.get('new_skus_configured') or [])
                configured.add(sku)
                payload['new_skus_configured'] = sorted(configured)
                po_doc.extracted_payload = payload
                po_doc.save(update_fields=['extracted_payload'])

                messages.success(request, f'SKU {sku} saved as Draft.')
                return _redirect_pending()
    else:
        po_defaults = {
            'default_unit_cost': po_unit_cost,
            'color_spec': po_color_spec,
            'application': po_application,
            'remarks': po_remarks,
        }
        if job_obj:
            initial = build_sku_recipe_initial_from_planning_job(job_obj, recipe=recipe, po_defaults=po_defaults)
        else:
            initial = build_sku_recipe_initial_from_recipe(recipe, po_defaults=po_defaults)
            initial['sku'] = sku
            initial['job_name'] = po_job_name
            if not initial.get('remarks'):
                initial['remarks'] = po_remarks
        form = SkuRecipeForm(instance=recipe, initial=initial)
        apply_sku_recipe_form_role_permissions(form, request.user, is_readonly=is_readonly, recipe=recipe)

    form.fields['sku'].widget.attrs['readonly'] = True

    current_recipe = recipe
    if request.method == 'POST' and form.is_valid() and 'obj' in locals():
        current_recipe = obj

    mismatch_alerts = []
    if current_recipe:
        cost_alert = _build_cost_mismatch_note(current_recipe.default_unit_cost, po_unit_cost)
        if cost_alert:
            mismatch_alerts.append(cost_alert)

    from planning.sku_classification import classify_po_line, is_job_repeat_classification_locked

    if not job_obj:
        job_obj = PlanningJob.objects.filter(sku__iexact=sku).order_by('-updated_at').first()
    if job_obj and (is_job_repeat_classification_locked(job_obj) or (job_obj.repeat_flag or '').strip()):
        repeat_flag_display = (job_obj.repeat_flag or '').strip() or 'New'
    else:
        repeat_flag_display, _reason = classify_po_line(
            sku,
            po_number,
            po_doc_created_at=po_doc.created_at,
            po_doc_id=po_doc.id,
            recipe=current_recipe,
        )

    active_plate_request, plate_block_reason = get_plate_request_block_for_master_entry(job_obj)

    context = {
        'form': form,
        'recipe': current_recipe,
        'sku': sku,
        'po_doc_id': po_doc_id,
        'po_number': po_number,
        'return_po': return_po,
        'return_q': return_q,
        'suggested_job_name': po_job_name,
        'suggested_qty': _format_display_qty(suggested_item.get('quantity')),
        'suggested_delivery_date': suggested_item.get('delivery_date') or '-',
        'recipe_status': (current_recipe.master_data_status if current_recipe else 'missing'),
        'missing_required_fields': _missing_required_master_fields(
            current_recipe, po_job_name, allow_missing_plate_set_no=True,
        ),
        'warning_master_fields': _warning_master_fields(current_recipe, po_job_name),
        'mismatch_alerts': mismatch_alerts,
        'is_readonly': is_readonly,
        'is_new_job': repeat_flag_display == 'New',
        'repeat_flag_display': repeat_flag_display,
        'jc_number': (job_obj.jc_number if job_obj else '') or '',
        'user_can_approve_qc': user_can_approve_qc,
        'active_plate_request': active_plate_request,
        'plate_block_reason': plate_block_reason,
        'plate_request_already_active': bool(active_plate_request),
    }
    context['can_admin_actions'] = is_admin_user
    from planning.sku_duplicate_alert import build_sku_duplicate_alert_for_sku

    context['sku_duplicate_alert'] = build_sku_duplicate_alert_for_sku(sku, current_job=job_obj)
    context.update(get_sku_recipe_form_ui_context(request.user, is_readonly=is_readonly))
    return render(request, 'planning/pending_sku_master_entry.html', context)


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def po_inbox(request):
    """PO intake queue after upload: split-ready documents with repeat/new counts."""
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'delete_po_intake':
            if not _user_is_admin(request.user):
                messages.error(request, 'Only administrators can delete PO intake records.')
                return redirect('planning:po_inbox')

            po_doc_id = (request.POST.get('po_doc_id') or '').strip()
            po_number = (request.POST.get('po_number') or '').strip()
            doc = None

            if po_doc_id:
                try:
                    doc = PoDocument.objects.get(id=int(po_doc_id))
                except (TypeError, ValueError, PoDocument.DoesNotExist):
                    messages.error(request, 'Invalid PO document selected for delete action.')
                    return redirect('planning:po_inbox')

                payload = doc.extracted_payload or {}
                po_number = (payload.get('po_number') or '').strip() or po_number

            if po_number:
                po_number_key = po_number
                docs_to_delete = PoDocument.objects.filter(extracted_payload__po_number__iexact=po_number_key)
                deleted_docs = docs_to_delete.count()
                for po_doc in docs_to_delete:
                    try:
                        if po_doc.po_file:
                            po_doc.po_file.delete(save=False)
                    except Exception:
                        pass
                    po_doc.delete()

                job_cards_deleted = JobCard.objects.filter(
                    Q(planning_job__po_number__iexact=po_number_key) | Q(PO_No__iexact=po_number_key)
                ).delete()[0]

                plate_requests_deleted = PlateRequest.objects.filter(
                    planning_job__po_number__iexact=po_number_key
                ).delete()[0]

                planning_job_qs = PlanningJob.objects.filter(po_number__iexact=po_number_key)
                jobs_deleted = planning_job_qs.delete()[0]

                messages.success(
                    request,
                    f'Deleted {document_type_label(po_number_key)} {po_number_key} from ERP: {deleted_docs} intake document(s), {jobs_deleted} planning job(s), {job_cards_deleted} job card(s), {plate_requests_deleted} plate request(s).',
                )
                return redirect('planning:po_inbox')

            if doc:
                try:
                    if doc.po_file:
                        doc.po_file.delete(save=False)
                except Exception:
                    pass
                doc.delete()
                messages.success(request, f'Deleted PO intake document {doc.id}.')
                return redirect('planning:po_inbox')

            messages.error(request, 'Invalid PO number for delete action.')
            return redirect('planning:po_inbox')

    search_query = (request.GET.get('q') or '').strip()
    per_page = _to_optional_positive_int(request.GET.get('per_page')) or 20
    page_number = request.GET.get('page') or 1

    docs_qs = PoDocument.objects.exclude(extracted_payload__isnull=True).values('id', 'created_at', 'extracted_payload').order_by('-created_at')

    # A search hits the whole table via the DB (cheap — no Python payload parsing
    # happens until after this filter), so an older PO/WO is still found regardless
    # of the cap below. The per-document classification cost that used to make a
    # large cap here expensive was O(items x total_documents) (see
    # sku_classification.build_sku_doc_index); now that it's a single batched
    # index build per request, processing every document is cheap enough that the
    # cap only exists as a safety net against pathological growth, not to bound
    # normal browsing — so "Total PO Documents" and the paginated list agree.
    if search_query:
        docs_qs = docs_qs.filter(
            Q(extracted_payload__po_number__icontains=search_query)
            | Q(extracted_payload__supplier_name__icontains=search_query)
        )
    dedupe_cap = 5000

    deduped_docs = []
    seen_po_numbers = set()
    for doc in docs_qs.iterator(chunk_size=100):
        payload = doc.get('extracted_payload') or {}
        po_number = str(payload.get('po_number') or '').strip().upper()
        if po_number:
            if po_number in seen_po_numbers:
                continue
            seen_po_numbers.add(po_number)

        deduped_docs.append(doc)
        if len(deduped_docs) >= dedupe_cap:
            break

    page_number_int = _to_optional_positive_int(page_number) or 1
    docs_to_load = deduped_docs

    doc_items = []
    all_sku_keys = set()
    for doc in docs_to_load:
        payload = doc.get('extracted_payload') or {}
        items = _po_payload_items(payload)
        doc_items.append((doc, payload, items))
        all_sku_keys.update(
            _sku_key(item.get('sku'))
            for item in items
            if item.get('sku')
        )

    recipe_map_all = {}
    if all_sku_keys:
        recipe_map_all = _build_recipe_map([
            {'sku': sku}
            for sku in all_sku_keys
        ])

    # Built once per request and reused for every row below — classifying each
    # line item against a per-item full PoDocument table scan (the old behavior)
    # made this page scale as O(items x total PO documents); this index turns
    # each item's check into an O(1)-ish in-memory lookup instead.
    from planning.sku_classification import build_sku_doc_index
    sku_doc_index = build_sku_doc_index() if doc_items else {}

    rows = []
    for doc, payload, items in doc_items:
        item_sku_keys = {
            _sku_key(item.get('sku'))
            for item in items
            if item.get('sku')
        }
        recipe_map = {
            key: recipe_map_all[key]
            for key in item_sku_keys
            if key in recipe_map_all
        }
        po_number_val = payload.get('po_number') or ''
        _, repeat_count, new_count, missing_skus = _annotate_items_with_recipe(
            items,
            recipe_map=recipe_map,
            current_po_number=po_number_val,
            po_doc_created_at=doc['created_at'],
            po_doc_id=doc['id'],
            sku_doc_index=sku_doc_index,
        )

        uploaded_at = doc['created_at']
        if uploaded_at and getattr(uploaded_at, 'tzinfo', None) is not None:
            uploaded_at = uploaded_at.astimezone().replace(tzinfo=None)

        rows.append(
            {
                'po_doc_id': doc['id'],
                'uploaded': uploaded_at,
                'po_number': payload.get('po_number') or '-',
                'document_type': payload.get('document_type') or 'PO',
                'supplier': payload.get('supplier_name') or '-',
                'item_count': len(items),
                'repeat_count': repeat_count,
                'new_count': new_count,
                'missing_count': len(missing_skus),
                'ignored_count': len([s for s in (payload.get('new_skus_ignored') or []) if s]),
                'repeat_jobs_created_count': payload.get('repeat_jobs_created_count') or 0,
                'repeat_jobs_updated_count': payload.get('repeat_jobs_updated_count') or 0,
                'repeat_jobs_locked_count': payload.get('repeat_jobs_locked_count') or 0,
                'repeat_jobs_missing_recipe_count': payload.get('repeat_jobs_missing_recipe_count') or 0,
                'new_skus_sent_to_planning_count': len(payload.get('new_skus_sent_to_planning') or []),
                'merged_duplicates': bool(payload.get('merged_duplicate_skus')),
                'merged_duplicate_skus': payload.get('merged_duplicate_skus') or [],
                'source_item_count': payload.get('source_item_count') or len(items),
            }
        )

    # Rows are already DB-filtered by search_query above (see docs_qs.filter(...)),
    # so no further Python-side filtering is needed here.
    paginator = Paginator(rows, per_page)
    page_obj = paginator.get_page(page_number)

    return render(
        request,
        'planning/po_inbox.html',
        {
            'rows': page_obj.object_list,
            'page_obj': page_obj,
            'search_query': search_query,
            'per_page': per_page,
            'can_admin_actions': _user_is_admin(request.user),
            # True all-time count (distinct PO/WO numbers). Normally equals
            # page_obj.paginator.count too — they only diverge if the dedupe_cap
            # safety net above was actually hit (see is_capped_view).
            'total_document_count': PoDocument.objects.exclude(extracted_payload__isnull=True)
                .values('extracted_payload__po_number').distinct().count(),
            'is_capped_view': len(deduped_docs) >= dedupe_cap,
        }
    )




@login_required
@permission_required('can_edit_jobcard')
def upload_po(request):
    """Upload a PO or Work Order PDF, extract its content, store it, and redirect to review.

    Both document types share the same layout from the Utopia ERP system (see
    planning/po_extractor.py), which auto-detects PURCHASE ORDER vs WORK ORDER
    from the PDF itself and tags the payload's `document_type` accordingly —
    no separate upload flow is needed per type. The document number (PO- or
    WO-prefixed) is stored in the same `po_number` payload field either way,
    so all downstream keying (dedupe, PO Intake grouping, PlanningJob.po_number,
    _sync_repeat_jobs_from_po) works unchanged regardless of which was uploaded.
    """
    if request.method == 'POST':
        pdf_file = request.FILES.get('po_pdf')
        if not pdf_file:
            messages.error(request, 'Please select a PDF file.')
            return redirect('planning:upload_po')

        if not pdf_file.name.lower().endswith('.pdf'):
            messages.error(request, 'Only PDF files are supported.')
            return redirect('planning:upload_po')

        try:
            extracted = extract_po_from_pdf(pdf_file)
        except ValueError as exc:
            messages.error(request, f'Extraction failed: {exc}')
            return redirect('planning:upload_po')
        except Exception as exc:
            messages.error(request, f'Unexpected error reading PDF: {exc}')
            return redirect('planning:upload_po')

        # Surface partial-extraction warning before saving so user sees it on review page.
        extraction_warning = extracted.pop('extraction_warning', None)
        expected_count = extracted.pop('expected_line_count', None)

        items = extracted.get('items', [])
        source_item_count = len(items)
        deduped_items, duplicate_skus = _deduplicate_po_items_by_sku(items)
        extracted['items'] = deduped_items
        extracted['source_item_count'] = source_item_count
        extracted['merged_duplicate_skus'] = sorted(set(duplicate_skus))

        # Reset file pointer so Django can save it
        pdf_file.seek(0)

        po_number = (extracted.get('po_number') or '').strip()
        if po_number in ('-', '—', 'N/A', 'NA'):
            po_number = ''
            extracted['po_number'] = ''
        doc_label = document_type_label(po_number)
        existing_doc = None
        if po_number:
            existing_doc = PoDocument.objects.filter(extracted_payload__po_number=po_number).order_by('-id').first()

        if existing_doc:
            existing_payload = existing_doc.extracted_payload or {}
            existing_items, _ = _deduplicate_po_items_by_sku(existing_payload.get('items', []))
            merged_items, added_skus, updated_skus, ignored_lines = _merge_po_items_for_existing_po(
                existing_items,
                deduped_items,
            )

            merged_payload = dict(existing_payload)
            merged_payload.update(extracted)
            merged_payload['items'] = merged_items
            merged_payload['source_item_count'] = len(merged_items)
            merged_payload['merged_duplicate_skus'] = sorted(
                set(existing_payload.get('merged_duplicate_skus') or []) | set(duplicate_skus)
            )
            configured_skus = sorted(set(existing_payload.get('new_skus_configured') or []))
            if configured_skus:
                merged_payload['new_skus_configured'] = configured_skus

            existing_doc.po_file = pdf_file
            existing_doc.extracted_payload = merged_payload
            existing_doc.extraction_status = 'processed'
            existing_doc.uploaded_by = request.user
            existing_doc.save(update_fields=['po_file', 'extracted_payload', 'extraction_status', 'uploaded_by'])
            po_doc = existing_doc
        else:
            po_doc = PoDocument.objects.create(
                po_file=pdf_file,
                extracted_payload=extracted,
                extraction_status='processed',
                uploaded_by=request.user,
            )

        sync_result = _sync_repeat_jobs_from_po(po_doc, actor=request.user)
        item_count = len(extracted.get('items', []))
        if existing_doc and ignored_lines:
            preview = ', '.join(
                f"{row['sku']} ({row['qty'] if row['qty'] is not None else '-'})"
                for row in ignored_lines[:8]
            )
            remainder = len(ignored_lines) - 8
            if remainder > 0:
                preview += f" +{remainder} more"
            messages.warning(
                request,
                f"Ignored duplicate line(s) for same {doc_label} (same SKU and Qty): {preview}.",
            )

        if extraction_warning:
            messages.warning(
                request,
                f"Partial extraction: {extraction_warning}",
            )
        else:
            if existing_doc:
                final_item_count = len((existing_doc.extracted_payload or {}).get('items', []))
                msg = (
                    f"{doc_label} {extracted.get('po_number', '?')} updated. "
                    f"Unique added SKU(s): {len(added_skus)}; corrected SKU(s): {len(updated_skus)}; "
                    f"current {doc_label} line count: {final_item_count}."
                )
                if duplicate_skus:
                    msg += f" Duplicate SKU lines merged in upload: {', '.join(sorted(set(duplicate_skus)))}."
                msg += f" Same {doc_label} + same SKU + same Qty lines are ignored."
            else:
                msg = (
                    f"{doc_label} {extracted.get('po_number', '?')} extracted with "
                    f"{item_count} of {expected_count or item_count} line items. Sent to PO Intake queue."
                )
                if duplicate_skus:
                    msg += f" Duplicate SKU lines merged: {', '.join(sorted(set(duplicate_skus)))}."
            if sync_result['created'] or sync_result['updated']:
                msg += (
                    f" Draft planning jobs synced from {doc_label} lines: created {sync_result['created']}, "
                    f"updated {sync_result['updated']}."
                )
            if sync_result['missing_recipe']:
                msg += f" SKU(s) pending approved master data: {sync_result['missing_recipe']}."
            messages.success(request, msg)
        for jc_number, matched_pr, matched_po in sync_result.get('pr_matched', []):
            messages.info(
                request,
                f'JC {jc_number} was opened under PR {matched_pr} — linked to {document_type_label(matched_po)} {matched_po} now; '
                f'no duplicate job card was created.',
            )
        return redirect('qc:po_review', doc_id=po_doc.id)

    return render(request, 'planning/po_upload.html')


@login_required
@permission_required('can_edit_jobcard')
def manual_po_entry(request):
    """Create a PO or Work Order intake record manually without uploading a PDF.

    Business logic is identical for both — only the number's prefix and the
    labels shown to the user differ. The selected `document_type` just tags
    the payload for display (see document_type_label()); nothing downstream
    branches on it.
    """
    if request.method == 'POST':
        document_type = (request.POST.get('document_type') or 'PO').strip().upper()
        if document_type not in ('PO', 'WO'):
            document_type = 'PO'
        doc_label = 'Work Order' if document_type == 'WO' else 'PO'

        po_number = (request.POST.get('po_number') or '').strip()
        if po_number in ('-', '—', 'N/A', 'NA'):
            po_number = ''
        pr_number = (request.POST.get('pr_number') or '').strip()

        if not po_number and not pr_number:
            messages.error(request, f'Either a {doc_label} Number or a PR Number is required.')
            return redirect('planning:manual_po_entry')

        items = []
        line_indexes = request.POST.getlist('item_index')
        for index in line_indexes:
            sku = (request.POST.get(f'manual_sku_{index}') or '').strip()
            if not sku:
                continue
            quantity = _to_int(request.POST.get(f'manual_quantity_{index}'))
            if quantity is None:
                messages.error(request, f'Quantity must be a valid number for line {index}.')
                return redirect('planning:manual_po_entry')

            unit_cost = _to_decimal(request.POST.get(f'manual_unit_cost_{index}'))
            if unit_cost is None:
                messages.error(request, f'Unit cost must be a valid number for line {index}.')
                return redirect('planning:manual_po_entry')

            net_total = None
            if quantity is not None and unit_cost is not None:
                net_total = unit_cost * Decimal(quantity)

            item = {
                'sku': sku,
                'job_name': (request.POST.get(f'manual_job_name_{index}') or '').strip() or sku,
                'quantity': quantity,
                'unit': (request.POST.get(f'manual_unit_{index}') or '').strip() or 'Pcs',
                'delivery_date': (request.POST.get(f'manual_delivery_date_{index}') or '').strip() or '',
                'unit_cost': _format_decimal_string(unit_cost),
                'net_total': _format_decimal_string(net_total),
            }
            items.append(item)

        if not items:
            messages.error(request, 'At least one PO line is required.')
            return redirect('planning:manual_po_entry')

        sku_keys = [_sku_key(item['sku']) for item in items if item.get('sku')]
        duplicate_sku_keys = [sku for sku in sku_keys if sku_keys.count(sku) > 1]
        if duplicate_sku_keys:
            messages.error(request, f'Duplicate SKUs are not allowed within the same {doc_label}. Please remove duplicate lines before saving.')
            return redirect('planning:manual_po_entry')

        supplier_name = (request.POST.get('supplier_name') or '').strip() or 'UTOPIA PRINTING & PACKAGING'
        buyer_name = (request.POST.get('buyer_name') or '').strip() or 'UTOPIA INDUSTRIES (PVT.) LTD.'
        grand_total = sum((Decimal(item['net_total']) if item.get('net_total') is not None else Decimal('0')) for item in items)

        payload = {
            'po_number': po_number,
            'document_type': document_type,
            'pr_number': pr_number,
            'po_date': (request.POST.get('po_date') or '').strip(),
            'approval_date': (request.POST.get('approval_date') or '').strip(),
            'department': (request.POST.get('department') or '').strip(),
            'delivery_location': (request.POST.get('delivery_location') or '').strip(),
            'supplier_name': supplier_name,
            'buyer_name': buyer_name,
            'grand_total': _format_decimal_string(grand_total),
            'items': items,
            'source_item_count': len(items),
        }

        manual_file = ContentFile(b'', name=f'manual_po_{po_number}_{timezone.now().strftime("%Y%m%d%H%M%S")}.txt')
        po_doc = PoDocument.objects.create(
            po_file=manual_file,
            extracted_payload=payload,
            extraction_status='processed',
            uploaded_by=request.user,
        )

        sync_result = _sync_repeat_jobs_from_po(po_doc, actor=request.user)
        ref_label = f'{doc_label} {po_number}' if po_number else f'PR {pr_number}'
        messages.success(request, f'Manual {ref_label} created with {len(items)} line(s).')
        if sync_result['created'] or sync_result['updated']:
            messages.success(
                request,
                f'Draft planning jobs synced from {doc_label} lines: created {sync_result["created"]}, updated {sync_result["updated"]}.',
            )
        if sync_result['missing_recipe']:
            messages.warning(request, f'SKU(s) pending approved master data: {sync_result["missing_recipe"]}.')
        for jc_number, matched_pr, matched_po in sync_result.get('pr_matched', []):
            messages.info(
                request,
                f'JC {jc_number} was opened under PR {matched_pr} — linked to {document_type_label(matched_po)} {matched_po} now; '
                f'no duplicate job card was created.',
            )

        return redirect('qc:po_review', doc_id=po_doc.id)

    return render(request, 'planning/manual_po_entry.html')


@login_required
@permission_required('can_edit_jobcard')
def po_review(request, doc_id):
    """Review extracted PO data and create PlanningJob records."""
    po_doc = get_object_or_404(PoDocument, id=doc_id)
    payload = po_doc.extracted_payload or {}
    items = _po_payload_items(payload)
    sku_counts = {}
    for item in payload.get('items', []) or []:
        sku_key = _sku_key(item.get('sku'))
        if sku_key:
            sku_counts[sku_key] = sku_counts.get(sku_key, 0) + 1
    duplicate_skus = [sku for sku, count in sku_counts.items() if count > 1]
    if duplicate_skus:
        messages.error(
            request,
            f'Duplicate SKUs are not allowed in the same PO. Please remove duplicate lines for: {", ".join(sorted(duplicate_skus))}.',
        )
    ignored_skus = sorted({s for s in (payload.get('new_skus_ignored') or []) if s})
    po_number = payload.get('po_number') or ''
    configured_new_skus = {_sku_key(sku) for sku in (payload.get('new_skus_configured') or []) if sku}
    recipe_map = _build_recipe_map(items)
    from planning.sku_classification import build_sku_doc_index
    sku_doc_index = build_sku_doc_index()
    annotated_items, repeat_count, new_count, missing_skus = _annotate_items_with_recipe(
        items,
        recipe_map,
        current_po_number=po_number,
        po_doc_created_at=po_doc.created_at,
        po_doc_id=po_doc.id,
        sku_doc_index=sku_doc_index,
    )

    item_sku_keys = {_sku_key(item.get('sku')) for item in annotated_items if item.get('sku')}
    existing_jobs_by_sku = {}
    if po_number and item_sku_keys:
        existing_jobs = PlanningJob.objects.filter(po_number=po_number).order_by('-updated_at', '-id')
        for job in existing_jobs:
            key = _sku_key(job.sku)
            if key in item_sku_keys and key not in existing_jobs_by_sku:
                existing_jobs_by_sku[key] = job

    existing_any_jobs_skus = set()
    if item_sku_keys:
        sku_any_query = Q()
        for sku_key in item_sku_keys:
            sku_any_query |= Q(sku__iexact=sku_key)
        query_qs = PlanningJob.objects.filter(sku_any_query)
        if po_number:
            query_qs = query_qs.exclude(po_number=po_number)
        existing_any_jobs_skus = {
            _sku_key(sku)
            for sku in query_qs.values_list('sku', flat=True)
            if sku
        }

    seen_skus_in_payload = set()
    for item in annotated_items:
        sku_key = _sku_key(item.get('sku'))
        from planning.sku_classification import classify_po_line

        line_label, _reason = classify_po_line(
            item.get('sku'),
            po_number,
            po_doc_created_at=po_doc.created_at,
            po_doc_id=po_doc.id,
            recipe=recipe_map.get(sku_key),
            explicit_repeat_flag=(item.get('repeat_flag') or item.get('repeat')),
        )
        item['forward_flag'] = line_label
        item['is_first_production'] = line_label == 'New'
        existing_job = existing_jobs_by_sku.get(sku_key)
        item['existing_job_id'] = existing_job.id if existing_job else None
        item['existing_jc_number'] = existing_job.jc_number if existing_job else ''
        if sku_key:
            seen_skus_in_payload.add(sku_key)

    # Display counts based on approved SKU master data availability, not existing PlanningJob history.
    repeat_count = sum(1 for item in annotated_items if item.get('is_repeat'))
    new_count = sum(1 for item in annotated_items if not item.get('is_repeat'))

    if request.method == 'POST':
        action = request.POST.get('action', '')
        if action == 'ignore':
            sku = (request.POST.get('sku') or '').strip()
            if not sku:
                messages.error(request, 'SKU is required for ignore action.')
                return redirect('qc:po_review', doc_id=po_doc.id)

            ignored = {
                _sku_key(s)
                for s in (payload.get('new_skus_ignored') or [])
                if s
            }
            ignored.add(_sku_key(sku))
            payload['new_skus_ignored'] = sorted(ignored)
            po_doc.extracted_payload = payload
            po_doc.save(update_fields=['extracted_payload'])

            po_number = (payload.get('po_number') or '').strip()
            if po_number:
                PlanningJob.objects.filter(
                    po_number__iexact=po_number,
                    sku__iexact=sku,
                    status__iexact='draft',
                    is_active=True,
                ).update(is_active=False, updated_at=timezone.now())

            messages.success(request, f'SKU {sku} ignored and removed from PO intake review.')
            return redirect('qc:po_review', doc_id=po_doc.id)

        if action == 'update_po_number':
            manual_po_number = (request.POST.get('manual_po_number') or '').strip()
            if manual_po_number in ('-', '—', 'N/A', 'NA'):
                manual_po_number = ''
            if not manual_po_number:
                messages.error(request, 'PO number is required to update the PO intake record.')
                return redirect('qc:po_review', doc_id=po_doc.id)

            payload['po_number'] = manual_po_number
            po_doc.extracted_payload = payload
            po_doc.save(update_fields=['extracted_payload'])
            messages.success(request, f'PO number updated to {manual_po_number}.')
            return redirect('qc:po_review', doc_id=po_doc.id)

        if action == 'add_manual_item':
            sku = (request.POST.get('manual_sku') or '').strip()
            if not sku:
                messages.error(request, 'SKU is required to add a manual PO line.')
                return redirect('qc:po_review', doc_id=po_doc.id)

            sku_key = _sku_key(sku)
            existing_skus = {_sku_key(item.get('sku')) for item in payload.get('items', []) if item.get('sku')}
            if sku_key in existing_skus:
                messages.error(request, f'SKU {sku} is already present on this PO. Duplicate SKUs are not allowed.')
                return redirect('qc:po_review', doc_id=po_doc.id)

            quantity = _to_int(request.POST.get('manual_quantity'))
            if quantity is None:
                messages.error(request, 'Quantity must be a valid number to add a manual PO line.')
                return redirect('qc:po_review', doc_id=po_doc.id)

            unit_cost_value = _to_decimal(request.POST.get('manual_unit_cost'))
            net_total_value = _to_decimal(request.POST.get('manual_net_total'))
            manual_item = {
                'sku': sku,
                'job_name': (request.POST.get('manual_job_name') or '').strip() or sku,
                'quantity': quantity,
                'unit': (request.POST.get('manual_unit') or '').strip() or '',
                'delivery_date': (request.POST.get('manual_delivery_date') or '').strip() or '',
                'unit_cost': _format_decimal_string(unit_cost_value),
                'net_total': _format_decimal_string(net_total_value),
                'print_sheet_size': (request.POST.get('manual_print_sheet_size') or '').strip() or '',
                'purchase_sheet_size': (request.POST.get('manual_purchase_sheet_size') or '').strip() or '',
                'ups': _to_optional_decimal(request.POST.get('manual_ups')),
                'machine_name': (request.POST.get('manual_machine_name') or '').strip() or '',
            }
            payload['items'] = list(payload.get('items', []))
            payload['items'].append(manual_item)
            po_doc.extracted_payload = payload
            po_doc.save(update_fields=['extracted_payload'])
            messages.success(request, f'Manual PO line for SKU {sku} added.')
            return redirect('qc:po_review', doc_id=po_doc.id)

        if action == 'create_jobs':
            sku_counts = {}
            for item in annotated_items:
                sku_key = _sku_key(item.get('sku'))
                if not sku_key:
                    continue
                sku_counts[sku_key] = sku_counts.get(sku_key, 0) + 1

            duplicate_skus = [sku for sku, count in sku_counts.items() if count > 1]
            if duplicate_skus:
                messages.error(
                    request,
                    'Duplicate SKUs are not allowed in the same PO. Remove duplicate lines before creating jobs.',
                )
                return redirect('qc:po_review', doc_id=po_doc.id)

            with transaction.atomic():
                created_count = 0
                updated_count = 0
                skipped_count = 0
                locked_count = 0
                missing_recipe_count = 0
                po_date_raw = payload.get('po_date')
                po_date = _parse_iso_date(po_date_raw)
                approval_date_raw = payload.get('approval_date')
                approval_date = _parse_iso_date(approval_date_raw)
                delivery_location = payload.get('delivery_location') or ''
                department = payload.get('department') or ''

                for item in annotated_items:
                    sku = (item.get('sku') or '').strip()
                    job_name = (item.get('job_name') or '').strip() or sku
                    if not sku:
                        skipped_count += 1
                        continue

                    field_prefix = f"item_{item['line_no']}_"
                    skip_flag = request.POST.get(f"{field_prefix}skip") == '1'

                    if skip_flag:
                        skipped_count += 1
                        continue

                    recipe = recipe_map.get(_sku_key(sku))
                    is_approved = bool(recipe and recipe.master_data_status == 'approved')
                    if not is_approved:
                        missing_recipe_count += 1
                        continue

                    sku_key = _sku_key(sku)
                    delivery_date = _parse_iso_date(item.get('delivery_date'))
                    plan_date = po_doc.created_at.date() if po_doc and getattr(po_doc, 'created_at', None) else (delivery_date or po_date)

                    existing_job = existing_jobs_by_sku.get(sku_key)
                    if existing_job:
                        if _normalize_status(existing_job.status) != 'draft':
                            locked_count += 1
                            continue
                        jc_number = existing_job.jc_number
                    else:
                        jc_number = allocate_next_jc_number(plan_date)

                    current_requirement = existing_job.requirement if existing_job else ''

                    from planning.sku_classification import repeat_flag_value_for_po_line

                    repeat_flag_value = repeat_flag_value_for_po_line(
                        item,
                        po_number=po_number,
                        po_doc_created_at=po_doc.created_at,
                        po_doc_id=po_doc.id,
                        recipe=recipe,
                        existing_job=existing_job,
                        sku_doc_index=sku_doc_index,
                    )
                    forward_as_new = repeat_flag_value == 'New'

                    qty = item.get('quantity')
                    order_qty = int(qty) if qty is not None else None

                    unit_cost_val = item.get('unit_cost')
                    unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None

                    defaults = {
                        'po_number': po_number,
                        'sku': sku,
                        'job_name': recipe.job_name or job_name,
                        'order_qty': order_qty,
                        'department': department,
                        'destination': delivery_location,
                        'po_approval_date': approval_date,
                        'delivery_date': delivery_date,
                        'unit_cost': unit_cost_dec if unit_cost_dec is not None else recipe.default_unit_cost,
                        'status': 'draft',
                        'repeat_flag': repeat_flag_value,
                        'requirement': _sync_new_sku_requirement(current_requirement, forward_as_new),
                        'material': recipe.material,
                        'color_spec': recipe.color_spec,
                        'application': recipe.application,
                        'size_w_mm': recipe.size_w_mm,
                        'size_h_mm': recipe.size_h_mm,
                        'ups': recipe.ups,
                        'print_sheet_size': recipe.print_sheet_size,
                        'purchase_sheet_size': recipe.purchase_sheet_size,
                    }
                    if plan_date:
                        defaults['plan_date'] = plan_date

                    job_obj, created = PlanningJob.objects.update_or_create(
                        jc_number=jc_number,
                        defaults=defaults,
                    )
                    if created:
                        created_count += 1
                    else:
                        updated_count += 1
                    existing_jobs_by_sku[sku_key] = job_obj
                    if sku_key:
                        existing_any_jobs_skus.add(sku_key)

            po_doc.extraction_status = 'processed'
            po_doc.save(update_fields=['extraction_status'])

            if created_count == 0 and updated_count == 0:
                if missing_recipe_count > 0:
                    messages.warning(
                        request,
                        'This PO contains only new SKUs with missing master data. Please configure them in Pending SKUs before sending to planning.',
                    )
                    return redirect('qc:pending_skus')
                messages.warning(
                    request,
                    f'No jobs created. Skipped {skipped_count}, missing-recipe {missing_recipe_count}, locked-skip {locked_count}. Add missing SKU master data from Pending SKUs and run create again.',
                )
                return redirect('qc:pending_skus')

            messages.success(
                request,
                f'Done. Created {created_count}, updated {updated_count}, skipped {skipped_count}, missing-recipe {missing_recipe_count}, locked-skip {locked_count} planning job(s).',
            )
            if missing_recipe_count > 0:
                messages.warning(
                    request,
                    f'{missing_recipe_count} SKU(s) are still pending master data. Open Pending SKUs tab to configure them.',
                )
            return redirect('qc:approval_queue')

    context = {
        'po_doc': po_doc,
        'payload': payload,
        'items': annotated_items,
        'items_json': json.dumps(annotated_items),
        'repeat_count': repeat_count,
        'new_count': new_count,
        'configured_new_count': len(configured_new_skus),
        'missing_skus': missing_skus,
        'ignored_skus': ignored_skus,
        'ignored_count': len(ignored_skus),
    }
    return render(request, 'planning/po_review.html', context)


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def po_new_skus(request, doc_id):
    po_doc = get_object_or_404(PoDocument, id=doc_id)
    payload = po_doc.extracted_payload or {}
    items = _po_payload_items(payload)

    recipe_map = _build_recipe_map(items)
    from planning.sku_classification import build_sku_doc_index
    _, _, _, missing_skus = _annotate_items_with_recipe(items, recipe_map, sku_doc_index=build_sku_doc_index())
    missing_recipe_defaults = {}
    if missing_skus:
        recipe_query = Q()
        for sku in missing_skus:
            recipe_query |= Q(sku__iexact=sku)
        for recipe in SkuRecipe.objects.filter(recipe_query):
            missing_recipe_defaults[recipe.sku.upper()] = recipe
    missing_sku_rows = [
        {
            'sku': sku,
            'recipe': missing_recipe_defaults.get(_sku_key(sku)),
        }
        for sku in missing_skus
    ]

    if request.method == 'POST':
        created_count = 0
        saved_skus = []
        for sku in missing_skus:
            prefix = f"sku_{sku}"
            job_name = (request.POST.get(f"{prefix}_job_name") or '').strip()
            material = (request.POST.get(f"{prefix}_material") or '').strip()
            color_spec = (request.POST.get(f"{prefix}_color_spec") or '').strip()
            application = (request.POST.get(f"{prefix}_application") or '').strip()
            print_sheet_size = (request.POST.get(f"{prefix}_print_sheet_size") or '').strip()
            purchase_sheet_size = (request.POST.get(f"{prefix}_purchase_sheet_size") or '').strip()
            ups = _to_optional_decimal(request.POST.get(f"{prefix}_ups"))

            unit_cost_raw = (request.POST.get(f"{prefix}_default_unit_cost") or '').strip()
            unit_cost = None
            if unit_cost_raw:
                try:
                    unit_cost = Decimal(unit_cost_raw)
                except InvalidOperation:
                    unit_cost = None

            if not job_name and not material:
                # Keep save requirements simple, but avoid empty recipe rows.
                continue

            SkuRecipe.objects.update_or_create(
                sku=sku,
                defaults={
                    'job_name': job_name,
                    'material': material,
                    'color_spec': color_spec,
                    'application': application,
                    'print_sheet_size': print_sheet_size,
                    'purchase_sheet_size': purchase_sheet_size,
                    'ups': ups,
                    'default_unit_cost': unit_cost,
                    'created_by': request.user,
                    'master_data_status': 'draft',
                    'reviewed_by': None,
                    'reviewed_at': None,
                    'approved_by': None,
                    'approved_at': None,
                },
            )
            created_count += 1
            saved_skus.append(sku)

        if saved_skus:
            configured = set(payload.get('new_skus_configured') or [])
            configured.update(saved_skus)
            payload['new_skus_configured'] = sorted(configured)
            po_doc.extracted_payload = payload
            po_doc.save(update_fields=['extracted_payload'])

        messages.success(
            request,
            f'SKU recipes saved: {created_count}. Planning jobs remain in draft until SKU master approval unlocks refresh.',
        )
        return redirect('qc:po_review', doc_id=po_doc.id)

    return render(
        request,
        'planning/po_new_skus.html',
        {
            'po_doc': po_doc,
            'payload': payload,
            'missing_skus': missing_skus,
            'missing_sku_rows': missing_sku_rows,
            'missing_recipe_defaults': missing_recipe_defaults,
            'example_form': SkuRecipeForm(),
        },
    )


@login_required
def request_wastage_machine_change(request, job_id):
    if request.method != 'POST':
        return redirect('planning:job_detail', job_id=job_id)
        
    job = get_object_or_404(PlanningJob, id=job_id)
    profile = getattr(request.user, 'profile', None)
    user_can_plan = profile.can_plan() if profile else False
    
    if not user_can_plan:
        messages.error(request, 'You do not have permission to request planning changes.')
        return redirect('planning:job_detail', job_id=job_id)
        
    status_now = _normalize_status(job.status)
    locked_statuses = {'qc_approved', 'released', 'in_production'}
    if status_now not in locked_statuses:
        messages.error(request, 'Reopen requests can only be submitted for QC approved, released, or in production job cards.')
        return redirect('planning:job_detail', job_id=job_id)
        
    has_pending = JobCardChangeRequest.objects.filter(planning_job=job, status='pending').exists()
    if has_pending:
        messages.error(request, 'There is already a pending reopen request for this job card.')
        return redirect('planning:job_detail', job_id=job_id)
        
    reason = (request.POST.get('reason') or '').strip()
    
    if not reason:
        messages.error(request, 'A justification reason is required to submit a reopen request.')
        return redirect('planning:job_detail', job_id=job_id)
        
    JobCardChangeRequest.objects.create(
        planning_job=job,
        request_type='reopen_to_draft',
        reason=reason,
        status='pending',
        requested_by=request.user
    )
    
    messages.success(request, f'Reopen request for {job.jc_number} has been submitted for Production Manager approval.')
    return redirect('planning:job_detail', job_id=job_id)


@login_required
@transaction.atomic
def approve_change_request(request, request_id):
    if request.method != 'POST':
        return redirect('planning:approval_queue')
        
    change_request = get_object_or_404(JobCardChangeRequest, id=request_id)
    profile = getattr(request.user, 'profile', None)
    user_can_approve_pm = profile.can_approve_pm() if profile else False
    
    if not user_can_approve_pm:
        messages.error(request, 'You do not have permission to approve change requests.')
        return redirect('planning:approval_queue')
        
    if change_request.status != 'pending':
        messages.error(request, 'This request is already resolved.')
        return redirect('planning:approval_queue')
        
    job = change_request.planning_job
    job_card = getattr(job, 'job_card', None)

    if change_request.is_cancellation:
        try:
            approve_job_cancellation(change_request, actor=request.user)
        except ValidationError as exc:
            _report_validation_error(request, exc, prefix=job.jc_number)
            return redirect('planning:approval_queue')

        change_request.status = 'approved'
        change_request.approved_by = request.user
        change_request.approved_at = timezone.now()
        change_request.save(update_fields=['status', 'approved_by', 'approved_at'])
        messages.success(request, f'Cancellation approved. Planning job {job.jc_number} is now cancelled.')
        return redirect('planning:approval_queue')

    old_status = job.status
    job.status = 'draft'
    job.issued_to_production = False
    job.save()
    
    if job_card:
        from core.jobcard_service import transition_job_card_status
        transition_job_card_status(
            job_card,
            'draft',
            actor=request.user,
            reason=f"Approved edit reopen request: {change_request.reason}"
        )
        
        ChangeLog.objects.create(
            entity_type='job_card',
            record_id=job_card.pk,
            record_label=str(job_card),
            action='reopen',
            changed_by=request.user,
            change_reason=f"Approved reopen request: {change_request.reason}",
            field_changes={
                'status': {
                    'label': 'Workflow Status',
                    'from': old_status,
                    'to': 'draft'
                }
            }
        )
            
    change_request.status = 'approved'
    change_request.approved_by = request.user
    change_request.approved_at = timezone.now()
    change_request.save(update_fields=['status', 'approved_by', 'approved_at'])
    
    messages.success(request, f'Reopen request for {job.jc_number} approved. The job card is now in Draft and can be edited.')
    return redirect('planning:approval_queue')


@login_required
def reject_change_request(request, request_id):
    if request.method != 'POST':
        return redirect('planning:approval_queue')
        
    change_request = get_object_or_404(JobCardChangeRequest, id=request_id)
    profile = getattr(request.user, 'profile', None)
    user_can_approve_pm = profile.can_approve_pm() if profile else False
    
    if not user_can_approve_pm:
        messages.error(request, 'You do not have permission to reject change requests.')
        return redirect('planning:approval_queue')
        
    if change_request.status != 'pending':
        messages.error(request, 'This request is already resolved.')
        return redirect('planning:approval_queue')
        
    rejection_reason = (request.POST.get('rejection_reason') or '').strip()
    
    change_request.status = 'rejected'
    change_request.approved_by = request.user
    change_request.approved_at = timezone.now()
    change_request.rejection_reason = rejection_reason
    change_request.save(update_fields=['status', 'approved_by', 'approved_at', 'rejection_reason'])
    
    messages.warning(request, f'Reopen request for {change_request.planning_job.jc_number} has been rejected.')
    return redirect('planning:approval_queue')



# ---------------------------------------------------------------------------
# Smart layout merge (ganging)
# ---------------------------------------------------------------------------

def _eligible_merge_jobs(request=None):
    """Jobs that have not been printed yet and are free to be ganged.

    Jobs whose plates already exist are still offered: the saving is not only the
    plate set but one make-ready, one setup and a longer run. The planner decides
    per member at accept time whether to scrap or retain the existing plate
    (see ``existing_plate_request`` on each returned job).
    """
    printed_job_ids = set(
        PlanningPrintRun.objects.filter(print_qty__gt=0).values_list('planning_job_id', flat=True)
    )
    merged_job_ids = set(
        MergeGroupItem.objects.filter(merge_group__status__in=MERGE_GROUP_OPEN_STATUSES)
        .values_list('planning_job_id', flat=True)
    )
    queryset = (
        PlanningJob.objects.filter(is_active=True, is_cancelled=False)
        .exclude(status__in=['in_production', 'completed', 'cancelled'])
        .exclude(issued_to_production=True)
        .exclude(job_process_type='cut_and_pack')
        .exclude(id__in=printed_job_ids | merged_job_ids)
    )
    if request is not None:
        material = (request.GET.get('material') or '').strip()
        if material:
            queryset = queryset.filter(material__icontains=material)
        plan_month = (request.GET.get('plan_month') or '').strip()
        if plan_month:
            queryset = queryset.filter(plan_month__iexact=plan_month)
    return queryset


def _annotate_existing_plate_requests(jobs):
    """Attach each job's live (non-archived) plate request, if any.

    Merging is still allowed when plates exist — the planner decides per member
    whether to scrap or retain them — so the board must show what is at stake.
    """
    from printing_plates.models import PlateRequest

    job_ids = [job.id for job in jobs]
    by_job = {}
    for plate_request in (
        PlateRequest.objects.exclude(status=PlateRequest.STATUS_ARCHIVED)
        .filter(planning_job_id__in=job_ids)
        .order_by('planning_job_id', '-requested_at', '-id')
    ):
        by_job.setdefault(plate_request.planning_job_id, plate_request)
    for job in jobs:
        job.existing_plate_request = by_job.get(job.id)
    return jobs


def _apply_merge_plate_dispositions(jobs, plate_actions, group, actor=None, reason=''):
    """Scrap or retain each member's pre-existing plate set for a new merge group.

    Only the lead keeps a live plate path — its own request (if any) is left
    alone, because the combined plate is raised against the lead.
    """
    from printing_plates.models import PlateRequest
    from printing_plates.services import retain_plate_for_reuse, scrap_plate_for_merge

    scrapped, retained = [], []
    for job in jobs:
        if job.id == group.lead_job_id:
            continue
        plate_request = (
            PlateRequest.objects.exclude(status=PlateRequest.STATUS_ARCHIVED)
            .filter(planning_job_id=job.id)
            .order_by('-requested_at', '-id')
            .first()
        )
        if not plate_request:
            continue
        if plate_actions.get(job.id) == 'retain':
            retain_plate_for_reuse(plate_request, actor=actor, merge_code=group.code, reason=reason)
            retained.append(plate_request)
        else:
            scrap_plate_for_merge(plate_request, actor=actor, merge_code=group.code, reason=reason)
            scrapped.append(plate_request)
    return scrapped, retained


def _merge_config_from_request(request):
    cfg = MergeConfig()
    try:
        window = int(request.GET.get('delivery_window') or 0)
    except (TypeError, ValueError):
        window = 0
    cfg.delivery_window_days = max(window, 0)
    return cfg


def _near_miss_blocked_jobs(jobs, cfg, limit=50):
    """Jobs that would be mergeable if someone completed their master data.

    Jobs that already look like a partner for a live candidate (same rough piece
    size and material) are listed first, but a job missing its size cannot be
    bucketed at all — and those are exactly the ones needing attention — so every
    blocked job is listed, capped so the board stays scannable.
    """
    mergeable_keys = set()
    blocked = []
    for job in jobs:
        if merge_blockers(job, cfg):
            blocked.append(job)
        else:
            key = size_key(job, cfg.size_tolerance_mm)
            if key:
                mergeable_keys.add((key, normalise_material(job.material)))

    rows = []
    for job in blocked:
        key = size_key(job, cfg.size_tolerance_mm)
        material = normalise_material(job.material)
        rows.append({
            'job': job,
            'reasons': merge_blockers(job, cfg),
            'is_near_miss': bool(key and material and (key, material) in mergeable_keys),
        })
    rows.sort(key=lambda row: (not row['is_near_miss'], len(row['reasons'])))
    return rows[:limit]


@login_required
@permission_required('can_view_planning_queue')
def planning_merge_board(request):
    cfg = _merge_config_from_request(request)
    jobs = _annotate_existing_plate_requests(list(_eligible_merge_jobs(request)))
    suggestions = build_suggestions(jobs, cfg)
    open_groups = (
        MergeGroup.objects.filter(status__in=MERGE_GROUP_OPEN_STATUSES)
        .prefetch_related('items__planning_job')
    )
    context = {
        'suggestions': suggestions,
        'open_groups': open_groups,
        'eligible_count': len(jobs),
        'blocked_rows': _near_miss_blocked_jobs(jobs, cfg),
        'config': cfg,
        'filters': {
            'material': (request.GET.get('material') or '').strip(),
            'plan_month': (request.GET.get('plan_month') or '').strip(),
            'delivery_window': cfg.delivery_window_days or '',
        },
    }
    return render(request, 'planning/planning_merge_board.html', context)


@login_required
@permission_required('can_plan')
def planning_merge_accept(request):
    if request.method != 'POST':
        return redirect('planning:merge_board')

    try:
        job_ids = sorted({int(value) for value in request.POST.getlist('job_ids')})
    except (TypeError, ValueError):
        job_ids = []
    if len(job_ids) < 2:
        messages.error(request, 'Select at least two jobs to merge.')
        return redirect('planning:merge_board')

    cfg = MergeConfig()
    with transaction.atomic():
        # Re-validate against the live eligibility set; never trust posted ups.
        jobs = list(_eligible_merge_jobs().select_for_update().filter(id__in=job_ids))
        if len(jobs) != len(job_ids):
            messages.error(request, 'One or more jobs are no longer eligible for merging. Refresh the board.')
            return redirect('planning:merge_board')

        # Per-member decision about an existing plate set: scrap it (the combined
        # plate supersedes it), retain it for a future run of that SKU, or drop
        # the job from this merge entirely.
        plate_actions = {
            job.id: (request.POST.get(f'plate_action_{job.id}') or 'scrap').strip().lower()
            for job in jobs
        }
        excluded = [job for job in jobs if plate_actions.get(job.id) == 'exclude']
        if excluded:
            jobs = [job for job in jobs if plate_actions.get(job.id) != 'exclude']
            if len(jobs) < 2:
                messages.error(
                    request,
                    'Excluding those jobs leaves fewer than two SKUs — nothing left to merge.',
                )
                return redirect('planning:merge_board')

        signatures = {bucket_signature(job, cfg) for job in jobs}
        if len(signatures) != 1 or None in signatures:
            messages.error(request, 'These jobs do not share the same size, material and colour specification.')
            return redirect('planning:merge_board')

        allocation = allocate_ups(jobs, jobs[0].ups_value, cfg)
        if not allocation:
            if excluded:
                messages.error(
                    request,
                    'After excluding '
                    + ', '.join(job.jc_number for job in excluded)
                    + ', the remaining quantities no longer split into whole ups within the '
                    f'{cfg.qty_tolerance_pct:g}% tolerance. Adjust the selection and try again.',
                )
            else:
                messages.error(request, 'Quantities cannot be split into whole ups within the 5% tolerance.')
            return redirect('planning:merge_board')

        savings = compute_savings(allocation, jobs, cfg)
        group_code = MergeGroup.next_code()
        group = MergeGroup.objects.create(
            code=group_code,
            artwork_code=MergeGroup.artwork_code_for(group_code),
            status='accepted',
            print_sheet_size=jobs[0].print_sheet_size or '',
            material=jobs[0].material or '',
            total_sheet_ups=allocation['sheet_ups'],
            run_sheets=allocation['run_sheets'],
            total_colors=jobs[0].total_colors or 0,
            plates_saved=savings['plates_saved'],
            makereadies_saved=savings['makereadies_saved'],
            setup_sheets_saved=savings['setup_sheets_saved'],
            mr_minutes_saved=savings['mr_minutes_saved'],
            impressions_saved=savings['impressions_saved'],
            notes=(request.POST.get('notes') or '').strip(),
            created_by=request.user,
            accepted_by=request.user,
            accepted_at=timezone.now(),
        )
        # The member with the most allocated ups leads: it carries the shared
        # plate request and the single printing run for the whole group.
        lead_item = max(allocation['items'], key=lambda item: item['allocated_ups'])
        MergeGroupItem.objects.bulk_create([
            MergeGroupItem(
                merge_group=group,
                planning_job=item['job'],
                allocated_ups=item['allocated_ups'],
                is_lead=(item is lead_item),
                source_awc_no=item['job'].awc_no_display or '',
                planned_produced_qty=item['planned_produced_qty'],
                net_qty=item['net_qty'],
                overage_pct=item['overage_pct'],
            )
            for item in allocation['items']
        ])
        group.lead_job = lead_item['job']
        group.save(update_fields=['lead_job'])

        scrapped, retained = _apply_merge_plate_dispositions(
            jobs, plate_actions, group, actor=request.user,
            reason=(request.POST.get('plate_reason') or '').strip(),
        )

    note = f'Merge group {group.code} created with {len(jobs)} jobs.'
    if excluded:
        note += ' Excluded: ' + ', '.join(job.jc_number for job in excluded) + '.'
    if scrapped:
        note += f' {len(scrapped)} existing plate set(s) scrapped as superseded.'
    if retained:
        note += f' {len(retained)} plate set(s) retained for a future run.'
    messages.success(request, note)
    return redirect('planning:merge_detail', group_id=group.id)


@login_required
@permission_required('can_view_planning_queue')
def planning_merge_detail(request, group_id):
    from printing_plates.models import PlateRequest
    from printing_plates.services import combined_plate_request_for_group, group_combined_plate_issued
    from planning.services import merge_layout_master_data_report

    group = get_object_or_404(
        MergeGroup.objects.prefetch_related('items__planning_job__job_card'),
        id=group_id,
    )
    # A member other than the lead still holding its own live plate set — either
    # a legacy group from before scrap/retain existed, or a retain choice.
    preexisting = (
        PlateRequest.objects.exclude(status=PlateRequest.STATUS_ARCHIVED)
        .filter(planning_job__in=[item.planning_job_id for item in group.items.all()])
        .exclude(planning_job_id=group.lead_job_id)
        .select_related('planning_job')
    )
    items = list(group.items.select_related('planning_job__job_card').all())
    # Mother-child: annotate each member so the child rows can show View/Print
    # like the regular jobs queue.
    printable_card_statuses = {'production_approved', 'released', 'in_production', 'completed', 'closed'}
    member_rows = []
    for item in items:
        job = item.planning_job
        card = getattr(job, 'job_card', None)
        member_rows.append({
            'item': item,
            'job': job,
            'can_print': bool(card and card.workflow_status in printable_card_statuses),
            'card_status_label': card.workflow_status_label if card else 'No job card yet',
        })
    combined_plate = combined_plate_request_for_group(group)

    released_members, unreleased_members = [], []
    for item in items:
        job_card = getattr(item.planning_job, 'job_card', None)
        if job_card and job_card.workflow_status in {'released', 'in_production', 'completed', 'closed'}:
            released_members.append(item)
        else:
            unreleased_members.append(item)

    return render(request, 'planning/planning_merge_detail.html', {
        'group': group,
        'preexisting_plate_requests': preexisting,
        'new_artwork_count': sum(1 for item in items if not (item.source_awc_no or '').strip()),
        'member_count': len(items),
        'combined_plate': combined_plate,
        'combined_plate_issued': group_combined_plate_issued(group),
        'unreleased_members': unreleased_members,
        'released_member_count': len(released_members),
        'all_members_released': not unreleased_members,
        'master_data_report': merge_layout_master_data_report(group) if group.is_open else [],
        'is_layout_approved': group.is_layout_approved,
        'material_origin_choices': PURCHASE_MATERIAL_ORIGIN_CHOICES,
        'member_rows': member_rows,
    })


@login_required
@permission_required('can_view_planning_queue')
def planning_merge_combined_sheet(request, group_id):
    """The single Combined Layout Sheet — the press document for a merged run.

    All combined-run detail lives here (not on the job cards, which stay in the
    familiar single-SKU format with just a watermark).
    """
    group = get_object_or_404(
        MergeGroup.objects.select_related('lead_job').prefetch_related('items__planning_job'),
        id=group_id,
    )
    return render(request, 'planning/planning_merge_combined_sheet.html', {'group': group})


@login_required
@permission_required('can_plan')
def planning_merge_approve_layout(request, group_id):
    """Group-level production approval for a combined layout.

    One approval releases every member for the merged run, standing in for each
    SKU's individual QC/PM/release gate. Refuses with a per-SKU list if any
    member's master data is incomplete.
    """
    from django.core.exceptions import ValidationError
    from planning.services import approve_merge_layout

    group = get_object_or_404(MergeGroup, id=group_id)
    if request.method != 'POST':
        return redirect('planning:merge_detail', group_id=group.id)
    if not group.is_open:
        messages.error(request, 'This merge group is closed.')
        return redirect('planning:merge_detail', group_id=group.id)

    try:
        approve_merge_layout(
            group, actor=request.user,
            combined_wastage=request.POST.get('combined_wastage'),
            material_origin=(request.POST.get('material_origin') or '').strip(),
        )
    except ValidationError as exc:
        messages.error(request, '; '.join(exc.messages))
        return redirect('planning:merge_detail', group_id=group.id)

    messages.success(
        request,
        f'Combined layout {group.code} approved for production — all '
        f'{group.items.count()} member cards released for the merged run.',
    )
    return redirect('planning:merge_detail', group_id=group.id)


@login_required
@permission_required('can_plan')
def planning_merge_raise_plate(request, group_id):
    """Raise the ONE combined plate request for the group, on the lead job only.

    Idempotent: if an open plate request already exists on the lead, it is
    reused rather than duplicated.
    """
    from printing_plates.services import combined_plate_request_for_group

    group = get_object_or_404(MergeGroup, id=group_id)
    if request.method != 'POST':
        return redirect('planning:merge_detail', group_id=group.id)
    if not group.is_open:
        messages.error(request, 'This merge group is closed.')
        return redirect('planning:merge_detail', group_id=group.id)
    if not group.lead_job_id:
        messages.error(request, 'This group has no lead job — cannot raise plates.')
        return redirect('planning:merge_detail', group_id=group.id)
    if not group.is_layout_approved:
        messages.error(
            request,
            'Approve the combined layout for production before raising its plate.',
        )
        return redirect('planning:merge_detail', group_id=group.id)

    existing = combined_plate_request_for_group(group)
    if existing:
        messages.info(request, f'Combined plate request #{existing.pk} is already open for {group.code}.')
        return redirect('planning:merge_detail', group_id=group.id)

    lead = group.lead_job
    plate_errors = [
        error for error in get_plate_making_prerequisite_errors(lead)
        if 'merged into layout' not in error  # that check does not apply to the lead itself
    ]
    if plate_errors:
        messages.error(request, ' '.join(plate_errors))
        return redirect('planning:merge_detail', group_id=group.id)

    from planning.sku_classification import plate_making_stage_for_repeat_flag

    lead.planning_stage = plate_making_stage_for_repeat_flag(lead.repeat_flag)
    lead.planning_stage_changed_at = timezone.now()
    lead.planning_stage_changed_by = request.user
    lead.save(update_fields=['planning_stage', 'planning_stage_changed_at', 'planning_stage_changed_by', 'updated_at'])

    plate_request = trigger_plate_request_for_planning_job(lead, request.user)
    if not plate_request:
        messages.error(
            request,
            f'Could not raise the combined plate for {lead.jc_number}. Check the job status and try again.',
        )
        return redirect('planning:merge_detail', group_id=group.id)

    group.status = 'artwork_ready'
    group.save(update_fields=['status'])
    messages.success(
        request,
        f'Combined plate request #{plate_request.pk} raised on lead job {lead.jc_number} for {group.code}.',
    )
    return redirect('printing_plates:request_detail', pk=plate_request.pk)


@login_required
@permission_required('can_plan')
def planning_merge_request_artwork(request, group_id):
    group = get_object_or_404(MergeGroup, id=group_id)
    if request.method != 'POST':
        return redirect('planning:merge_detail', group_id=group.id)
    if not group.is_open:
        messages.error(request, 'This merge group is closed.')
        return redirect('planning:merge_detail', group_id=group.id)

    group.status = 'artwork_requested'
    group.designer_notes = (request.POST.get('designer_notes') or '').strip()
    group.designer_requested_at = timezone.now()
    group.save(update_fields=['status', 'designer_notes', 'designer_requested_at'])
    messages.success(request, f'Combined artwork requested from design for {group.code}.')
    return redirect('planning:merge_detail', group_id=group.id)


@login_required
@permission_required('can_plan')
def planning_merge_cancel(request, group_id):
    group = get_object_or_404(MergeGroup, id=group_id)
    if request.method != 'POST':
        return redirect('planning:merge_detail', group_id=group.id)

    # Once the group's own combined plate is on the floor, a planner cannot
    # silently unpick the group — that is a production change requiring PM
    # approval. This checks the group's own plate request, not the lead's
    # (or a repeat SKU's) historical plate_set_no text, which does not mean a
    # combined plate exists.
    from printing_plates.services import group_combined_plate_issued

    if group_combined_plate_issued(group):
        messages.error(
            request,
            f'The combined plate for {group.code} is already issued to production. '
            f'Raise a change request for PM approval instead of cancelling the merge.',
        )
        return redirect('planning:merge_detail', group_id=group.id)

    group.status = 'cancelled'
    group.cancelled_at = timezone.now()
    group.save(update_fields=['status', 'cancelled_at'])
    messages.success(request, f'Merge group {group.code} cancelled; its jobs are back in the pool.')
    return redirect('planning:merge_board')
