import csv
import base64
import io
import json
from datetime import datetime
from decimal import Decimal, InvalidOperation

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Count, Q, Sum
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse

from core.jc_numbering import allocate_next_jc_number
from core.views import permission_required
from .forms import PlanningJobEditForm, SkuRecipeForm
from .models import PlanningDispatchRun, PlanningJob, PlanningPrintRun, PoDocument, SkuRecipe
from .po_extractor import extract_po_from_pdf


PLANNING_STATUSES = [
    ('draft', 'Draft'),
    ('reviewed', 'Reviewed'),
    ('approved', 'Approved'),
    ('closed', 'Closed'),
]
PLANNING_STATUS_SET = {value for value, _ in PLANNING_STATUSES}


def _clean_number(raw_value):
    if raw_value is None:
        return None
    text = str(raw_value).strip().replace(',', '')
    if not text:
        return None
    return text


def _to_int(raw_value):
    cleaned = _clean_number(raw_value)
    if cleaned is None:
        return None
    try:
        return int(float(cleaned))
    except ValueError:
        return None


def _to_decimal(raw_value):
    cleaned = _clean_number(raw_value)
    if cleaned is None:
        return None
    try:
        return Decimal(cleaned)
    except (InvalidOperation, ValueError):
        return None


def _to_date(raw_value):
    if not raw_value:
        return None
    text = str(raw_value).strip()
    if not text:
        return None

    for fmt in ('%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _parse_iso_date(raw_value):
    if not raw_value:
        return None
    try:
        return datetime.strptime(str(raw_value).strip(), '%Y-%m-%d').date()
    except ValueError:
        return None


def _normalize_status(raw_value, default='draft'):
    value = (raw_value or '').strip().lower()
    if value in {'open', 'pending'}:
        return 'draft'
    if value in PLANNING_STATUS_SET:
        return value
    return default


def _parse_date_filter(raw_value):
    value = (raw_value or '').strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, '%Y-%m-%d').date()
    except ValueError:
        return None


def _build_qr_image_base64(data):
    if not data:
        return ''
    try:
        import qrcode
    except ImportError:
        return ''

    qr = qrcode.QRCode(border=2, box_size=3)
    qr.add_data(data)
    qr.make(fit=True)
    image = qr.make_image(fill_color='black', back_color='white')
    buffer = io.BytesIO()
    image.save(buffer, format='PNG')
    return base64.b64encode(buffer.getvalue()).decode('ascii')


def _sku_key(sku):
    return (sku or '').strip().upper()


def _build_recipe_map(items):
    sku_values = sorted({_sku_key(item.get('sku')) for item in items if item.get('sku')})
    if not sku_values:
        return {}

    recipe_query = Q()
    for sku in sku_values:
        recipe_query |= Q(sku__iexact=sku)

    recipes = SkuRecipe.objects.filter(recipe_query)
    return {recipe.sku.upper(): recipe for recipe in recipes}


def _annotate_items_with_recipe(items, recipe_map):
    annotated = []
    repeat_count = 0
    new_count = 0
    missing_skus = []

    for item in items:
        sku = (item.get('sku') or '').strip()
        key = _sku_key(sku)
        has_recipe = bool(key and key in recipe_map)
        item_copy = dict(item)
        item_copy['is_repeat'] = has_recipe
        item_copy['recipe_status'] = 'Repeat' if has_recipe else 'New'
        annotated.append(item_copy)

        if has_recipe:
            repeat_count += 1
        else:
            new_count += 1
            if sku:
                missing_skus.append(sku)

    return annotated, repeat_count, new_count, sorted(set(missing_skus))


@login_required
@permission_required('can_edit_jobcard')
def planning_home(request):
    queryset = PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs').all()

    if request.method == 'POST' and request.POST.get('action') == 'bulk_update_status':
        selected_ids = []
        for raw_id in request.POST.getlist('selected_ids'):
            try:
                selected_ids.append(int(raw_id))
            except (TypeError, ValueError):
                continue

        target_status = _normalize_status(request.POST.get('target_status'), default='')
        if target_status not in PLANNING_STATUS_SET:
            messages.error(request, 'Please select a valid target status for bulk update.')
            return redirect('planning:home')

        if not selected_ids:
            messages.error(request, 'Select at least one planning row for bulk update.')
            return redirect('planning:home')

        updated = 0
        skipped_locked = 0
        for job in PlanningJob.objects.filter(id__in=selected_ids):
            current_status = _normalize_status(job.status)
            if current_status == 'approved' and target_status not in {'approved', 'reviewed'}:
                skipped_locked += 1
                continue
            if current_status == target_status:
                continue

            job.status = target_status
            if target_status == 'approved':
                job.issued_to_production = True
            elif current_status == 'approved' and target_status == 'reviewed':
                job.issued_to_production = False
            job.save(update_fields=['status', 'issued_to_production', 'updated_at'])
            updated += 1

        messages.success(
            request,
            f'Bulk status update complete. Updated {updated}, locked-skip {skipped_locked}.',
        )
        return redirect('planning:home')

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
        queryset = queryset.filter(status__iexact=status_filter)
    if department_filter:
        queryset = queryset.filter(department__icontains=department_filter)
    if machine_filter:
        queryset = queryset.filter(machine_name__icontains=machine_filter)
    if from_date:
        queryset = queryset.filter(plan_date__gte=from_date)
    if to_date:
        queryset = queryset.filter(plan_date__lte=to_date)

    status_rows = (
        queryset.values('status')
        .annotate(total=Count('id'))
        .order_by('status')
    )
    status_counts = {
        _normalize_status(row['status']): row['total']
        for row in status_rows
    }

    paginator = Paginator(queryset, 50)
    page_number = request.GET.get('page')
    jobs = paginator.get_page(page_number)
    return render(
        request,
        'planning/planning_home.html',
        {
            'jobs': jobs,
            'status_counts': status_counts,
            'status_choices': PLANNING_STATUSES,
            'filters': {
                'q': q,
                'status': status_filter,
                'department': department_filter,
                'machine': machine_filter,
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
                'ups': _to_int(row.get('Ups')),
                'print_sheet_size': (row.get('Print Sheet Size') or '').strip(),
                'print_sheets': _to_int(row.get('Print Sheets')),
                'wastage_sheets': _to_int(row.get('Wastage')),
                'actual_sheet_required': _to_int(row.get('Actual Sheet require')),
                'purchase_sheet_size': (row.get('Purchase Sheet Size') or '').strip(),
                'purchase_sheet_ups': _to_int(row.get('Purchase Sheet ups')),
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
                'purchase_material': (row.get('Purchase Material') or '').strip(),
                'stock_qty': _to_decimal(row.get('Stock')),
                'daily_demand': _to_decimal(row.get('Daily Demand')),
                'department': (row.get('Department') or '').strip(),
                'plate_set_no': (row.get('Plate Set No') or '').strip(),
                'awc_no': (row.get('AWC No.') or '').strip(),
                'aging_days': _to_int(row.get('Aging')),
                'die_cutting': (row.get('Die cutting') or '').strip(),
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
        return redirect('planning:home')

    return render(request, 'planning/planning_import.html')


@login_required
@permission_required('can_edit_jobcard')
def planning_job_detail(request, job_id):
    job = get_object_or_404(
        PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs'),
        id=job_id,
    )
    return render(
        request,
        'planning/planning_job_detail.html',
        {
            'job': job,
            'status_now': _normalize_status(job.status),
        },
    )


@login_required
@permission_required('can_edit_jobcard')
def planning_job_edit(request, job_id):
    job = get_object_or_404(PlanningJob, id=job_id)
    current_status = _normalize_status(job.status)

    if current_status == 'approved':
        messages.error(request, 'Approved records are locked. Unlock to Reviewed before editing.')
        return redirect('planning:job_detail', job_id=job.id)

    if request.method == 'POST':
        form = PlanningJobEditForm(request.POST, instance=job)
        if form.is_valid():
            edited = form.save(commit=False)
            edited.status = _normalize_status(edited.status)
            edited.job_card_version = (job.job_card_version or 1) + 1
            edited.save()
            messages.success(request, f'Planning job {edited.jc_number} updated.')
            return redirect('planning:job_detail', job_id=edited.id)
    else:
        form = PlanningJobEditForm(instance=job)

    return render(request, 'planning/planning_job_edit.html', {'job': job, 'form': form})


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def planning_job_status_update(request, job_id):
    if request.method != 'POST':
        return redirect('planning:job_detail', job_id=job_id)

    job = get_object_or_404(PlanningJob, id=job_id)
    current_status = _normalize_status(job.status)
    transition = (request.POST.get('transition') or '').strip()

    transitions = {
        'submit_review': ('draft', 'reviewed'),
        'approve': ('reviewed', 'approved'),
        'unlock': ('approved', 'reviewed'),
        'mark_closed': (None, 'closed'),
        'reopen': ('closed', 'draft'),
    }
    if transition not in transitions:
        messages.error(request, 'Unknown status transition request.')
        return redirect('planning:job_detail', job_id=job.id)

    required_from, target_status = transitions[transition]
    if required_from and current_status != required_from:
        messages.error(request, f'Transition not allowed from {current_status} to {target_status}.')
        return redirect('planning:job_detail', job_id=job.id)

    if current_status == target_status:
        messages.info(request, f'Job already in {target_status} status.')
        return redirect('planning:job_detail', job_id=job.id)

    job.status = target_status
    if target_status == 'approved':
        job.issued_to_production = True
    if transition == 'unlock':
        job.issued_to_production = False
    job.save(update_fields=['status', 'issued_to_production', 'updated_at'])

    messages.success(request, f'Job status updated: {current_status} -> {target_status}.')
    return redirect('planning:job_detail', job_id=job.id)


@login_required
@permission_required('can_edit_jobcard')
def planning_job_card_print(request, job_id):
    job = get_object_or_404(
        PlanningJob.objects.prefetch_related('print_runs', 'dispatch_runs'),
        id=job_id,
    )
    scan_url = request.build_absolute_uri(reverse('planning:scan_open', args=[job.jc_number]))
    context = {
        'job': job,
        'now_ts': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'status_now': _normalize_status(job.status),
        'scan_url': scan_url,
        'qr_code_b64': _build_qr_image_base64(scan_url),
    }
    return render(request, 'planning/planning_job_card_print.html', context)


@login_required
@permission_required('can_edit_jobcard')
def planning_report(request):
    queryset = PlanningJob.objects.all()

    from_date = _parse_date_filter(request.GET.get('from_date'))
    to_date = _parse_date_filter(request.GET.get('to_date'))
    if from_date:
        queryset = queryset.filter(plan_date__gte=from_date)
    if to_date:
        queryset = queryset.filter(plan_date__lte=to_date)

    totals = queryset.aggregate(
        total_jobs=Count('id'),
        total_order_qty=Sum('order_qty'),
        approved_jobs=Count('id', filter=Q(status__iexact='approved')),
        closed_jobs=Count('id', filter=Q(status__iexact='closed')),
    )

    by_status = (
        queryset.values('status')
        .annotate(total=Count('id'), order_qty=Sum('order_qty'))
        .order_by('status')
    )
    by_department = (
        queryset.values('department')
        .annotate(total=Count('id'), order_qty=Sum('order_qty'))
        .order_by('-total', 'department')[:20]
    )
    by_machine = (
        queryset.values('machine_name')
        .annotate(total=Count('id'), order_qty=Sum('order_qty'))
        .order_by('-total', 'machine_name')[:20]
    )

    context = {
        'totals': totals,
        'by_status': by_status,
        'by_department': by_department,
        'by_machine': by_machine,
        'filters': {
            'from_date': request.GET.get('from_date', ''),
            'to_date': request.GET.get('to_date', ''),
        },
    }
    return render(request, 'planning/planning_report.html', context)


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
@permission_required('can_edit_jobcard')
def upload_po(request):
    """Upload a PO PDF, extract its content, store it, and redirect to review."""
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

        # Reset file pointer so Django can save it
        pdf_file.seek(0)

        po_doc = PoDocument.objects.create(
            po_file=pdf_file,
            extracted_payload=extracted,
            extraction_status='processed',
            uploaded_by=request.user,
        )
        messages.success(
            request,
            f"PO {extracted.get('po_number', '?')} extracted with "
            f"{len(extracted.get('items', []))} line items. Review and confirm below.",
        )
        return redirect('planning:po_review', doc_id=po_doc.id)

    return render(request, 'planning/po_upload.html')


@login_required
@permission_required('can_edit_jobcard')
def po_review(request, doc_id):
    """Review extracted PO data and create PlanningJob records."""
    po_doc = get_object_or_404(PoDocument, id=doc_id)
    payload = po_doc.extracted_payload or {}
    items = payload.get('items', [])
    recipe_map = _build_recipe_map(items)
    annotated_items, repeat_count, new_count, missing_skus = _annotate_items_with_recipe(items, recipe_map)

    if request.method == 'POST':
        action = request.POST.get('action', '')
        if action == 'create_jobs':
            with transaction.atomic():
                created_count = 0
                updated_count = 0
                skipped_count = 0
                locked_count = 0
                missing_recipe_count = 0
                po_number = payload.get('po_number', '')
                po_date_raw = payload.get('po_date')
                po_date = _parse_iso_date(po_date_raw)
                delivery_location = payload.get('delivery_location', '')
                department = payload.get('department', '')

                for item in annotated_items:
                    sku = (item.get('sku') or '').strip()
                    job_name = (item.get('job_name') or '').strip() or sku
                    if not sku:
                        skipped_count += 1
                        continue

                    recipe = recipe_map.get(_sku_key(sku))
                    if not recipe:
                        missing_recipe_count += 1
                        continue

                    field_prefix = f"item_{item['line_no']}_"
                    skip_flag = request.POST.get(f"{field_prefix}skip") == '1'

                    if skip_flag:
                        skipped_count += 1
                        continue

                    delivery_date = _parse_iso_date(item.get('delivery_date'))
                    plan_date = delivery_date or po_date

                    # Re-use existing PO+SKU job to avoid duplicates; otherwise allocate a new serial.
                    existing_job = PlanningJob.objects.filter(
                        po_number=po_number,
                        sku=sku,
                    ).order_by('-updated_at', '-id').first()

                    if existing_job:
                        if _normalize_status(existing_job.status) == 'approved':
                            locked_count += 1
                            continue
                        jc_number = existing_job.jc_number
                    else:
                        jc_number = allocate_next_jc_number(plan_date)

                    qty = item.get('quantity')
                    order_qty = int(qty) if qty is not None else None

                    unit_cost_val = item.get('unit_cost')
                    unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None

                    defaults = {
                        'po_number': po_number,
                        'sku': sku,
                        'job_name': recipe.job_name or job_name,
                        'order_qty': order_qty,
                        'department': recipe.department or department,
                        'destination': delivery_location,
                        'unit_cost': unit_cost_dec if unit_cost_dec is not None else recipe.default_unit_cost,
                        'status': 'draft',
                        'repeat_flag': 'Repeat',
                        'material': recipe.material,
                        'color_spec': recipe.color_spec,
                        'application': recipe.application,
                        'size_w_mm': recipe.size_w_mm,
                        'size_h_mm': recipe.size_h_mm,
                        'ups': recipe.ups,
                        'print_sheet_size': recipe.print_sheet_size,
                        'purchase_sheet_size': recipe.purchase_sheet_size,
                        'purchase_material': recipe.purchase_material,
                        'machine_name': recipe.machine_name,
                    }
                    if plan_date:
                        defaults['plan_date'] = plan_date

                    _, created = PlanningJob.objects.update_or_create(
                        jc_number=jc_number,
                        defaults=defaults,
                    )
                    if created:
                        created_count += 1
                    else:
                        updated_count += 1

            po_doc.extraction_status = 'processed'
            po_doc.save(update_fields=['extraction_status'])

            messages.success(
                request,
                f'Done. Created {created_count}, updated {updated_count}, skipped {skipped_count}, missing-recipe {missing_recipe_count}, locked-skip {locked_count} planning job(s).',
            )
            return redirect('planning:home')

    context = {
        'po_doc': po_doc,
        'payload': payload,
        'items': annotated_items,
        'items_json': json.dumps(annotated_items),
        'repeat_count': repeat_count,
        'new_count': new_count,
        'missing_skus': missing_skus,
    }
    return render(request, 'planning/po_review.html', context)


@login_required
@permission_required('can_edit_jobcard')
@transaction.atomic
def po_new_skus(request, doc_id):
    po_doc = get_object_or_404(PoDocument, id=doc_id)
    payload = po_doc.extracted_payload or {}
    items = payload.get('items', [])

    recipe_map = _build_recipe_map(items)
    _, _, _, missing_skus = _annotate_items_with_recipe(items, recipe_map)

    if request.method == 'POST':
        created_count = 0
        for sku in missing_skus:
            prefix = f"sku_{sku}"
            job_name = (request.POST.get(f"{prefix}_job_name") or '').strip()
            material = (request.POST.get(f"{prefix}_material") or '').strip()
            color_spec = (request.POST.get(f"{prefix}_color_spec") or '').strip()
            application = (request.POST.get(f"{prefix}_application") or '').strip()
            machine_name = (request.POST.get(f"{prefix}_machine_name") or '').strip()
            department = (request.POST.get(f"{prefix}_department") or '').strip()
            print_sheet_size = (request.POST.get(f"{prefix}_print_sheet_size") or '').strip()
            purchase_material = (request.POST.get(f"{prefix}_purchase_material") or '').strip()

            unit_cost_raw = (request.POST.get(f"{prefix}_default_unit_cost") or '').strip()
            unit_cost = None
            if unit_cost_raw:
                try:
                    unit_cost = Decimal(unit_cost_raw)
                except InvalidOperation:
                    unit_cost = None

            if not job_name and not material and not machine_name:
                # Keep save requirements simple, but avoid empty recipe rows.
                continue

            SkuRecipe.objects.update_or_create(
                sku=sku,
                defaults={
                    'job_name': job_name,
                    'material': material,
                    'color_spec': color_spec,
                    'application': application,
                    'machine_name': machine_name,
                    'department': department,
                    'print_sheet_size': print_sheet_size,
                    'purchase_material': purchase_material,
                    'default_unit_cost': unit_cost,
                    'created_by': request.user,
                },
            )
            created_count += 1

        messages.success(request, f'SKU recipes saved: {created_count}. You can continue PO review now.')
        return redirect('planning:po_review', doc_id=po_doc.id)

    return render(
        request,
        'planning/po_new_skus.html',
        {
            'po_doc': po_doc,
            'payload': payload,
            'missing_skus': missing_skus,
            'example_form': SkuRecipeForm(),
        },
    )
