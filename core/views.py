from django.shortcuts import render, get_object_or_404, redirect
from django.http import HttpResponse, JsonResponse
from django.contrib import messages
from django.utils import timezone
from django.views.decorators.http import require_POST
from django.db.models import Q
import csv
import json
from datetime import datetime

from .bulk_upload import process_jobcard_upload, get_template_headers, get_template_example
from .models import JobCard, Production, Machine, Operator, Department, Material, Dispatch

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


def home(request):
    return render(request, 'home.html')


def bulk_upload_jobcards(request):
    context = {}

    if request.method == "POST":
        file = request.FILES.get('file')

        if not file:
            context = {
                'success_count': 0,
                'error_count': 1,
                'errors': [{'row': 0, 'errors': 'Please choose a file to upload.'}],
            }
            return render(request, "upload.html", context)

        result = process_jobcard_upload(file)
        context = result

        return render(request, "upload.html", context)

    return render(request, "upload.html", context)


@require_POST
def quick_add_master(request):
    """Create master dropdown values for planner workflow without admin dependency."""
    master_type = (request.POST.get('type') or '').strip().lower()
    name = (request.POST.get('name') or '').strip()

    if master_type not in {'material', 'machine', 'department'}:
        return JsonResponse({'ok': False, 'error': 'Invalid master type.'}, status=400)

    if not name:
        return JsonResponse({'ok': False, 'error': 'Name is required.'}, status=400)

    model_map = {
        'material': Material,
        'machine': Machine,
        'department': Department,
    }
    model = model_map[master_type]

    # Case-insensitive duplicate check for better UX.
    existing = model.objects.filter(name__iexact=name).first()
    if existing:
        return JsonResponse({
            'ok': True,
            'created': False,
            'id': existing.id,
            'name': existing.name,
            'type': master_type,
            'message': 'Already exists. Selected existing value.'
        })

    obj = model.objects.create(name=name)
    return JsonResponse({
        'ok': True,
        'created': True,
        'id': obj.id,
        'name': obj.name,
        'type': master_type,
        'message': 'Created successfully.'
    })


def job_card_entry(request):
    """Manual job card entry form"""
    if request.method == "POST":
        try:
            job_card_no = (request.POST.get('job_card_no') or '').strip()
            sku = (request.POST.get('sku') or '').strip()
            order_qty = int(request.POST.get('order_qty') or 0)

            if not job_card_no:
                raise ValueError("Job card number is required")
            if not sku:
                raise ValueError("SKU is required")
            if order_qty <= 0:
                raise ValueError("Order quantity must be greater than 0")

            if JobCard.objects.filter(job_card_no=job_card_no).exists():
                raise ValueError(f"Job card number {job_card_no} already exists")

            po_date_raw = (request.POST.get('po_date') or '').strip()
            po_date = datetime.strptime(po_date_raw, "%Y-%m-%d").date() if po_date_raw else None
            month_value = po_date.strftime("%B") if po_date else ((request.POST.get('month') or '').strip() or None)

            material = Material.objects.filter(id=request.POST.get('material')).first() if request.POST.get('material') else None
            machine = Machine.objects.filter(id=request.POST.get('machine_name')).first() if request.POST.get('machine_name') else None
            department = Department.objects.filter(id=request.POST.get('department')).first() if request.POST.get('department') else None

            JobCard.objects.create(
                job_card_no=job_card_no,
                month=month_value,
                po_date=po_date,
                PO_No=(request.POST.get('po_no') or '').strip() or None,
                SKU=sku,
                material=material,
                colour=int(request.POST.get('colour') or 0) or None,
                application=(request.POST.get('application') or '').strip() or None,
                order_qty=order_qty,
                total_impressions_required=int(request.POST.get('total_impressions_required') or 0) or None,
                ups=int(request.POST.get('ups') or 0) or None,
                print_sheet_size=(request.POST.get('print_sheet_size') or '').strip() or None,
                wastage=int(request.POST.get('wastage') or 0),
                purchase_sheet_size=(request.POST.get('purchase_sheet_size') or '').strip() or None,
                purchase_sheet_ups=int(request.POST.get('purchase_sheet_ups') or 0) or None,
                remarks=(request.POST.get('remarks') or '').strip() or None,
                destination=(request.POST.get('destination') or '').strip() or None,
                machine_name=machine,
                department=department,
                die_cutting=(request.POST.get('die_cutting') or '').strip() or None,
                status=(request.POST.get('status') or 'Open').strip() or 'Open',
            )

            messages.success(request, f"Job card {job_card_no} created successfully")
            return redirect('job_card_entry')

        except Exception as e:
            messages.error(request, f"Error creating job card: {str(e)}")

    context = {
        'today': timezone.now().date(),
        'materials': Material.objects.all().order_by('name'),
        'machines': Machine.objects.filter(is_active=True).order_by('name'),
        'departments': Department.objects.all().order_by('name'),
    }
    return render(request, 'job_card_entry.html', context)


def job_card_records(request):
    """Job card records list page"""
    query = (request.GET.get('q') or '').strip()
    status = (request.GET.get('status') or '').strip()

    jobcards = JobCard.objects.select_related('material', 'machine_name', 'department').order_by('-created_at')

    if query:
        jobcards = jobcards.filter(
            Q(job_card_no__icontains=query) |
            Q(SKU__icontains=query) |
            Q(PO_No__icontains=query)
        )

    if status:
        jobcards = jobcards.filter(status=status)

    context = {
        'jobcards': jobcards,
        'q': query,
        'status': status,
    }
    return render(request, 'job_card_records.html', context)


def production_entry(request):
    """Production data entry form for operators"""
    if request.method == "POST":
        job_card_id = request.POST.get('job_card')
        machine_id = request.POST.get('machine')
        operator_id = request.POST.get('operator')
        shift = request.POST.get('shift')
        date = request.POST.get('date')
        impressions = request.POST.get('impressions')
        output_sheets = request.POST.get('output_sheets')
        waste_sheets = request.POST.get('waste_sheets')
        downtime_minutes = request.POST.get('downtime_minutes')
        planned_time = request.POST.get('planned_time')
        run_time = request.POST.get('run_time')
        setup_time = request.POST.get('setup_time')
        downtime_category = request.POST.get('downtime_category')
        waste_reason = request.POST.get('waste_reason')
        remarks = request.POST.get('remarks')

        try:
            output_sheets_val = int(output_sheets) if output_sheets else 0
            waste_sheets_val = int(waste_sheets) if waste_sheets else 0
            downtime_val = float(downtime_minutes) if downtime_minutes else 0
            planned_time_val = float(planned_time) if planned_time else 0
            run_time_val = float(run_time) if run_time else 0
            setup_time_val = float(setup_time) if setup_time else 0

            if planned_time_val <= 0:
                raise ValueError("Planned time must be greater than 0")
            if run_time_val <= 0:
                raise ValueError("Run time must be greater than 0")
            if downtime_val > 0 and not downtime_category:
                raise ValueError("Downtime category is required when downtime is greater than 0")
            if waste_sheets_val > 0 and not waste_reason:
                raise ValueError("Waste reason is required when waste sheets are greater than 0")
            if (run_time_val + downtime_val + setup_time_val) > planned_time_val:
                raise ValueError("Run time + downtime + setup time cannot exceed planned time")

            job_card = get_object_or_404(JobCard, pk=job_card_id)
            machine = get_object_or_404(Machine, pk=machine_id)
            operator = get_object_or_404(Operator, pk=operator_id)

            # Create production record
            Production.objects.create(
                job_card=job_card,
                machine=machine,
                operator=operator,
                shift=shift,
                date=date,
                impressions=int(impressions) if impressions else 0,
                output_sheets=output_sheets_val,
                waste_sheets=waste_sheets_val,
                planned_time=planned_time_val,
                run_time=run_time_val,
                setup_time=setup_time_val,
                downtime=downtime_val,
                downtime_category=downtime_category,
                waste_reason=waste_reason
            )

            messages.success(request, f'Production data saved successfully for Job Card {job_card.job_card_no}')
            return redirect('production_entry')

        except Exception as e:
            messages.error(request, f'Error saving production data: {str(e)}')

    # Get data for form dropdowns
    job_cards = JobCard.objects.filter(status__in=['Open', 'In Progress']).order_by('-created_at')
    machines = Machine.objects.filter(is_active=True)
    operators = Operator.objects.all()

    context = {
        'job_cards': job_cards,
        'machines': machines,
        'operators': operators,
        'today': timezone.now().date(),
    }

    return render(request, 'production_entry.html', context)


def production_records(request):
    """Production records list page"""
    query = (request.GET.get('q') or '').strip()
    shift = (request.GET.get('shift') or '').strip()

    records = Production.objects.select_related('job_card', 'machine', 'operator').order_by('-date', '-id')

    if query:
        records = records.filter(
            Q(job_card__job_card_no__icontains=query) |
            Q(machine__name__icontains=query) |
            Q(operator__name__icontains=query)
        )

    if shift:
        records = records.filter(shift=shift)

    context = {
        'records': records,
        'q': query,
        'shift': shift,
    }
    return render(request, 'production_records.html', context)


def dispatch_entry(request):
    """Dispatch entry form"""
    if request.method == 'POST':
        try:
            job_card_id = request.POST.get('job_card')
            dc_no = (request.POST.get('dc_no') or '').strip() or None
            dispatch_date_raw = request.POST.get('dispatch_date')
            dispatch_qty = int(request.POST.get('dispatch_qty') or 0)

            if not job_card_id:
                raise ValueError('Job card is required')
            if dispatch_qty <= 0:
                raise ValueError('Dispatch quantity must be greater than 0')

            job_card = get_object_or_404(JobCard, pk=job_card_id)
            dispatch_date = datetime.strptime(dispatch_date_raw, "%Y-%m-%d").date() if dispatch_date_raw else timezone.now().date()

            Dispatch.objects.create(
                job_card=job_card,
                dc_no=dc_no,
                dispatch_date=dispatch_date,
                dispatch_qty=dispatch_qty,
            )

            messages.success(request, f'Dispatch saved for {job_card.job_card_no}')
            return redirect('dispatch_entry')
        except Exception as e:
            messages.error(request, f'Error saving dispatch: {str(e)}')

    context = {
        'job_cards': JobCard.objects.order_by('-created_at')[:200],
        'today': timezone.now().date(),
    }
    return render(request, 'dispatch_entry.html', context)


def dispatch_records(request):
    """Dispatch records list page"""
    query = (request.GET.get('q') or '').strip()

    records = Dispatch.objects.select_related('job_card').order_by('-dispatch_date', '-id')
    if query:
        records = records.filter(
            Q(job_card__job_card_no__icontains=query) |
            Q(dc_no__icontains=query)
        )

    context = {
        'records': records,
        'q': query,
    }
    return render(request, 'dispatch_records.html', context)


def production_dashboard(request):
    """Real-time production dashboard with OEE metrics"""
    from django.db.models import Sum, Count
    from datetime import timedelta

    # Get date range (default to last 7 days)
    try:
        days = int(request.GET.get('days', 7))
    except (TypeError, ValueError):
        days = 7
    if days < 1:
        days = 1

    end_date = timezone.now().date()
    start_date = end_date - timedelta(days=days - 1)

    period_productions = Production.objects.filter(date__gte=start_date, date__lte=end_date)
    period_dispatches = Dispatch.objects.filter(dispatch_date__gte=start_date, dispatch_date__lte=end_date)

    # Calculate period metrics
    total_impressions = period_productions.aggregate(total=Sum('impressions'))['total'] or 0
    total_downtime = period_productions.aggregate(total=Sum('downtime'))['total'] or 0
    total_output = period_productions.aggregate(total=Sum('output_sheets'))['total'] or 0
    total_waste = period_productions.aggregate(total=Sum('waste_sheets'))['total'] or 0

    # Dispatch metrics for same period
    total_dispatch_qty = period_dispatches.aggregate(total=Sum('dispatch_qty'))['total'] or 0
    dispatch_count = period_dispatches.count()
    dispatched_job_cards_count = period_dispatches.values('job_card').distinct().count()
    avg_dispatch_qty = (total_dispatch_qty / dispatch_count) if dispatch_count else 0
    dispatch_fulfillment_pct = (total_dispatch_qty / total_output * 100) if total_output > 0 else 0

    # Calculate OEE (simplified - you may want to adjust based on your machine standards)
    available_time_minutes = period_productions.count() * 480  # Assuming 8 hours per production record
    actual_run_time = available_time_minutes - total_downtime

    availability = (actual_run_time / available_time_minutes * 100) if available_time_minutes > 0 else 0
    performance = 85.0  # Placeholder - would need machine standards
    quality = ((total_output / (total_output + total_waste)) * 100) if (total_output + total_waste) > 0 else 100

    oee_value = (availability * performance * quality) / 10000

    # Recent productions (last 10)
    recent_productions = period_productions.select_related('job_card', 'machine', 'operator').order_by('-date', '-id')[:10]
    recent_dispatches = period_dispatches.select_related('job_card').order_by('-dispatch_date', '-id')[:10]

    # Production summary by date
    production_by_date = period_productions\
        .values('date')\
        .annotate(
            total_impressions=Sum('impressions'),
            total_output=Sum('output_sheets'),
            total_waste=Sum('waste_sheets'),
            total_downtime=Sum('downtime'),
            count=Count('id')
        )\
        .order_by('-date')

    # Machine utilization
    machine_utilization = period_productions\
        .values('machine__name')\
        .annotate(
            total_impressions=Sum('impressions'),
            total_downtime=Sum('downtime'),
            production_count=Count('id')
        )\
        .order_by('-total_impressions')[:5]

    # Downtime analysis
    downtime_by_category = period_productions\
        .exclude(downtime_category__isnull=True)\
        .exclude(downtime_category='')\
        .values('downtime_category')\
        .annotate(
            total_minutes=Sum('downtime'),
            count=Count('id')
        )\
        .order_by('-total_minutes')[:5]

    # Dispatch trends and top dispatched job cards
    dispatch_by_date = period_dispatches\
        .values('dispatch_date')\
        .annotate(
            total_dispatch=Sum('dispatch_qty'),
            count=Count('id')
        )\
        .order_by('-dispatch_date')

    top_dispatched_jobcards = period_dispatches\
        .values('job_card__job_card_no')\
        .annotate(total_dispatch=Sum('dispatch_qty'))\
        .order_by('-total_dispatch')[:5]

    # Convert query data into JSON-serializable lists for charts.
    production_by_date_data = [
        {
            'date_only': item['date'].isoformat() if item['date'] else None,
            'total_impressions': float(item['total_impressions'] or 0),
            'total_output': float(item['total_output'] or 0),
            'total_waste': float(item['total_waste'] or 0),
            'total_downtime': float(item['total_downtime'] or 0),
            'count': int(item['count'] or 0),
        }
        for item in production_by_date
    ]

    machine_utilization_data = [
        {
            'machine_name': item['machine__name'] or 'Unknown',
            'total_impressions': float(item['total_impressions'] or 0),
            'total_downtime': float(item['total_downtime'] or 0),
            'production_count': int(item['production_count'] or 0),
        }
        for item in machine_utilization
    ]

    downtime_by_category_data = [
        {
            'downtime_category': item['downtime_category'],
            'total_minutes': float(item['total_minutes'] or 0),
            'count': int(item['count'] or 0),
        }
        for item in downtime_by_category
    ]

    dispatch_by_date_data = [
        {
            'date_only': item['dispatch_date'].isoformat() if item['dispatch_date'] else None,
            'total_dispatch': float(item['total_dispatch'] or 0),
            'count': int(item['count'] or 0),
        }
        for item in dispatch_by_date
    ]

    top_dispatched_jobcards_data = [
        {
            'job_card_no': item['job_card__job_card_no'] or 'Unknown',
            'total_dispatch': float(item['total_dispatch'] or 0),
        }
        for item in top_dispatched_jobcards
    ]

    context = {
        'today': end_date,
        'days': days,
        'period_label': 'Today' if days == 1 else f'Last {days} Days',
        'oee_value': round(oee_value, 2),
        'availability_value': round(availability, 2),
        'performance_value': round(performance, 2),
        'quality_value': round(quality, 2),
        'total_impressions': total_impressions,
        'total_output': total_output,
        'total_waste': total_waste,
        'total_downtime': total_downtime,
        'total_dispatch_qty': total_dispatch_qty,
        'dispatch_count': dispatch_count,
        'dispatched_job_cards_count': dispatched_job_cards_count,
        'avg_dispatch_qty': round(avg_dispatch_qty, 2),
        'dispatch_fulfillment_pct': round(dispatch_fulfillment_pct, 2),
        'recent_productions': recent_productions,
        'recent_dispatches': recent_dispatches,
        'production_by_date_json': json.dumps(production_by_date_data),
        'machine_utilization_json': json.dumps(machine_utilization_data),
        'downtime_by_category_json': json.dumps(downtime_by_category_data),
        'dispatch_by_date_json': json.dumps(dispatch_by_date_data),
        'top_dispatched_jobcards_json': json.dumps(top_dispatched_jobcards_data),
    }

    return render(request, 'production_dashboard.html', context)


def download_template(request):
    """Download template in CSV or Excel format"""
    file_format = request.GET.get('format', 'csv').lower()
    
    headers = get_template_headers()
    example = get_template_example()
    
    if file_format == 'excel' and EXCEL_AVAILABLE:
        try:
            # Generate Excel file
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.title = "Job Cards"
            
            # Add headers with styling
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            
            for col_num, header in enumerate(headers, 1):
                cell = worksheet.cell(row=1, column=col_num, value=header)
                cell.fill = header_fill
                cell.font = header_font
            
            # Add example row
            for col_num, value in enumerate(example, 1):
                worksheet.cell(row=2, column=col_num, value=value)
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Send file
            response = HttpResponse(
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename="jobcard_template.xlsx"'
            workbook.save(response)
            return response
        
        except Exception as e:
            # Fallback to CSV if Excel generation fails
            print(f"Excel generation failed: {e}")
            import traceback
            traceback.print_exc()
    
    # Generate CSV file (default/fallback)
    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="jobcard_template.csv"'
    
    writer = csv.writer(response)
    writer.writerow(headers)
    writer.writerow(example)
    
    return response