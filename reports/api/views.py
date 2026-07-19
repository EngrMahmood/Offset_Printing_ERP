from __future__ import annotations

from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
from django.http import HttpResponse, JsonResponse
from django.views.decorators.http import require_GET
from django.views.decorators.http import require_POST

# Ensure built-in reports are registered on import.
from reports.report_registry import builtin_reports  # noqa: F401
from reports.report_engine import run_report
from reports.report_engine.engine import bump_cache_version
from reports.report_registry import registry
from reports.export.services import export_as_csv, export_as_pdf, export_as_xlsx
from reports.models import MachinePlanningJcSelection, ScheduledReport
from reports.scheduler.services import calculate_next_run
from reports.services import _can_edit_jc_selection


@require_GET
def list_reports(request):
    items = []
    for report in registry.all():
        items.append(
            {
                'slug': report.slug,
                'title': report.title,
                'description': report.description,
                'department': report.department,
                'category': report.category,
                'navigation_group': report.navigation_group,
                'icon': report.icon,
                'filters': list(report.filters),
                'supported_exports': list(report.supported_exports),
                'supported_charts': list(report.supported_charts),
                'drilldown_support': report.drilldown_support,
            }
        )
    return JsonResponse({'ok': True, 'count': len(items), 'items': items})


@require_GET
def run_report_api(request, slug):
    try:
        payload = run_report(slug, request)
    except KeyError:
        return JsonResponse({'ok': False, 'error': 'Report not found'}, status=404)
    except PermissionDenied as exc:
        return JsonResponse({'ok': False, 'error': str(exc)}, status=403)
    return JsonResponse({'ok': True, 'payload': payload})


@require_GET
def export_report_api(request, slug):
    export_type = (request.GET.get('type') or 'csv').strip().lower()
    
    # Inject _export=true to bypass pagination/limits in the executor
    q_params = request.GET.copy()
    q_params['_export'] = 'true'
    request.GET = q_params

    try:
        payload = run_report(slug, request)
    except KeyError:
        return JsonResponse({'ok': False, 'error': 'Report not found'}, status=404)
    except PermissionDenied as exc:
        return JsonResponse({'ok': False, 'error': str(exc)}, status=403)

    filename_base = slug.replace('_', '-').strip() or 'report'
    if export_type == 'xlsx':
        try:
            content = export_as_xlsx(payload)
        except RuntimeError as exc:
            return JsonResponse({'ok': False, 'error': str(exc)}, status=400)
        response = HttpResponse(
            content,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        response['Content-Disposition'] = f'attachment; filename="{filename_base}.xlsx"'
        return response

    if export_type == 'pdf':
        try:
            content = export_as_pdf(payload)
        except RuntimeError as exc:
            return JsonResponse({'ok': False, 'error': str(exc)}, status=400)
        response = HttpResponse(content, content_type='application/pdf')
        response['Content-Disposition'] = f'attachment; filename="{filename_base}.pdf"'
        return response

    content = export_as_csv(payload)
    response = HttpResponse(content, content_type='text/csv; charset=utf-8')
    response['Content-Disposition'] = f'attachment; filename="{filename_base}.csv"'
    return response


@login_required
@require_POST
def machine_planning_jc_selection_api(request, jc_number):
    """V2 plan item 3: toggle whether a JC is included in its combined
    Machine Planning run. Shared state (not per-user); only Planner/Manager/
    Admin (or superuser) may change it."""
    if not _can_edit_jc_selection(request.user):
        return JsonResponse({'ok': False, 'error': 'Only Planner, Manager, or Admin can change JC selection.'}, status=403)

    is_excluded_raw = (request.POST.get('is_excluded') or '').strip().lower()
    is_excluded = is_excluded_raw in {'1', 'true', 'yes', 'on'}

    selection, _ = MachinePlanningJcSelection.objects.update_or_create(
        jc_number=jc_number,
        defaults={'is_excluded': is_excluded, 'updated_by': request.user},
    )
    bump_cache_version()
    return JsonResponse({'ok': True, 'jc_number': jc_number, 'is_excluded': selection.is_excluded})


@require_GET
def list_schedules_api(request):
    qs = ScheduledReport.objects.filter(is_active=True).order_by('-created_at')
    items = [
        {
            'id': row.id,
            'name': row.name,
            'report_slug': row.report_slug,
            'frequency': row.frequency,
            'recipients': row.recipients,
            'filters': row.filters,
            'next_run_at': row.next_run_at.isoformat() if row.next_run_at else None,
            'last_run_at': row.last_run_at.isoformat() if row.last_run_at else None,
        }
        for row in qs
    ]
    return JsonResponse({'ok': True, 'count': len(items), 'items': items})


@require_POST
def create_schedule_api(request):
    report_slug = (request.POST.get('report_slug') or '').strip().lower()
    if not registry.get(report_slug):
        return JsonResponse({'ok': False, 'error': 'Invalid report slug'}, status=400)

    frequency = (request.POST.get('frequency') or ScheduledReport.FREQ_WEEKLY).strip().lower()
    if frequency not in {choice[0] for choice in ScheduledReport.FREQUENCY_CHOICES}:
        return JsonResponse({'ok': False, 'error': 'Invalid frequency'}, status=400)

    schedule = ScheduledReport.objects.create(
        name=(request.POST.get('name') or f'{report_slug} schedule').strip(),
        report_slug=report_slug,
        frequency=frequency,
        recipients=(request.POST.get('recipients') or '').strip(),
        filters={},
        next_run_at=calculate_next_run(frequency),
        created_by=request.user if getattr(request, 'user', None) and request.user.is_authenticated else None,
    )
    return JsonResponse({'ok': True, 'id': schedule.id})
