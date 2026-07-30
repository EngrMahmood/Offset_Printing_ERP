import logging

from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.http import Http404, HttpResponseForbidden
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from reports.report_registry import registry

logger = logging.getLogger(__name__)

# Ensure built-in reports are loaded.
from reports.report_registry import builtin_reports  # noqa: F401


def reports_home(request):
    return render(
        request,
        'reports/index.html',
        {
            'reports': registry.all(),
        },
    )


def report_detail(request, report_type: str):
    from django.template import TemplateDoesNotExist
    from django.template.loader import get_template
    from reports.report_engine.engine import run_report
    
    report = registry.get(report_type)
    if report is None:
        raise Http404('Report not found')
        
    template_name = f'reports/{report_type}.html'
    try:
        get_template(template_name)
        has_custom = True
    except TemplateDoesNotExist:
        template_name = 'reports/report_detail.html'
        has_custom = False
    
    context = {
        'report': report,
    }
    if has_custom:
        try:
            payload = run_report(report_type, request)
            context['data'] = payload.get('data') or {}
        except PermissionDenied:
            return HttpResponseForbidden("You don't have access to this report.")
        except Exception:
            logger.exception('Failed to build report %s', report_type)

    return render(
        request,
        template_name,
        context,
    )


@require_POST
def kpi_scorecard_save_note(request):
    from reports.kpi_services import KPI_COMPUTE_FUNCS
    from reports.models import KPIActionNote
    from reports.report_engine.engine import _has_access, bump_cache_version

    report = registry.get('kpi-scorecard')
    if report is None or not _has_access(request, report.permissions):
        return HttpResponseForbidden("You don't have access to this report.")

    kpi_slug = (request.POST.get('kpi_slug') or '').strip()
    period_type = (request.POST.get('period_type') or '').strip()
    period_key = (request.POST.get('period_key') or '').strip()
    status = (request.POST.get('status') or '').strip()
    note_text = request.POST.get('note') or ''

    if kpi_slug in KPI_COMPUTE_FUNCS and period_type in ('month', 'quarter') and period_key:
        KPIActionNote.objects.update_or_create(
            kpi_slug=kpi_slug,
            period_type=period_type,
            period_key=period_key,
            defaults={
                'note': note_text,
                'status': status,
                'updated_by': request.user if request.user.is_authenticated else None,
            },
        )
        bump_cache_version()
        messages.success(request, 'Action plan saved.')
    else:
        messages.error(request, 'Could not save action plan — invalid KPI or period.')

    redirect_url = reverse('reports:detail', args=['kpi-scorecard'])
    query = request.POST.get('return_query') or ''
    if query:
        redirect_url = f'{redirect_url}?{query}'
    return redirect(redirect_url)
