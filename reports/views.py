from django.http import Http404
from django.shortcuts import render

from reports.report_registry import registry

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
        except Exception:
            pass

    return render(
        request,
        template_name,
        context,
    )
