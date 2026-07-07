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
    report = registry.get(report_type)
    if report is None:
        raise Http404('Report not found')
    return render(
        request,
        'reports/report_detail.html',
        {
            'report': report,
        },
    )
