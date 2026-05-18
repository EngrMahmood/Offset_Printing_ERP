from django.http import Http404
from django.shortcuts import render

from .services import build_overview_context, build_report_context


def reports_home(request):
    context = build_overview_context(request)
    return render(request, 'reports/index.html', context)


def report_detail(request, report_type: str):
    context = build_report_context(report_type, request)
    if context is None:
        raise Http404('Report not found')
    return render(request, 'reports/report_detail.html', context)
