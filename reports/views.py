import logging

from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.http import Http404, HttpResponse, HttpResponseForbidden, JsonResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from reports.report_registry import registry

logger = logging.getLogger(__name__)

# Ensure built-in reports are loaded.
from reports.report_registry import builtin_reports  # noqa: F401


def sku_suggest(request):
    """Type-ahead for the SKU History search box: SKUs containing the typed text."""
    from reports.report_engine.engine import _has_access
    from reports.sku_history import suggest_skus

    report = registry.get('sku-history')
    if report is None or not _has_access(request, report.permissions):
        return JsonResponse({'results': []}, status=403)
    return JsonResponse({'results': suggest_skus(request.GET.get('q', ''))})


def _hourly_report_allowed(request):
    from reports.report_engine.engine import _has_access

    report = registry.get('hourly-report')
    return report is not None and _has_access(request, report.permissions)


def hourly_report_page(request):
    """Hourly Report: live snapshot of printing, packing, dispatch and planning, with a WhatsApp share button."""
    from django.utils import timezone

    from reports.hourly_report import build_hourly_report
    from reports.models import HourlyReportSnapshot

    if not request.user.is_authenticated:
        return redirect('login')
    if not _hourly_report_allowed(request):
        return HttpResponseForbidden("You don't have access to this report.")
    report = build_hourly_report()
    history = [
        {
            'time': timezone.localtime(s.taken_at).strftime('%H:%M'),
            'by': (s.shared_by.get_full_name() or s.shared_by.username) if s.shared_by_id else '',
            'impressions': (s.payload.get('totals') or {}).get('day_impressions', 0),
            'packed': (s.payload.get('totals') or {}).get('packed_pcs', 0),
            'dispatched': (s.payload.get('totals') or {}).get('dispatched_pcs', 0),
        }
        for s in HourlyReportSnapshot.objects.filter(business_date=report['business_date']).select_related('shared_by')[:12]
    ]
    return render(request, 'reports/hourly_report_page.html', {
        'report': registry.get('hourly-report'), 'data': report, 'history': history,
    })


def hourly_report_image(request):
    """PNG of the current Hourly Report, for sharing next to the text."""
    from reports.hourly_report import build_hourly_report, render_image

    if not request.user.is_authenticated:
        return redirect('login')
    if not _hourly_report_allowed(request):
        return HttpResponseForbidden("You don't have access to this report.")
    shared_by = (request.user.get_full_name() or request.user.username) if request.GET.get('by') else ''
    response = HttpResponse(render_image(build_hourly_report(), shared_by=shared_by), content_type='image/png')
    response['Cache-Control'] = 'no-store'
    response['Content-Disposition'] = 'inline; filename="hourly-report.png"'
    return response


@require_POST
def hourly_report_share(request):
    """Record that the supervisor is sharing the report now and hand back exactly what to send:
    the figures are rebuilt here, so what is logged is what goes out."""
    from reports.hourly_report import build_hourly_report, build_text
    from reports.models import HourlyReportSnapshot

    if not request.user.is_authenticated:
        return JsonResponse({'ok': False}, status=401)
    if not _hourly_report_allowed(request):
        return JsonResponse({'ok': False}, status=403)
    report = build_hourly_report()
    name = request.user.get_full_name() or request.user.username
    text = build_text(report, shared_by=name)
    snapshot = HourlyReportSnapshot.objects.create(
        business_date=report['business_date'], shift=report['shift'], shared_by=request.user,
        channel=(request.POST.get('channel') or 'whatsapp')[:20],
        payload={'totals': report['totals'], 'text': text},
    )
    return JsonResponse({
        'ok': True, 'snapshot_id': snapshot.id, 'text': text,
        'image_url': reverse('reports:hourly_report_image') + '?by=1',
    })


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
    if 'quarter' in report.filters:
        from django.utils import timezone

        this_year = timezone.localdate().year
        context['quarter_years'] = list(range(this_year, max(this_year - 4, 2019), -1))
    if 'product_type' in report.filters:
        from core.models import ProductType

        context['product_types'] = list(ProductType.objects.order_by('name').values_list('name', flat=True))
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
