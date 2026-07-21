from __future__ import annotations

from reports.services import (
    build_daily_production_context,
    build_dispatch_tracking_context,
    build_job_planning_context,
    build_machine_planning_context,
    build_plates_planning_context,
    build_production_insights_context,
    build_qc_approvals_context,
    build_raw_material_cutting_context,
    build_wastage_report_context,
)
from reports.report_registry.metadata import ReportDefinition
from reports.report_registry.registry import registry


def _machine_planning_executor(request, filters):
    return build_machine_planning_context(request)


def _daily_production_executor(request, filters):
    return build_daily_production_context(request)


def _dispatch_tracking_executor(request, filters):
    return build_dispatch_tracking_context(request)


def _job_planning_executor(request, filters):
    return build_job_planning_context(request)


def _plates_planning_executor(request, filters):
    return build_plates_planning_context(request)


def _production_insights_executor(request, filters):
    return build_production_insights_context(request)


def _qc_approvals_executor(request, filters):
    return build_qc_approvals_context(request)


def _raw_material_cutting_executor(request, filters):
    return build_raw_material_cutting_context(request)


def _wastage_report_executor(request, filters):
    return build_wastage_report_context(request)


registry.register(
    ReportDefinition(
        slug='daily-production',
        title='Daily Production',
        description='Day-by-day printing impressions, packing output, and dispatch quantities.',
        department='production',
        permissions=('core.view_reports',),
        filters=('date_range', 'machine'),
        supported_exports=('csv', 'xlsx', 'pdf'),
        supported_charts=('line', 'bar', 'stacked_bar'),
        drilldown_support=False,
        cache_timeout=300,
        icon='fa-calendar-day',
        category='execution',
        navigation_group='operations',
        executor=_daily_production_executor,
    )
)

registry.register(
    ReportDefinition(
        slug='machine-planning',
        title='Machine Planning',
        description='Planned load by machine, backlog, and machine capacity vs actual output.',
        department='planning',
        permissions=('core.view_reports',),
        filters=('date_range', 'department', 'machine', 'status', 'po', 'job_card', 'sku'),
        supported_exports=('csv', 'xlsx', 'pdf'),
        supported_charts=('line', 'bar', 'stacked_bar', 'top_n'),
        drilldown_support=True,
        cache_timeout=300,
        icon='fa-industry',
        category='planning',
        navigation_group='operations',
        executor=_machine_planning_executor,
    )
)

registry.register(
    ReportDefinition(
        slug='job-planning',
        title='Job Planning',
        description='Job status mix, due-date pressure, repeat mix, and planning readiness.',
        department='planning',
        permissions=('core.view_reports',),
        filters=('date_range', 'department', 'status', 'po', 'job_card', 'sku', 'customer'),
        supported_exports=('csv', 'xlsx', 'pdf'),
        supported_charts=('bar', 'donut', 'line', 'top_n'),
        drilldown_support=True,
        cache_timeout=300,
        icon='fa-calendar-check',
        category='planning',
        navigation_group='operations',
        executor=_job_planning_executor,
    )
)

registry.register(
    ReportDefinition(
        slug='plates-planning',
        title='Plates Planning',
        description='Plate set readiness, blocked jobs, and QC gate visibility.',
        department='printing_plates',
        permissions=('core.view_reports',),
        filters=('date_range', 'status', 'po', 'job_card', 'sku'),
        supported_exports=('csv', 'xlsx', 'pdf'),
        supported_charts=('bar', 'donut', 'top_n'),
        drilldown_support=True,
        cache_timeout=300,
        icon='fa-clone',
        category='planning',
        navigation_group='operations',
        executor=_plates_planning_executor,
    )
)

registry.register(
    ReportDefinition(
        slug='production-insights',
        title='Production Insights',
        description='Output, waste, downtime, OEE proxies, and machine/shift efficiency.',
        department='production',
        permissions=('core.view_reports',),
        filters=('date_range', 'machine', 'shift', 'operator', 'po', 'job_card', 'sku'),
        supported_exports=('csv', 'xlsx', 'pdf'),
        supported_charts=('line', 'bar', 'stacked_bar', 'pareto', 'top_n'),
        drilldown_support=True,
        cache_timeout=300,
        icon='fa-chart-line',
        category='execution',
        navigation_group='operations',
        executor=_production_insights_executor,
    )
)

registry.register(
    ReportDefinition(
        slug='qc-approvals',
        title='QC Approvals',
        description='Approval queue status, rejection trends, and QC turnaround time.',
        department='qc',
        permissions=('core.view_reports',),
        filters=('date_range', 'status', 'po', 'job_card', 'sku', 'approval_status'),
        supported_exports=('csv', 'xlsx', 'pdf'),
        supported_charts=('bar', 'donut', 'line', 'top_n'),
        drilldown_support=True,
        cache_timeout=300,
        icon='fa-check-circle',
        category='quality',
        navigation_group='quality',
        executor=_qc_approvals_executor,
    )
)

registry.register(
    ReportDefinition(
        slug='dispatch-tracking',
        title='Dispatch Tracking',
        description='Fulfillment rates, dispatch completion, and delivery backlog.',
        department='dispatch',
        permissions=('core.view_reports',),
        filters=('date_range', 'destination', 'status', 'po', 'job_card', 'customer'),
        supported_exports=('csv', 'xlsx', 'pdf'),
        supported_charts=('line', 'bar', 'donut', 'top_n'),
        drilldown_support=True,
        cache_timeout=300,
        icon='fa-truck',
        category='execution',
        navigation_group='operations',
        executor=_dispatch_tracking_executor,
    )
)

registry.register(
    ReportDefinition(
        slug='raw-material-cutting-request',
        title='Raw Material Cutting Request',
        description='Material cutting requirements for released jobs, including sheet sizes and quantities.',
        department='production',
        permissions=('core.view_reports',),
        filters=('date_range', 'material', 'status', 'po', 'job_card', 'sku', 'machine'),
        supported_exports=('csv', 'xlsx', 'pdf'),
        supported_charts=('top_n', 'bar'),
        drilldown_support=True,
        cache_timeout=300,
        icon='fa-cut',
        category='execution',
        navigation_group='operations',
        executor=_raw_material_cutting_executor,
    )
)

registry.register(
    ReportDefinition(
        slug='wastage-report',
        title='Wastage Report',
        description='Process-wise wastage analysis including printing, sorting, and dispatch gaps (tentative vs finalized).',
        department='production',
        permissions=('core.view_reports',),
        filters=('date_range', 'status', 'po', 'job_card', 'sku', 'machine', 'wastage_status', 'high_wastage'),
        supported_exports=('csv', 'xlsx', 'pdf'),
        supported_charts=('bar', 'donut', 'line', 'top_n'),
        drilldown_support=True,
        cache_timeout=300,
        icon='fa-trash-alt',
        category='execution',
        navigation_group='operations',
        executor=_wastage_report_executor,
    )
)
