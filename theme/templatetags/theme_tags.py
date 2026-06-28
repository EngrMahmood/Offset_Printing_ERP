from django import template

register = template.Library()


_APP_LABELS = {
    'qc': 'QC',
    'planning': 'Planning',
    'core': 'Core',
    'supply_chain': 'Supply Chain',
    'theme': 'Theme',
}

_URL_LABELS = {
    'approval_queue': 'Approval Queue',
    'master_review': 'SKU Master Review',
    'pending_skus': 'Pending SKUs',
    'pending_skus_ignored': 'Ignored SKUs',
    'pending_sku_master_entry': 'SKU Master Entry',
    'po_inbox': 'PO Inbox',
    'po_upload': 'Upload PO',
    'po_review': 'PO Review',
    'manual_po_entry': 'Manual PO Entry',
    'sku_recipes': 'SKU Recipes',
    'sku_recipes_archived': 'Archived SKUs',
    'sku_recipe_bulk_upload': 'Bulk Upload',
    'sku_recipes_pending_review': 'Pending Review',
    'jobs': 'Planning Jobs',
    'job_detail': 'Job Detail',
    'job_edit': 'Edit Job',
    'planning_archived_jobs': 'Archived Jobs',
    'scan': 'Production Scan',
    'planning_welcome': 'Overview',
    'monthly_demand': 'Monthly Demand',
    'opening': 'Stock Opening',
    'receiving': 'Stock Receiving',
    'issuance': 'Stock Issuance',
    'adjustment': 'Stock Adjustment',
    'items': 'Item Master',
    'item_edit': 'Edit Item',
    'consumption_reports': 'Consumption Reports',
    'kpi_dashboard': 'Inventory KPIs',
    'jc_links': 'Job Card Links',
    'physical_counts': 'Physical Stock Count',
}


@register.filter
def pretty_app_name(value):
    """Convert app name like 'qc' to human label 'QC'."""
    return _APP_LABELS.get(value or '', (value or '').title())


@register.filter
def pretty_url_name(value):
    """Convert url_name like 'approval_queue' to 'Approval Queue'."""
    if not value:
        return ''
    if value in _URL_LABELS:
        return _URL_LABELS[value]
    return value.replace('_', ' ').title()


@register.filter
def jc_status_badge(status):
    """Map job card workflow_status to erp-badge variant class."""
    mapping = {
        'planning_approved': 'erp-badge-approved',
        'qc_approved': 'erp-badge-approved',
        'production_approved': 'erp-badge-approved',
        'completed': 'erp-badge-archived',
        'closed': 'erp-badge-archived',
        'in_production': 'erp-badge-archived',
        'qc_rejected': 'erp-badge-rejected',
        'pm_rejected': 'erp-badge-rejected',
        'released': 'erp-badge-hold',
        'draft': 'erp-badge-draft',
        'pending_data': 'erp-badge-draft',
    }
    return mapping.get((status or '').lower(), 'erp-badge-pending')


@register.filter
def recipe_status_badge(status):
    """Map SKU recipe master_data_status to erp-badge variant class."""
    mapping = {
        'approved': 'erp-badge-approved',
        'reviewed': 'erp-badge-hold',
        'pending_review': 'erp-badge-pending',
        'draft': 'erp-badge-draft',
        'archived': 'erp-badge-archived',
    }
    return mapping.get((status or '').lower(), 'erp-badge-pending')


@register.simple_tag(takes_context=True)
def is_active_app(context, app_name):
    request = context.get('request')
    if not request or not getattr(request, 'resolver_match', None):
        return ''
    return 'is-active' if request.resolver_match.app_name == app_name else ''


@register.simple_tag(takes_context=True)
def is_active_url(context, url_name):
    request = context.get('request')
    if not request or not getattr(request, 'resolver_match', None):
        return ''
    return 'is-active' if request.resolver_match.url_name == url_name else ''
