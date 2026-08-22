from django import template

register = template.Library()

_PRIORITY_BADGES = {
    'CRITICAL': 'erp-badge-rejected',
    'MAJOR': 'erp-badge-pending',
    'MEDIUM': 'erp-badge-hold',
    'LOW': 'erp-badge-neutral',
}

_STATUS_BADGES = {
    'PENDING_APPROVAL': 'erp-badge-neutral',
    'REPORTED': 'erp-badge-pending',
    'DIAGNOSED': 'erp-badge-hold',
    'AWAITING_PARTS': 'erp-badge-pending',
    'AWAITING_VENDOR': 'erp-badge-pending',
    'IN_PROGRESS': 'erp-badge-hold',
    'COMPLETED': 'erp-badge-approved',
    'VERIFIED': 'erp-badge-approved',
    'CLOSED': 'erp-badge-archived',
    'CANCELLED': 'erp-badge-rejected',
    'REJECTED': 'erp-badge-rejected',
}


@register.filter
def priority_badge(priority):
    return _PRIORITY_BADGES.get(priority, 'erp-badge-neutral')


@register.filter
def status_badge(status):
    return _STATUS_BADGES.get(status, 'erp-badge-neutral')
