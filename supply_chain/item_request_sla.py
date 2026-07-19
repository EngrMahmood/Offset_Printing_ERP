from django.utils import timezone

SLA_DAYS_BY_STATUS = {
    'MGR_REVIEW': 2,
    'SC_REVIEW': 2,
    'NEEDS_REVISION': 3,
    'APPROVED': 1,
    'IN_PROCUREMENT': 14,
}


def stage_entered_at(item_request, approvals=None):
    """Best-known timestamp the request entered its current status."""
    if approvals is None:
        approvals = item_request.approvals.all()
    for approval in reversed(list(approvals)):
        if approval.action in ('SUBMIT', 'RESUBMIT') and item_request.status == 'MGR_REVIEW':
            return approval.created_at
        if approval.action == 'APPROVE' and approval.stage == 'MANAGER' and item_request.status == 'SC_REVIEW':
            return approval.created_at
        if approval.action == 'APPROVE' and approval.stage == 'SUPPLY_CHAIN' and item_request.status in ('APPROVED', 'IN_PROCUREMENT'):
            return approval.created_at
        if approval.action == 'REVISE' and item_request.status == 'NEEDS_REVISION':
            return approval.created_at
    return item_request.created_at


def sla_days_limit(status):
    return SLA_DAYS_BY_STATUS.get(status)


def sla_status(item_request, approvals=None):
    """Return (days_in_stage, days_limit, breached: bool|None)."""
    limit = sla_days_limit(item_request.status)
    if limit is None:
        return None, None, None
    entered_at = stage_entered_at(item_request, approvals)
    days_in_stage = (timezone.now() - entered_at).days
    return days_in_stage, limit, days_in_stage > limit
