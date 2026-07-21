from __future__ import annotations

from typing import Any

PLANNING_NAV_ROLES = {'admin', 'manager', 'planner'}
QC_NAV_ROLES = {'admin', 'manager', 'qc', 'production_manager'}
PRODUCTION_NAV_ROLES = {'admin', 'manager', 'planner', 'production_manager', 'production', 'operator'}
DISPATCH_NAV_ROLES = {'admin', 'manager', 'dispatch'}
MASTER_DATA_NAV_ROLES = {'admin', 'manager', 'production_manager'}
MIGRATION_NAV_ROLES = {'admin', 'manager', 'planner'}
REPORTS_NAV_ROLES = {'admin', 'manager', 'planner', 'production_manager'}
JOB_SUMMARY_NAV_ROLES = {'admin', 'manager', 'planner', 'production_manager', 'production', 'qc', 'dispatch'}
PRINTING_PLATES_NAV_ROLES = {'admin', 'manager', 'planner', 'production_manager', 'graphics_designer'}
GUIDE_NAV_ROLES = {
    'admin', 'manager', 'planner', 'qc', 'production_manager', 'production', 'dispatch', 'finance', 'operator'
}
SUPPLY_CHAIN_NAV_ROLES = {'admin', 'manager', 'supply_chain'}
AUDIT_NAV_ROLES = {'admin', 'manager', 'planner', 'production_manager', 'production', 'qc', 'dispatch'}
ITEM_REQUEST_NAV_ROLES = {
    'admin', 'manager', 'supply_chain', 'planner', 'production_manager', 'production',
    'graphics_designer', 'operator', 'dispatch', 'qc', 'storekeeper', 'finance',
}
MAINTENANCE_NAV_ROLES = {'admin', 'manager', 'maintenance_engineer', 'production_manager'}


def _role_from_request(request: Any) -> str:
    profile = getattr(getattr(request, 'user', None), 'profile', None)
    return (getattr(profile, 'role', '') or '').strip().lower()


def get_nav_permissions(request: Any) -> dict[str, bool | str]:
    role = _role_from_request(request)
    is_authenticated = bool(getattr(getattr(request, 'user', None), 'is_authenticated', False))
    is_superuser = bool(getattr(getattr(request, 'user', None), 'is_superuser', False))

    if is_superuser:
        role = role or 'admin'

    def _allow(roles: set[str]) -> bool:
        if not is_authenticated:
            return False
        return is_superuser or role in roles

    def _pending_item_reviews() -> int:
        if not _allow(ITEM_REQUEST_NAV_ROLES):
            return 0
        # Imported lazily so core does not hard-depend on supply_chain at import time.
        from supply_chain.item_request_service import pending_review_count
        try:
            return pending_review_count(request.user)
        except Exception:
            return 0

    return {
        'role': role,
        'item_request_pending_count': _pending_item_reviews(),
        'can_access_dashboard': is_authenticated,
        'can_access_planning': _allow(PLANNING_NAV_ROLES),
        'can_access_qc': _allow(QC_NAV_ROLES),
        'can_access_production': _allow(PRODUCTION_NAV_ROLES),
        'can_access_dispatch': _allow(DISPATCH_NAV_ROLES),
        'can_access_master_data': _allow(MASTER_DATA_NAV_ROLES),
        'can_access_migration': _allow(MIGRATION_NAV_ROLES),
        'can_access_reports': _allow(REPORTS_NAV_ROLES),
        'can_access_job_summary': _allow(JOB_SUMMARY_NAV_ROLES),
        'can_access_printing_plates': _allow(PRINTING_PLATES_NAV_ROLES),
        'can_access_guides': _allow(GUIDE_NAV_ROLES),
        'can_access_supply_chain': _allow(SUPPLY_CHAIN_NAV_ROLES),
        'can_access_audit': _allow(AUDIT_NAV_ROLES),
        'can_access_item_request': _allow(ITEM_REQUEST_NAV_ROLES),
        'can_access_maintenance': _allow(MAINTENANCE_NAV_ROLES),
        'can_access_tasks': is_authenticated,
    }
