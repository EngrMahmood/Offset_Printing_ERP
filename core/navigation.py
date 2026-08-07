from __future__ import annotations

from typing import Any

# Permission codes for each nav module. Actual grants are DB-driven
# (core.models.Role / Permission, editable from Settings -> Roles & Permissions);
# these constants are only used to seed default grants for the built-in roles
# (see core.management.commands.seed_access_control).
NAV_PERMISSION_CODES = {
    'planning': 'nav.planning',
    'qc': 'nav.qc',
    'production': 'nav.production',
    'dispatch': 'nav.dispatch',
    'master_data': 'nav.master_data',
    'migration': 'nav.migration',
    'reports': 'nav.reports',
    'job_summary': 'nav.job_summary',
    'printing_plates': 'nav.printing_plates',
    'guides': 'nav.guides',
    'supply_chain': 'nav.supply_chain',
    'audit': 'nav.audit',
    'item_request': 'nav.item_request',
    'maintenance': 'nav.maintenance',
    'chat': 'nav.chat',
}

PLANNING_NAV_ROLES = {'admin', 'manager', 'planner'}
QC_NAV_ROLES = {'admin', 'manager', 'qc', 'production_manager'}
PRODUCTION_NAV_ROLES = {'admin', 'manager', 'planner', 'production_manager', 'production', 'operator'}
DISPATCH_NAV_ROLES = {'admin', 'manager', 'dispatch'}
MASTER_DATA_NAV_ROLES = {'admin', 'manager', 'production_manager'}
MIGRATION_NAV_ROLES = {'admin', 'manager', 'planner'}
REPORTS_NAV_ROLES = {'admin', 'manager', 'planner', 'production_manager', 'viewer'}
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
MAINTENANCE_NAV_ROLES = {'admin', 'manager', 'maintenance_engineer', 'production_manager', 'production', 'operator'}
# Chat is an org-wide utility — open to every role by default (see chat.management.commands.seed_chat_permissions,
# which owns the actual Permission/Role grants; this set exists only for parity with the other *_NAV_ROLES constants).
CHAT_NAV_ROLES = {
    'admin', 'manager', 'planner', 'production_manager', 'production', 'graphics_designer',
    'operator', 'dispatch', 'qc', 'storekeeper', 'finance', 'supply_chain', 'maintenance_engineer',
}


def _role_from_request(request: Any) -> str:
    profile = getattr(getattr(request, 'user', None), 'profile', None)
    return (getattr(profile, 'role', '') or '').strip().lower()


def _nav_layout_from_request(request: Any) -> dict:
    profile = getattr(getattr(request, 'user', None), 'profile', None)
    layout = getattr(profile, 'nav_layout', None) or {}
    if not isinstance(layout, dict):
        return {}
    # "row1" replaced the old single-row "pinned" key when the two-row nav
    # customization was added; accept either shape so accounts that saved a
    # layout before this change don't lose it.
    row1 = layout.get('row1', layout.get('pinned'))
    row2 = layout.get('row2', [])
    overflow = layout.get('overflow')
    if not isinstance(row1, list) or not isinstance(row2, list) or not isinstance(overflow, list):
        return {}
    return {'row1': row1, 'row2': row2, 'overflow': overflow}


def get_nav_permissions(request: Any) -> dict[str, bool | str]:
    role = _role_from_request(request)
    is_authenticated = bool(getattr(getattr(request, 'user', None), 'is_authenticated', False))
    is_superuser = bool(getattr(getattr(request, 'user', None), 'is_superuser', False))
    is_staff = bool(getattr(getattr(request, 'user', None), 'is_staff', False))

    if is_superuser:
        role = role or 'admin'

    def _allow(nav_key: str) -> bool:
        if not is_authenticated:
            return False
        if is_superuser:
            return True
        from core.permissions import user_has_permission
        return user_has_permission(request.user, NAV_PERMISSION_CODES[nav_key])

    def _pending_item_reviews() -> int:
        if not _allow('item_request'):
            return 0
        # Imported lazily so core does not hard-depend on supply_chain at import time.
        from supply_chain.item_request_service import pending_review_count
        try:
            return pending_review_count(request.user)
        except Exception:
            return 0

    can_review_overrides = is_authenticated and (is_staff or role in {'admin', 'manager', 'production_manager'})
    # Who may set a per-job pass-count override (supervisory, includes production).
    can_set_pass_override = is_authenticated and (
        is_staff or role in {'admin', 'manager', 'production_manager', 'production'}
    )

    def _pending_override_reviews() -> int:
        if not can_review_overrides:
            return 0
        # Imported lazily to avoid a model import at module load time.
        from core.models import EditOverrideRequest
        try:
            return EditOverrideRequest.objects.filter(status='pending').count()
        except Exception:
            return 0

    def _my_override_actionable() -> int:
        """Count of the current user's override requests that are approved and
        ready to act on (unexpired, not yet consumed)."""
        if not is_authenticated:
            return 0
        from django.utils import timezone
        from core.models import EditOverrideRequest
        try:
            return EditOverrideRequest.objects.filter(
                requested_by=request.user,
                status='approved',
                consumed_at__isnull=True,
                expires_at__gt=timezone.now(),
            ).count()
        except Exception:
            return 0

    return {
        'role': role,
        'nav_layout': _nav_layout_from_request(request),
        'item_request_pending_count': _pending_item_reviews(),
        'can_access_dashboard': is_authenticated,
        'can_access_planning': _allow('planning'),
        'can_access_qc': _allow('qc'),
        'can_access_production': _allow('production'),
        'can_access_dispatch': _allow('dispatch'),
        'can_access_master_data': _allow('master_data'),
        'can_access_migration': _allow('migration'),
        'can_access_reports': _allow('reports'),
        'can_access_job_summary': _allow('job_summary'),
        'can_access_printing_plates': _allow('printing_plates'),
        'can_access_guides': _allow('guides'),
        'can_access_supply_chain': _allow('supply_chain'),
        'can_access_audit': _allow('audit'),
        'can_access_item_request': _allow('item_request'),
        'can_access_maintenance': _allow('maintenance'),
        'can_access_chat': _allow('chat'),
        'can_access_tasks': is_authenticated,
        'can_access_floor_dashboard': is_authenticated,
        # Managers/admins who can approve edit-lock override requests.
        'can_review_overrides': can_review_overrides,
        'override_pending_count': _pending_override_reviews(),
        # Supervisory set of roles who may set a per-job pass-count override.
        'can_set_pass_override': can_set_pass_override,
        # The current user's approved-and-ready override requests (for "My Requests").
        'my_override_actionable_count': _my_override_actionable(),
    }
