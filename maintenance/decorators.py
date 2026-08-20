from functools import wraps

from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied

from core.navigation import get_nav_permissions


def maintenance_required(view_func):
    @login_required
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        if not get_nav_permissions(request).get('can_access_maintenance'):
            raise PermissionDenied
        return view_func(request, *args, **kwargs)

    # Gates on the soft-coded nav.maintenance Permission — configurable from
    # Settings -> Roles & Access Control.
    _wrapped._rbac_configurable = True
    return _wrapped


def maintenance_manager_required(view_func):
    @login_required
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        nav = get_nav_permissions(request)
        if not (nav.get('can_access_maintenance') and nav.get('role') in ('admin', 'manager')):
            raise PermissionDenied
        return view_func(request, *args, **kwargs)

    # nav.maintenance is soft-coded, but the extra admin/manager check on top
    # is hardcoded — not fully UI-configurable.
    _wrapped._rbac_hardcoded = True
    return _wrapped


def maintenance_staff_required(view_func):
    """Gate for actions beyond raising/viewing a complaint — triaging, working a
    ticket, adding spare/service lines, raising demand. Open to the maintenance
    engineer and above, not to the operator/supervisor who merely reported it."""
    @login_required
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        nav = get_nav_permissions(request)
        allowed_roles = ('admin', 'manager', 'production_manager', 'maintenance_engineer')
        if not (nav.get('can_access_maintenance') and nav.get('role') in allowed_roles):
            raise PermissionDenied
        return view_func(request, *args, **kwargs)

    _wrapped._rbac_hardcoded = True
    return _wrapped
