from functools import wraps

from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied

from core.navigation import get_nav_permissions
from core.permissions import user_has_permission


def bom_required(view_func):
    @login_required
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        if not get_nav_permissions(request).get('can_access_bom'):
            raise PermissionDenied
        return view_func(request, *args, **kwargs)

    _wrapped._rbac_configurable = True
    return _wrapped


def bom_manage_required(view_func):
    @login_required
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        nav = get_nav_permissions(request)
        if not nav.get('can_access_bom'):
            raise PermissionDenied
        if not (request.user.is_superuser or user_has_permission(request.user, 'action.manage_bom')):
            raise PermissionDenied
        return view_func(request, *args, **kwargs)

    _wrapped._rbac_configurable = True
    return _wrapped


def bom_approve_required(view_func):
    @login_required
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        nav = get_nav_permissions(request)
        if not nav.get('can_access_bom'):
            raise PermissionDenied
        if not (request.user.is_superuser or user_has_permission(request.user, 'action.approve_bom')):
            raise PermissionDenied
        return view_func(request, *args, **kwargs)

    _wrapped._rbac_configurable = True
    return _wrapped
