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

    return _wrapped


def maintenance_manager_required(view_func):
    @login_required
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        nav = get_nav_permissions(request)
        if not (nav.get('can_access_maintenance') and nav.get('role') in ('admin', 'manager')):
            raise PermissionDenied
        return view_func(request, *args, **kwargs)

    return _wrapped
