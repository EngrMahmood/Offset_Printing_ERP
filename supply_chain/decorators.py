from functools import wraps

from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied

from core.navigation import get_nav_permissions


def supply_chain_required(view_func):
    @login_required
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        if not get_nav_permissions(request).get('can_access_supply_chain'):
            raise PermissionDenied
        return view_func(request, *args, **kwargs)

    # Gates on the soft-coded nav.supply_chain Permission — configurable from
    # Settings -> Roles & Access Control, so audit_view_permissions shouldn't
    # flag views using this decorator as unguarded.
    _wrapped._rbac_configurable = True
    return _wrapped


def item_request_access_required(view_func):
    @login_required
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        if not get_nav_permissions(request).get('can_access_item_request'):
            raise PermissionDenied
        return view_func(request, *args, **kwargs)

    _wrapped._rbac_configurable = True
    return _wrapped
