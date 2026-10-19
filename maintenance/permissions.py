"""Central role / object-level checks for the maintenance module.

Module access itself is the soft-coded ``nav.maintenance`` permission
(see ``decorators.maintenance_required``). These helpers decide who counts as
maintenance staff and which tickets a non-staff user may see.
"""
from django.core.exceptions import PermissionDenied

STAFF_ROLES = ('admin', 'manager', 'production_manager', 'maintenance_engineer')


def role_of(user):
    profile = getattr(user, 'profile', None)
    return (getattr(profile, 'role', '') or '').strip().lower()


def is_maintenance_staff(user):
    return bool(getattr(user, 'is_superuser', False) or role_of(user) in STAFF_ROLES)


def can_view_record(user, record):
    """Staff see every ticket; everyone else only the tickets they reported."""
    return is_maintenance_staff(user) or record.reported_by_id == user.id


def require_record_access(user, record):
    if not can_view_record(user, record):
        raise PermissionDenied
