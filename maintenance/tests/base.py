from django.contrib.auth import get_user_model
from django.test import TestCase, override_settings

from core.models import Machine, Permission, Role, UserProfile
from maintenance.models import MachineDowntime, MaintenanceRecord
from maintenance.services import generate_record_no

User = get_user_model()
PASSWORD = 'pw-for-tests-12345'


def make_user(username, role, with_module_access=True):
    user = User.objects.create_user(username, password=PASSWORD)
    profile, _ = UserProfile.objects.get_or_create(user=user)
    profile.role = role
    profile.save()
    role_row, _ = Role.objects.get_or_create(slug=role, defaults={'display_name': role})
    if with_module_access:
        permission, _ = Permission.objects.get_or_create(code='nav.maintenance', defaults={'name': 'Maintenance'})
        role_row.permissions.add(permission)
    # Reload so a profile cached on the instance by the post-save signal can't mask the role.
    return User.objects.get(pk=user.pk)


@override_settings(PASSWORD_HASHERS=['django.contrib.auth.hashers.MD5PasswordHasher'])
class MaintenanceTestCase(TestCase):
    """Fixtures shared by the maintenance tests: one machine, an operator, an
    engineer, a manager and a user without module access."""

    def setUp(self):
        from supply_chain.models import ItemRequestType

        for code, name in (('MNT', 'Maintenance'), ('SRV', 'Service')):
            ItemRequestType.objects.get_or_create(code=code, defaults={'name': name, 'is_active': True})
        self.machine = Machine.objects.create(name='SM74-A', machine_type='offset_printing')
        self.other_machine = Machine.objects.create(name='GTO-B', machine_type='offset_printing')
        self.operator = make_user('op1', 'operator')
        self.operator2 = make_user('op2', 'operator')
        self.engineer = make_user('eng1', 'maintenance_engineer')
        self.manager = make_user('mgr1', 'manager')
        self.outsider = make_user('planner1', 'planner', with_module_access=False)

    def login(self, user):
        self.client.force_login(user)

    def make_record(self, machine=None, reporter=None, status='REPORTED', stopped=False, **extra):
        from django.utils import timezone

        record = MaintenanceRecord.objects.create(
            machine=machine or self.machine,
            reported_date=timezone.localdate(),
            reported_by=reporter or self.operator,
            fault_description=extra.pop('fault_description', 'Feed table jam'),
            status=status,
            machine_stopped=stopped,
            **extra,
        )
        generate_record_no(record)
        if stopped:
            from maintenance.services import open_downtime
            open_downtime(record)
        return record

    def open_downtimes(self, machine=None):
        return MachineDowntime.objects.filter(machine=machine or self.machine, end_at__isnull=True)
