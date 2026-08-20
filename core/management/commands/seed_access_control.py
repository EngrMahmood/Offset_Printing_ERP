from django.core.management.base import BaseCommand
from django.db import transaction

from core.models import Role, Permission, UserProfile
from core import navigation as nav

# nav_key -> (display name, description)
NAV_PERMISSION_INFO = {
    'planning': ('Planning', 'Access the Planning module'),
    'qc': ('QC', 'Access the QC module'),
    'production': ('Production', 'Access the Production module'),
    'dispatch': ('Dispatch', 'Access the Dispatch module'),
    'master_data': ('Master Data', 'Access Master Data management'),
    'migration': ('Migration', 'Access the Migration module'),
    'reports': ('Reports', 'Access Reports'),
    'job_summary': ('Jobs Summary', 'Access the Jobs Summary dashboard'),
    'printing_plates': ('Printing Plates', 'Access the Printing Plates module'),
    'guides': ('Guides', 'Access Guides'),
    'supply_chain': ('Supply Chain', 'Access the Supply Chain module'),
    'audit': ('Audit', 'Access the Audit module'),
    'item_request': ('Item Requests', 'Access Item Requests'),
    'maintenance': ('Maintenance', 'Access the Maintenance module'),
}

# nav_key -> legacy hardcoded role set (default grants for built-in roles)
NAV_DEFAULT_ROLES = {
    'planning': nav.PLANNING_NAV_ROLES,
    'qc': nav.QC_NAV_ROLES,
    'production': nav.PRODUCTION_NAV_ROLES,
    'dispatch': nav.DISPATCH_NAV_ROLES,
    'master_data': nav.MASTER_DATA_NAV_ROLES,
    'migration': nav.MIGRATION_NAV_ROLES,
    'reports': nav.REPORTS_NAV_ROLES,
    'job_summary': nav.JOB_SUMMARY_NAV_ROLES,
    'printing_plates': nav.PRINTING_PLATES_NAV_ROLES,
    'guides': nav.GUIDE_NAV_ROLES,
    'supply_chain': nav.SUPPLY_CHAIN_NAV_ROLES,
    'audit': nav.AUDIT_NAV_ROLES,
    'item_request': nav.ITEM_REQUEST_NAV_ROLES,
    'maintenance': nav.MAINTENANCE_NAV_ROLES,
}

# action_key -> (permission code, display name, description, category, legacy default role set)
# Mirrors the hardcoded tuples in each UserProfile.can_*() method (core/models.py) so that
# flipping those methods onto Permission checks doesn't change behavior for existing roles.
ACTION_PERMISSIONS = {
    'view_plate_queue': (
        'View Plate Queue', 'Can view the printing plate request queue', 'Printing Plates',
        {'admin', 'manager', 'planner', 'graphics_designer'},
    ),
    'create_plate_request': (
        'Create Plate Request', 'Can create a new plate request', 'Printing Plates',
        {'admin', 'manager', 'planner', 'graphics_designer'},
    ),
    'send_plate': (
        'Send Plate', 'Can mark a plate as sent', 'Printing Plates',
        {'admin', 'manager', 'planner', 'graphics_designer'},
    ),
    'receive_plate': (
        'Receive Plate', 'Can mark a plate as received', 'Printing Plates',
        {'admin', 'manager', 'planner', 'graphics_designer'},
    ),
    'archive_plate': (
        'Archive Plate', 'Can archive a plate request', 'Printing Plates',
        {'admin', 'manager', 'planner', 'graphics_designer'},
    ),
    'edit_jobcard': (
        'Edit Job Card', 'Can create/edit job cards', 'Planning Actions',
        {'admin', 'manager', 'planner', 'production_manager'},
    ),
    'approve_planning': (
        'Approve Planning', 'Can approve the planning queue', 'Planning Actions',
        {'admin', 'manager', 'planner'},
    ),
    'cancel_planning_job': (
        'Cancel Planning Job', 'Can cancel a planning job the customer no longer needs', 'Planning Actions',
        {'admin', 'manager', 'planner'},
    ),
    'view_jobcard': (
        'View Job Cards (Read-Only)', 'Can view job card / SKU recipe reference pages without edit rights',
        'Planning Actions',
        {'admin', 'manager', 'planner', 'production_manager'},
    ),
    'view_planning_queue': (
        'View Planning Queue', 'Can view the planning queue', 'Planning Actions',
        {'admin', 'manager', 'planner', 'production_manager', 'qc'},
    ),
    'plan': (
        'Plan Jobs', 'Can create, edit, and manage planning jobs', 'Planning Actions',
        {'admin', 'manager', 'planner'},
    ),
    'view_approval_queue': (
        'View Approval Queue', 'Can access the approval queue page', 'Planning Actions',
        {'admin', 'manager', 'planner', 'qc', 'production_manager'},
    ),
    'edit_production': (
        'Edit Production', 'Can log production data', 'Production Actions',
        {'admin', 'manager', 'production_manager', 'production', 'operator'},
    ),
    'start_production': (
        'Start Production', 'Can move a released Job Card into production execution', 'Production Actions',
        {'admin', 'manager', 'production_manager', 'production'},
    ),
    'manage_operators': (
        'Manage Operators', 'Can assign operators to shifts/jobs', 'Production Actions',
        {'admin', 'manager', 'production_manager', 'production'},
    ),
    'view_production_records': (
        'View Production/Packing Records (Read-Only)', 'Can view the production and packing records ledger without edit rights',
        'Production Actions',
        {'admin', 'manager', 'production_manager', 'production', 'operator'},
    ),
    'approve_qc': (
        'Approve QC', 'Can perform QC checks', 'QC & Approvals',
        {'admin', 'qc', 'manager'},
    ),
    'view_qc_queue': (
        'View QC Queue', 'Can view the QC queue', 'QC & Approvals',
        {'admin', 'manager', 'qc', 'production_manager'},
    ),
    'approve_pm': (
        'Approve PM (Production Manager)', 'Can perform production manager approval', 'QC & Approvals',
        {'admin', 'manager', 'production_manager'},
    ),
    'view_pm_queue': (
        'View PM Queue', 'Can view the production manager approval queue', 'QC & Approvals',
        {'admin', 'manager', 'production_manager'},
    ),
    'approve_dispatch': (
        'Approve Dispatch', 'Can approve/edit dispatch', 'Dispatch',
        {'admin', 'manager', 'dispatch'},
    ),
    'view_dispatch_records': (
        'View Dispatch Records (Read-Only)', 'Can view the dispatch records ledger without edit rights', 'Dispatch',
        {'admin', 'manager', 'dispatch'},
    ),
    'manage_masters': (
        'Manage Masters', 'Can manage machines, operators, materials, departments', 'Master Data & Masters',
        {'admin', 'manager', 'production_manager'},
    ),
    'archive_records': (
        'Archive Records', 'Can archive and restore operational records', 'Master Data & Masters',
        {'admin', 'manager', 'production_manager'},
    ),
    'view_analytics': (
        'View Analytics', 'Can view dashboard and analytics', 'Reports & Analytics',
        {'admin', 'manager', 'planner', 'production', 'dispatch', 'finance'},
    ),
    'view_reports': (
        'View Reports', 'Can view financial/operational reports', 'Reports & Analytics',
        {'admin', 'manager', 'finance'},
    ),
    'view_job_summary': (
        'View Job Summary', 'Can view the Jobs Summary dashboard', 'Reports & Analytics',
        {'admin', 'manager', 'planner', 'production_manager', 'production', 'qc', 'dispatch'},
    ),
    'view_sku_master_review_queue': (
        'View SKU Master Review Queue', 'Can view the SKU master review queue', 'SKU Master Review',
        {'admin', 'qc', 'manager', 'planner'},
    ),
    'approve_sku_master_review': (
        'Approve SKU Master Review', 'Can approve/reject SKUs in master review', 'SKU Master Review',
        {'admin', 'qc', 'manager'},
    ),
    'supply_chain_admin': (
        'Supply Chain Admin', 'Elevated admin actions within Supply Chain (approvals, edits, admin views)', 'Supply Chain',
        {'admin', 'manager'},
    ),
    'manage_tasks': (
        'Manage Tasks & Teams', 'Can manage teams and edit/administer other users\' tasks', 'Tasks',
        {'admin', 'manager'},
    ),
    'view_production_wip': (
        'View Production WIP Board', 'Can view the production work-in-progress board', 'Production Actions',
        {'admin', 'manager', 'planner', 'production_manager', 'production', 'operator', 'viewer'},
    ),
    'manage_production_wip_statuses': (
        'Manage Production WIP Statuses', 'Can add new WIP status columns to the production board', 'Production Actions',
        {'admin'},
    ),
    'finalize_job_card': (
        'Finalize Job Cards', 'Can manually close/reopen job cards stuck near completion', 'Production Actions',
        {'admin', 'manager'},
    ),
    'set_pass_override': (
        'Set Print Pass Count Override', 'Can override a job\'s planned print pass count', 'Production Actions',
        {'admin', 'manager', 'production_manager', 'production'},
    ),
    'view_released_jobs': (
        'View Released Jobs (Plate Status)', 'Can view released jobs and request plate replacement', 'Production Actions',
        {'admin', 'manager', 'planner', 'production_manager', 'production', 'operator', 'graphics_designer'},
    ),
}


class Command(BaseCommand):
    help = (
        "Seeds the soft-coded access control system: creates a Role row per "
        "legacy UserProfile role, a Permission row per nav module, and grants "
        "them to roles exactly matching the current hardcoded nav access, so "
        "behavior is unchanged until an admin edits it from Settings."
    )

    @transaction.atomic
    def handle(self, *args, **options):
        role_by_slug = {}
        for slug, label in UserProfile.ROLE_CHOICES:
            role, created = Role.objects.get_or_create(
                slug=slug,
                defaults={'display_name': label.split('—')[0].strip(), 'description': label, 'is_system': True},
            )
            role_by_slug[slug] = role
            self.stdout.write(f"{'Created' if created else 'Found'} role: {slug}")

        permission_by_key = {}
        for key, (name, description) in NAV_PERMISSION_INFO.items():
            code = nav.NAV_PERMISSION_CODES[key]
            permission, created = Permission.objects.get_or_create(
                code=code,
                defaults={'name': name, 'description': description, 'category': 'Navigation & Modules'},
            )
            permission_by_key[key] = permission
            self.stdout.write(f"{'Created' if created else 'Found'} permission: {code}")

        for key, allowed_roles in NAV_DEFAULT_ROLES.items():
            permission = permission_by_key[key]
            for slug in allowed_roles:
                role = role_by_slug.get(slug)
                if role is None:
                    continue
                role.permissions.add(permission)

        action_permission_by_key = {}
        for key, (name, description, category, allowed_roles) in ACTION_PERMISSIONS.items():
            code = f'action.{key}'
            permission, created = Permission.objects.get_or_create(
                code=code,
                defaults={'name': name, 'description': description, 'category': category},
            )
            action_permission_by_key[key] = (permission, allowed_roles)
            self.stdout.write(f"{'Created' if created else 'Found'} permission: {code}")

        for key, (permission, allowed_roles) in action_permission_by_key.items():
            for slug in allowed_roles:
                role = role_by_slug.get(slug)
                if role is None:
                    continue
                role.permissions.add(permission)

        self.stdout.write(self.style.SUCCESS('Access control seed complete.'))
