import os, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from django.test import RequestFactory
from django.contrib.auth.models import User
from django.template.loader import render_to_string

# Create a mock user and request
factory = RequestFactory()
user = User(username='testuser', is_staff=True)
request = factory.get('/')
request.user = user

templates_to_check = [
    'planning/planning_welcome.html',
    'planning/planning_home.html',
    'planning/approval_queue.html',
    'planning/planning_job_detail.html',
    'planning/planning_job_edit.html',
    'planning/po_inbox.html',
    'planning/pending_skus.html',
    'planning/planning_report.html',
    'planning/planning_scan.html',
]

class MockJob:
    def __init__(self):
        self.id = 1
        self.job_id = '123'
        self.card_number = 'C123'
        self.customer_name = 'Test'
        self.product_name = 'Test Product'
        self.status = 'Planning'
    def __str__(self):
        return str(self.id)

errors = []
for tpl in templates_to_check:
    ctx = {
        'request': request,
        'user': user,
        'user_role': 'admin',
        'can_edit_jobcard': True,
        'can_view_reports': True,
        'can_manage_masters': True,
        'planning_jobs': [],
        'qc_jobs': [],
        'pm_jobs': [],
        'release_jobs': [],
        'jobs': type('Page', (), {'paginator': type('P', (), {'count': 0, 'num_pages': 1})(), 'number': 1, 'has_other_pages': False, '__iter__': lambda self: iter([]), 'has_previous': False, 'has_next': False})(),
        'job': MockJob(),
        'filters': {'q': '', 'status': '', 'department': '', 'machine': '', 'from_date': '', 'to_date': ''},
        'status_choices': [],
        'status_counts': {},
        'rows': [],
        'pending_rows': [],
        'pending_count': 0,
        'po_summary': [],
        'po_filter': '',
        'q': '',
        'totals': {},
        'by_status': [],
        'by_department': [],
        'messages': [],
        'current_step': 'planning',
    }
    try:
        render_to_string(tpl, ctx, request=request)
        print(f'OK: {tpl}')
    except Exception as e:
        errors.append(f'ERROR {tpl}: {e}')
        print(f'ERROR: {tpl}: {e}')

if not errors:
    print('All templates rendered successfully.')
else:
    print(f'{len(errors)} template(s) had errors.')
