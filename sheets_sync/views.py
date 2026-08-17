from django.contrib.auth.decorators import login_required, user_passes_test
from django.shortcuts import render

from sheets_sync.models import SheetsSyncSetting, SheetsSyncLog, SheetsRowIndex
from sheets_sync.registry import SYNCED_MODELS
from sheets_sync.queue_worker import get_worker_status


def is_admin_user(user):
    if not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    profile = getattr(user, 'profile', None)
    if profile and getattr(profile, 'role', '') == 'admin':
        return True
    return False


@login_required
@user_passes_test(is_admin_user)
def dashboard(request):
    setting = SheetsSyncSetting.get_settings()
    worker_status = get_worker_status()

    tabs = []
    for entry in SYNCED_MODELS:
        tab_name = entry['tab_name']
        last_log = SheetsSyncLog.objects.filter(tab_name=tab_name).order_by('-timestamp').first()
        tabs.append({
            'tab_name': tab_name,
            'dotted_path': entry['dotted_path'],
            'row_count': SheetsRowIndex.objects.filter(tab_name=tab_name).count(),
            'last_log': last_log,
        })

    recent_failures = SheetsSyncLog.objects.filter(status='FAILED').order_by('-timestamp')[:15]
    recent_activity = SheetsSyncLog.objects.order_by('-timestamp')[:30]

    spreadsheet_url = ''
    if setting.spreadsheet_id:
        spreadsheet_url = f'https://docs.google.com/spreadsheets/d/{setting.spreadsheet_id}/edit'

    context = {
        'setting': setting,
        'spreadsheet_url': spreadsheet_url,
        'worker_status': worker_status,
        'tabs': tabs,
        'recent_failures': recent_failures,
        'recent_activity': recent_activity,
        'total_rows': sum(t['row_count'] for t in tabs),
    }
    return render(request, 'sheets_sync/dashboard.html', context)
