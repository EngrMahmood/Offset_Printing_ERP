import os
import datetime
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.contrib.auth.decorators import login_required, user_passes_test
from django.http import FileResponse, Http404, HttpResponseForbidden
from django.utils import timezone
from django.db.models import Sum, Count, Q

from backup.models import BackupSetting, BackupHistory, RestoreHistory
from backup.forms import BackupSettingForm
from backup.services import create_backup, get_database_engine
from backup.restore import simulate_restore, perform_restore

def is_admin_user(user):
    """Verifies that the user has admin rights (superuser or profile role == admin)."""
    if not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    # Fallback to profile check if available
    profile = getattr(user, 'profile', None)
    if profile and getattr(profile, 'role', '') == 'admin':
        return True
    return False

def get_next_backup_time(setting):
    if not setting.backup_enabled:
        return None
    now = timezone.localtime(timezone.now())
    scheduled_time = setting.backup_time
    
    # Today's scheduled backup time
    scheduled_today = now.replace(hour=scheduled_time.hour, minute=scheduled_time.minute, second=0, microsecond=0)
    
    if setting.frequency == 'DAILY':
        if now < scheduled_today:
            return scheduled_today
        else:
            return scheduled_today + datetime.timedelta(days=1)
    elif setting.frequency == 'WEEKLY':
        # Sunday is 6
        days_ahead = (6 - now.weekday()) % 7
        if days_ahead == 0 and now >= scheduled_today:
            days_ahead = 7
        next_date = scheduled_today + datetime.timedelta(days=days_ahead)
        return next_date
    elif setting.frequency == 'MONTHLY':
        if now.day == 1 and now < scheduled_today:
            return scheduled_today
        if now.month == 12:
            next_month = now.replace(year=now.year + 1, month=1, day=1, hour=scheduled_time.hour, minute=scheduled_time.minute, second=0, microsecond=0)
        else:
            next_month = now.replace(month=now.month + 1, day=1, hour=scheduled_time.hour, minute=scheduled_time.minute, second=0, microsecond=0)
        return next_month
    return None

@login_required
@user_passes_test(is_admin_user)
def backup_dashboard(request):
    setting = BackupSetting.get_settings()
    
    # Backup stats
    backups_qs = BackupHistory.objects.all()
    total_backups = backups_qs.count()
    success_backups = backups_qs.filter(status='SUCCESS')
    success_count = success_backups.count()
    
    success_rate = 100.0
    if total_backups > 0:
        success_rate = round((success_count / total_backups) * 100, 1)
        
    storage_used_bytes = success_backups.aggregate(total_size=Sum('file_size'))['total_size'] or 0
    # Human readable size
    storage_used_mb = round(storage_used_bytes / (1024 * 1024), 2)
    
    last_backup = success_backups.first()
    next_backup = get_next_backup_time(setting)
    
    last_restore_record = RestoreHistory.objects.filter(status='SUCCESS').first()
    last_restore_date = last_restore_record.timestamp if last_restore_record else None
    
    # Recent errors
    recent_errors = BackupHistory.objects.filter(status='FAILED')[:5]
    
    # Recent backups history list
    history_list = backups_qs.order_by('-start_time')[:30]
    
    context = {
        'setting': setting,
        'last_backup': last_backup,
        'next_backup': next_backup,
        'total_backups': total_backups,
        'storage_used_mb': storage_used_mb,
        'last_restore_date': last_restore_date,
        'success_rate': success_rate,
        'recent_errors': recent_errors,
        'history_list': history_list,
        'db_engine': get_database_engine().upper(),
    }
    return render(request, 'backup/dashboard.html', context)

@login_required
@user_passes_test(is_admin_user)
def run_manual_backup(request):
    if request.method == 'POST':
        history = create_backup(backup_type='MANUAL', user=request.user)
        if history.status == 'SUCCESS':
            messages.success(request, f"Backup created successfully: {history.file_name}")
        else:
            messages.error(request, f"Backup failed: {history.error_message}")
    return redirect('backup:dashboard')

@login_required
@user_passes_test(is_admin_user)
def update_settings(request):
    setting = BackupSetting.get_settings()
    if request.method == 'POST':
        form = BackupSettingForm(request.POST, instance=setting)
        if form.is_valid():
            form.save()
            messages.success(request, "Backup settings updated successfully.")
            return redirect('backup:dashboard')
    else:
        form = BackupSettingForm(instance=setting)
        
    return render(request, 'backup/settings.html', {'form': form, 'setting': setting})

@login_required
@user_passes_test(is_admin_user)
def restore_backup(request, backup_id):
    backup = get_object_or_404(BackupHistory, id=backup_id)
    
    if request.method == 'POST':
        confirm = request.POST.get('confirm_restore')
        if confirm == 'RESTORE':
            try:
                perform_restore(backup.id, user=request.user)
                messages.success(request, "Database restored successfully. The ERP has been reverted to this backup.")
            except Exception as e:
                messages.error(request, f"Database restoration failed: {str(e)}")
            return redirect('backup:dashboard')
        else:
            messages.error(request, "Restoration cancelled. Please type 'RESTORE' to confirm.")
            return redirect('backup:dashboard')
            
    # Verify simulation
    sim = simulate_restore(backup.id)
    context = {
        'backup': backup,
        'sim': sim,
    }
    return render(request, 'backup/confirm_restore.html', context)

@login_required
@user_passes_test(is_admin_user)
def download_backup(request, backup_id):
    backup = get_object_or_404(BackupHistory, id=backup_id)
    if not backup.backup_location:
        raise Http404("No backup file path defined.")
        
    paths = [p.strip() for p in backup.backup_location.split(",")]
    zip_path = None
    for p in paths:
        if os.path.exists(p):
            zip_path = p
            break
            
    if not zip_path or not os.path.exists(zip_path):
        raise Http404("Physical backup file does not exist on disk.")
        
    response = FileResponse(open(zip_path, 'rb'), as_attachment=True, filename=backup.file_name)
    return response

@login_required
@user_passes_test(is_admin_user)
def delete_backup(request, backup_id):
    backup = get_object_or_404(BackupHistory, id=backup_id)
    if request.method == 'POST':
        # Remove physical files
        if backup.backup_location:
            paths = [p.strip() for p in backup.backup_location.split(",")]
            for p in paths:
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception as e:
                        messages.warning(request, f"Could not delete physical file: {str(e)}")
        backup.delete()
        messages.success(request, "Backup history and file deleted successfully.")
    return redirect('backup:dashboard')
