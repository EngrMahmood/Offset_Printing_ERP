from django import forms
from backup.models import BackupSetting

class BackupSettingForm(forms.ModelForm):
    class Meta:
        model = BackupSetting
        fields = [
            'backup_enabled',
            'backup_time',
            'frequency',
            'local_backup_folder',
            'cloud_onedrive_folder',
            'cloud_gdrive_folder',
            'keep_daily',
            'keep_weekly',
            'keep_monthly',
            'include_media',
            'include_logs',
            'enable_notifications',
            'enable_encryption',
            'encryption_password',
        ]
        widgets = {
            'backup_time': forms.TimeInput(attrs={'type': 'time', 'class': 'erp-input'}),
            'frequency': forms.Select(attrs={'class': 'erp-select'}),
            'local_backup_folder': forms.TextInput(attrs={'class': 'erp-input'}),
            'cloud_onedrive_folder': forms.TextInput(attrs={'class': 'erp-input'}),
            'cloud_gdrive_folder': forms.TextInput(attrs={'class': 'erp-input'}),
            'keep_daily': forms.NumberInput(attrs={'class': 'erp-input', 'min': 1}),
            'keep_weekly': forms.NumberInput(attrs={'class': 'erp-input', 'min': 1}),
            'keep_monthly': forms.NumberInput(attrs={'class': 'erp-input', 'min': 1}),
            'encryption_password': forms.PasswordInput(render_value=True, attrs={'class': 'erp-input'}),
        }
