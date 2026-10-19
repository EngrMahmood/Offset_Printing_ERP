import os

from django import forms
from django.core.exceptions import ValidationError

ALLOWED_ATTACHMENT_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.webp', '.gif', '.pdf', '.mp4', '.mov', '.doc', '.docx', '.xls', '.xlsx'}
MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024


def validate_attachment(upload):
    """Reject oversized or unexpected file types before they are stored."""
    extension = os.path.splitext(upload.name or '')[1].lower()
    if extension not in ALLOWED_ATTACHMENT_EXTENSIONS:
        raise ValidationError('This file type is not allowed. Upload a photo, PDF, video or Office document.')
    if upload.size > MAX_ATTACHMENT_BYTES:
        raise ValidationError('The file is larger than 10 MB.')

from .models import (
    FaultCategory, MachineDowntime, MaintenanceAttachment, MaintenanceRecord, MaintenanceServiceJob,
    MaintenanceSparePart, PreventiveMaintenancePlan,
)


class MaintenanceRecordForm(forms.ModelForm):
    class Meta:
        model = MaintenanceRecord
        fields = [
            'machine', 'reported_date', 'fault_types', 'maintenance_type', 'execution_type', 'priority',
            'fault_description', 'proposed_solution', 'spare_parts_needed', 'repair_needed', 'repair_details',
            'assigned_to', 'labour_hours', 'remarks',
        ]
        widgets = {
            'machine': forms.Select(attrs={'class': 'erp-select'}),
            'reported_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'fault_types': forms.CheckboxSelectMultiple(attrs={'class': 'erp-checkbox-group'}),
            'maintenance_type': forms.Select(attrs={'class': 'erp-select'}),
            'execution_type': forms.Select(attrs={'class': 'erp-select'}),
            'priority': forms.Select(attrs={'class': 'erp-select'}),
            'fault_description': forms.Textarea(attrs={'class': 'erp-input', 'rows': 3}),
            'proposed_solution': forms.Textarea(attrs={'class': 'erp-input', 'rows': 2}),
            'repair_details': forms.Textarea(attrs={'class': 'erp-input', 'rows': 2}),
            'assigned_to': forms.Select(attrs={'class': 'erp-select'}),
            'labour_hours': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.5'}),
            'remarks': forms.Textarea(attrs={'class': 'erp-input', 'rows': 2}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['fault_types'].queryset = FaultCategory.objects.filter(is_active=True)
        self.fields['assigned_to'].required = False
        self.fields['proposed_solution'].required = False
        self.fields['repair_details'].required = False
        self.fields['labour_hours'].required = False
        self.fields['remarks'].required = False


class ComplaintForm(forms.ModelForm):
    """Minimal shop-floor intake: what machine, what's wrong, is it stopped.
    Nothing technical — the engineer classifies everything else at triage."""

    machine_stopped = forms.ChoiceField(
        label='Has the machine stopped?',
        choices=[('yes', 'Yes - production has stopped'), ('no', 'No - it is still running')],
        widget=forms.RadioSelect(attrs={'class': 'erp-radio-group'}),
        initial='no',
    )
    photo = forms.FileField(
        required=False, label='Photo (optional)',
        widget=forms.ClearableFileInput(attrs={'class': 'erp-input', 'accept': 'image/*'}),
        validators=[validate_attachment],
    )

    class Meta:
        model = MaintenanceRecord
        fields = ['machine', 'fault_description']
        labels = {
            'fault_description': "What's wrong with the machine?",
        }
        widgets = {
            'machine': forms.Select(attrs={'class': 'erp-select'}),
            'fault_description': forms.Textarea(attrs={'class': 'erp-input', 'rows': 4, 'placeholder': 'Describe the problem in your own words...'}),
        }


class TriageForm(forms.ModelForm):
    """The engineer's classification screen: turns a raw complaint into a
    properly assessed record and moves it from REPORTED to DIAGNOSED."""

    class Meta:
        model = MaintenanceRecord
        fields = [
            'fault_types', 'maintenance_type', 'priority', 'execution_type', 'proposed_solution',
            'spare_parts_needed', 'repair_needed', 'repair_details', 'assigned_to',
        ]
        widgets = {
            'fault_types': forms.CheckboxSelectMultiple(attrs={'class': 'erp-checkbox-group'}),
            'maintenance_type': forms.Select(attrs={'class': 'erp-select'}),
            'priority': forms.Select(attrs={'class': 'erp-select'}),
            'execution_type': forms.Select(attrs={'class': 'erp-select'}),
            'proposed_solution': forms.Textarea(attrs={'class': 'erp-input', 'rows': 2}),
            'repair_details': forms.Textarea(attrs={'class': 'erp-input', 'rows': 2}),
            'assigned_to': forms.Select(attrs={'class': 'erp-select'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        from django.contrib.auth import get_user_model

        User = get_user_model()
        self.fields['fault_types'].queryset = FaultCategory.objects.filter(is_active=True)
        self.fields['maintenance_type'].required = True
        self.fields['proposed_solution'].required = False
        self.fields['repair_details'].required = False
        self.fields['assigned_to'].required = False
        self.fields['assigned_to'].queryset = User.objects.filter(
            is_active=True, profile__role__in=('admin', 'manager', 'production_manager', 'maintenance_engineer'),
        )


class MaintenanceSparePartForm(forms.ModelForm):
    class Meta:
        model = MaintenanceSparePart
        fields = ['description', 'quantity', 'uom', 'existing_sku']
        widgets = {
            'description': forms.TextInput(attrs={'class': 'erp-input'}),
            'quantity': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.01'}),
            'uom': forms.TextInput(attrs={'class': 'erp-input'}),
            'existing_sku': forms.Select(attrs={'class': 'erp-select'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['existing_sku'].required = False


class MaintenanceServiceJobForm(forms.ModelForm):
    class Meta:
        model = MaintenanceServiceJob
        fields = ['vendor', 'scope', 'sent_out_date', 'returned_date']
        widgets = {
            'vendor': forms.TextInput(attrs={'class': 'erp-input'}),
            'scope': forms.Textarea(attrs={'class': 'erp-input', 'rows': 3}),
            'sent_out_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'returned_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['vendor'].required = False
        self.fields['sent_out_date'].required = False
        self.fields['returned_date'].required = False


class MachineDowntimeForm(forms.ModelForm):
    class Meta:
        model = MachineDowntime
        fields = ['start_at', 'end_at', 'reason']
        widgets = {
            'start_at': forms.DateTimeInput(attrs={'class': 'erp-input', 'type': 'datetime-local'}),
            'end_at': forms.DateTimeInput(attrs={'class': 'erp-input', 'type': 'datetime-local'}),
            'reason': forms.Select(attrs={'class': 'erp-select'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['end_at'].required = False


class PreventiveMaintenancePlanForm(forms.ModelForm):
    class Meta:
        model = PreventiveMaintenancePlan
        fields = [
            'machine', 'title', 'interval_type', 'interval_value', 'next_due_at', 'next_due_impressions',
            'assigned_to', 'instructions', 'checklist', 'is_active',
        ]
        widgets = {
            'machine': forms.Select(attrs={'class': 'erp-select'}),
            'title': forms.TextInput(attrs={'class': 'erp-input'}),
            'interval_type': forms.Select(attrs={'class': 'erp-select'}),
            'interval_value': forms.NumberInput(attrs={'class': 'erp-input', 'min': 1}),
            'next_due_at': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}, format='%Y-%m-%d'),
            'next_due_impressions': forms.NumberInput(attrs={'class': 'erp-input', 'min': 0}),
            'assigned_to': forms.Select(attrs={'class': 'erp-select'}),
            'instructions': forms.Textarea(attrs={'class': 'erp-input', 'rows': 3}),
            'checklist': forms.Textarea(attrs={'class': 'erp-input', 'rows': 6, 'placeholder': 'One check per line'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        from django.contrib.auth import get_user_model

        from .permissions import STAFF_ROLES

        self.fields['next_due_at'].required = False
        self.fields['next_due_impressions'].required = False
        self.fields['assigned_to'].required = False
        self.fields['assigned_to'].queryset = get_user_model().objects.filter(
            is_active=True, profile__role__in=STAFF_ROLES,
        )
        self.fields['interval_value'].min_value = 1

    def clean(self):
        cleaned = super().clean()
        interval_type = cleaned.get('interval_type')
        interval = cleaned.get('interval_value')
        if interval is not None and interval < 1:
            self.add_error('interval_value', 'The interval must be at least 1.')
        if interval_type == 'DAYS':
            if not cleaned.get('next_due_at') and interval:
                from datetime import timedelta

                from django.utils import timezone

                cleaned['next_due_at'] = timezone.localdate() + timedelta(days=interval)
            cleaned['next_due_impressions'] = None
        elif interval_type == 'IMPRESSIONS':
            machine = cleaned.get('machine')
            if cleaned.get('next_due_impressions') is None and machine and interval:
                from .services import machine_impression_total

                # Start counting from the machine's current logged total.
                cleaned['next_due_impressions'] = machine_impression_total(machine) + interval
            cleaned['next_due_at'] = None
        return cleaned


class MaintenanceAttachmentForm(forms.ModelForm):
    class Meta:
        model = MaintenanceAttachment
        fields = ['file', 'caption']
        widgets = {
            'file': forms.ClearableFileInput(attrs={'class': 'erp-input'}),
            'caption': forms.TextInput(attrs={'class': 'erp-input'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['caption'].required = False
        self.fields['file'].validators.append(validate_attachment)
