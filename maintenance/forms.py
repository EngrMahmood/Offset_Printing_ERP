from django import forms

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
        label='Is the machine running?',
        choices=[('no', "No — it's stopped"), ('yes', "Yes — still running, but needs attention")],
        widget=forms.RadioSelect(attrs={'class': 'erp-radio-group'}),
        initial='yes',
    )
    photo = forms.FileField(
        required=False, label='Photo (optional)',
        widget=forms.ClearableFileInput(attrs={'class': 'erp-input', 'accept': 'image/*'}),
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
        fields = ['machine', 'title', 'interval_type', 'interval_value', 'next_due_at', 'is_active']
        widgets = {
            'machine': forms.Select(attrs={'class': 'erp-select'}),
            'title': forms.TextInput(attrs={'class': 'erp-input'}),
            'interval_type': forms.Select(attrs={'class': 'erp-select'}),
            'interval_value': forms.NumberInput(attrs={'class': 'erp-input'}),
            'next_due_at': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['next_due_at'].required = False


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
