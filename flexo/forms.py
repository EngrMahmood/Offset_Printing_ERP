from django import forms
from django.forms import inlineformset_factory

from .models import (
    FlexoJobCard, FlexoPlanningDispatchRun, FlexoPlanningJob,
    FlexoPrintingLogEntry, FlexoSlittingCuttingLogEntry,
)

_INPUT = {'class': 'erp-input'}
_SELECT = {'class': 'erp-select'}
_DATE = {'class': 'erp-input', 'type': 'date'}


class FlexoPlanningJobForm(forms.ModelForm):
    class Meta:
        model = FlexoPlanningJob
        fields = [
            'po_number', 'po_received_date', 'plan_month', 'sku', 'job_name',
            'repeat_flag', 'material', 'color_spec',
            'size_w_mm', 'size_h_mm', 'die_code',
            'size_h_inches', 'awc_no', 'teeth', 'order_roll_qty', 'printing_cylinder',
            'order_qty', 'ups', 'roll_size', 'meter_per_roll', 'roll_qty_meter',
            'wastage_qty_meter', 'actual_roll_qty_meter', 'remarks',
            'roll_qty_received', 'actual_wastage_mtr', 'received_qty_mtr',
            'status', 'rejected_qty', 'balance_qty', 'destination', 'cost', 'amount',
            'machine_name', 'purchase_material_origin', 'stock_qty', 'daily_demand',
            'department',
        ]
        widgets = {
            'po_number': forms.TextInput(attrs=_INPUT),
            'po_received_date': forms.DateInput(attrs=_DATE),
            'plan_month': forms.TextInput(attrs=_INPUT),
            'sku': forms.TextInput(attrs=_INPUT),
            'job_name': forms.TextInput(attrs=_INPUT),
            'repeat_flag': forms.Select(attrs=_SELECT),
            'material': forms.TextInput(attrs=_INPUT),
            'color_spec': forms.TextInput(attrs=_INPUT),
            'size_w_mm': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'size_h_mm': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'die_code': forms.TextInput(attrs=_INPUT),
            'size_h_inches': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'awc_no': forms.TextInput(attrs=_INPUT),
            'teeth': forms.TextInput(attrs=_INPUT),
            'order_roll_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'printing_cylinder': forms.TextInput(attrs=_INPUT),
            'order_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'ups': forms.NumberInput(attrs={**_INPUT, 'step': '0.01'}),
            'roll_size': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'meter_per_roll': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'roll_qty_meter': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'wastage_qty_meter': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'actual_roll_qty_meter': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'remarks': forms.TextInput(attrs=_INPUT),
            'roll_qty_received': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'actual_wastage_mtr': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'received_qty_mtr': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'status': forms.Select(attrs=_SELECT),
            'rejected_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'balance_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'destination': forms.TextInput(attrs=_INPUT),
            'cost': forms.NumberInput(attrs={**_INPUT, 'step': '0.0001'}),
            'amount': forms.NumberInput(attrs={**_INPUT, 'step': '0.01'}),
            'machine_name': forms.Select(attrs=_SELECT),
            'purchase_material_origin': forms.Select(attrs=_SELECT),
            'stock_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'daily_demand': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'department': forms.Select(attrs=_SELECT),
        }


FlexoPlanningDispatchRunFormSet = inlineformset_factory(
    FlexoPlanningJob, FlexoPlanningDispatchRun,
    fields=['dispatch_index', 'delivery_date', 'dc_no', 'delivered_qty'],
    widgets={
        'dispatch_index': forms.NumberInput(attrs=_INPUT),
        'delivery_date': forms.DateInput(attrs=_DATE),
        'dc_no': forms.TextInput(attrs=_INPUT),
        'delivered_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
    },
    extra=1, can_delete=True,
)


class FlexoJobCardForm(forms.ModelForm):
    class Meta:
        model = FlexoJobCard
        fields = [
            'delivery_location', 'department', 'days_plan', 'daily_demand',
            'material_type', 'size_w_mm', 'size_h_mm', 'size_h_inches',
            'print_roll_size', 'roll_meter_qty', 'ups', 'machine', 'wastage',
            'color', 'special_instructions', 'mtr_per_roll',
            'die_code', 'awc_no', 'printing_cylinder',
            'prepared_by', 'checked_by', 'artwork_developed_by', 'approved_by',
            'status',
        ]
        widgets = {
            'delivery_location': forms.TextInput(attrs=_INPUT),
            'department': forms.Select(attrs=_SELECT),
            'days_plan': forms.NumberInput(attrs=_INPUT),
            'daily_demand': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'material_type': forms.TextInput(attrs=_INPUT),
            'size_w_mm': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'size_h_mm': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'size_h_inches': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'print_roll_size': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'roll_meter_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'ups': forms.NumberInput(attrs={**_INPUT, 'step': '0.01'}),
            'machine': forms.Select(attrs=_SELECT),
            'wastage': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'color': forms.TextInput(attrs=_INPUT),
            'special_instructions': forms.Textarea(attrs={**_INPUT, 'rows': 2}),
            'mtr_per_roll': forms.TextInput(attrs=_INPUT),
            'die_code': forms.TextInput(attrs=_INPUT),
            'awc_no': forms.TextInput(attrs=_INPUT),
            'printing_cylinder': forms.TextInput(attrs=_INPUT),
            'prepared_by': forms.TextInput(attrs=_INPUT),
            'checked_by': forms.TextInput(attrs=_INPUT),
            'artwork_developed_by': forms.TextInput(attrs=_INPUT),
            'approved_by': forms.TextInput(attrs=_INPUT),
            'status': forms.Select(attrs=_SELECT),
        }


class FlexoPrintingLogEntryForm(forms.ModelForm):
    class Meta:
        model = FlexoPrintingLogEntry
        fields = [
            'date', 'machine', 'operator', 'shift',
            'jumbo_roll_qty', 'baby_roll_qty', 'given_meter_qty', 'finished_meter_qty',
            'wastage', 'signature_name',
        ]
        widgets = {
            'date': forms.DateInput(attrs=_DATE),
            'machine': forms.Select(attrs=_SELECT),
            'operator': forms.Select(attrs=_SELECT),
            'shift': forms.Select(attrs=_SELECT),
            'jumbo_roll_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'baby_roll_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'given_meter_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'finished_meter_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'wastage': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'signature_name': forms.TextInput(attrs=_INPUT),
        }


class FlexoSlittingCuttingLogEntryForm(forms.ModelForm):
    class Meta:
        model = FlexoSlittingCuttingLogEntry
        fields = ['date', 'operator', 'cut_qty', 'cut_length_m', 'wastage', 'signature_name']
        widgets = {
            'date': forms.DateInput(attrs=_DATE),
            'operator': forms.Select(attrs=_SELECT),
            'cut_qty': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'cut_length_m': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'wastage': forms.NumberInput(attrs={**_INPUT, 'step': '0.001'}),
            'signature_name': forms.TextInput(attrs=_INPUT),
        }
