import re
from decimal import Decimal

from django import forms

from .models import JobCardLayout, PLANNING_QC_GATE_STATUSES, PlanningJob, PURCHASE_MATERIAL_ORIGIN_CHOICES, SkuRecipe


_COLOR_PLUS_RE = re.compile(r'^(\d+)\s*\+\s*(\d+)$')
_COLOR_SINGLE_RE = re.compile(r'^(\d+)\s*(?:colou?r(?:s)?)?$', re.IGNORECASE)


def _normalize_color_spec_value(raw_value):
    raw_text = str(raw_value or '').strip()
    if not raw_text:
        return ''

    lowered = raw_text.lower()
    if re.search(r'color|colour|colours|colors', lowered):
        usable = True
    elif re.search(r'\d+\s*c\b', lowered) or ('c' in lowered and re.search(r'\d', lowered)):
        usable = True
    elif any(sep in lowered for sep in ['+', '/', '-']):
        usable = True
    elif raw_text.isdigit() or re.fullmatch(r'\d+\.\d+', raw_text):
        usable = True
    else:
        usable = False

    if not usable:
        return raw_text

    normalized = lowered.replace('colours', 'color').replace('colour', 'color').replace('colors', 'color')
    normalized = normalized.replace('c/', '+').replace('c+', '+').replace('/', '+').replace('-', '+')
    normalized = re.sub(r'[^0-9\+\s]+', '', normalized).strip()
    normalized = re.sub(r'\s+', '+', normalized)
    normalized = re.sub(r'\++', '+', normalized)

    plus_match = _COLOR_PLUS_RE.fullmatch(normalized)
    if plus_match:
        return f"{int(plus_match.group(1))}+{int(plus_match.group(2))}"

    single_match = _COLOR_SINGLE_RE.fullmatch(normalized)
    if single_match:
        return f"{int(single_match.group(1))} color"

    numbers = re.findall(r'[0-9]+', normalized)
    if len(numbers) == 1:
        return f"{int(numbers[0])} color"
    if len(numbers) == 2:
        return f"{int(numbers[0])}+{int(numbers[1])}"

    return value


def _normalize_application_value(raw_value):
    value = str(raw_value or '').strip()
    if not value:
        return ''
    lowered = value.lower()
    if lowered in {'no', 'none', 'n/a', 'na', 'nil', 'not applicable'}:
        return 'NO'
    if 'uv' in lowered or 'u.v' in lowered:
        return 'UV'
    if 'matt' in lowered or 'matte' in lowered:
        return 'Lamination Matt'
    if 'lamination' in lowered or 'lam' in lowered or 'lamin' in lowered:
        return 'Lamination Gloss'
    if 'gloss' in lowered or 'shine' in lowered:
        return 'Lamination Gloss'
    if 'varnish' in lowered or 'op' in lowered:
        return 'NO'
    return value


APPLICATION_CHOICES = [
    ('', 'Select Application'),
    ('UV', 'UV'),
    ('Lamination Gloss', 'Lamination Gloss'),
    ('Lamination Matt', 'Lamination Matt'),
    ('NO', 'NO'),
]


class PlanningJobFinalizationForm(forms.ModelForm):
    """Phase 4 finalization form — editable fields only (pre-QC execution prep).

    Access rule: status must be 'draft' or 'pending_qc'.
    SKU master fields are read-only; changes require reopen_sku → re-approval.
    """

    class Meta:
        model = PlanningJob
        fields = [
            'delivery_date',
            'wastage_sheets',
            'plate_set_no',
            'machine_name',
            'planned_total_impressions',
            'purchase_material_origin',
            'destination',
            'remarks',
            'requirement',
            'status',
        ]
        widgets = {
            'delivery_date': forms.DateInput(attrs={'type': 'date'}),
            'remarks': forms.Textarea(attrs={'rows': 3}),
            'requirement': forms.Textarea(attrs={'rows': 3}),
        }
        labels = {
            'planned_total_impressions': 'Total Impressions',
            'requirement': 'Special Instructions',
            'purchase_material_origin': 'Purchase Material Origin',
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        required_fields = [
            'delivery_date',
            'wastage_sheets',
            'plate_set_no',
            'machine_name',
            'planned_total_impressions',
            'purchase_material_origin',
            'destination',
            'requirement',
            'status',
        ]
        for field_name in required_fields:
            if field_name in self.fields:
                self.fields[field_name].required = True
                self.fields[field_name].widget.attrs.setdefault('required', 'required')

        if 'purchase_material_origin' in self.fields:
            self.fields['purchase_material_origin'].widget = forms.Select(choices=PURCHASE_MATERIAL_ORIGIN_CHOICES)
        if 'plate_set_no' in self.fields:
            self.fields['plate_set_no'].widget.attrs.setdefault('placeholder', 'Plate set reference')

    def clean(self):
        cleaned = super().clean()
        required_messages = {
            'delivery_date': 'Delivery Date is required.',
            'wastage_sheets': 'Wastage Sheets is required.',
            'plate_set_no': 'Plate Set is required.',
            'machine_name': 'Machine Name is required.',
            'planned_total_impressions': 'Total Impressions is required.',
            'purchase_material_origin': 'Purchase Material Origin is required.',
            'destination': 'Destination is required.',
            'requirement': 'Special Instructions is required.',
            'status': 'Status is required.',
        }

        for field_name, message in required_messages.items():
            value = cleaned.get(field_name)
            if field_name == 'wastage_sheets':
                if value is None:
                    self.add_error(field_name, message)
            else:
                if not str(value or '').strip():
                    self.add_error(field_name, message)

        status = (cleaned.get('status') or self.instance.status or '').strip().lower()
        if status in PLANNING_QC_GATE_STATUSES:
            qc_required_messages = {
                'plate_set_no': 'Plate Set is required before QC approval.',
                'wastage_sheets': 'Wastage is required before QC approval.',
                'machine_name': 'Machine Name is required before QC approval.',
                'purchase_material_origin': 'Purchase Material Origin is required before QC approval.',
            }
            for field_name, message in qc_required_messages.items():
                value = cleaned.get(field_name)
                if field_name == 'wastage_sheets':
                    if value is None:
                        self.add_error(field_name, message)
                else:
                    if not str(value or '').strip():
                        self.add_error(field_name, message)

        return cleaned


# Backward-compatible alias so existing view imports keep working
PlanningJobEditForm = PlanningJobFinalizationForm


class SkuRecipeForm(forms.ModelForm):
    class Meta:
        model = SkuRecipe
        fields = [
            'sku',
            'job_name',
            'material',
            'color_spec',
            'application',
            'size_w_mm',
            'size_h_mm',
            'ups',
            'print_sheet_size',
            'purchase_sheet_size',
            'purchase_sheet_ups',
            'default_unit_cost',
            'daily_demand',
            'awc_no',
            'die_cutting',
            'notes',
        ]
        widgets = {
            'notes': forms.Textarea(attrs={'rows': 3}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        app_field = self.fields['application']
        app_field.widget = forms.Select(choices=APPLICATION_CHOICES)
        app_field.required = True
        app_field.widget.attrs.setdefault('required', 'required')

        self.fields['job_name'].widget.attrs.setdefault('required', 'required')
        self.fields['material'].widget.attrs.setdefault('required', 'required')
        self.fields['color_spec'].widget.attrs.setdefault('required', 'required')
        self.fields['application'].widget.attrs.setdefault('required', 'required')
        self.fields['print_sheet_size'].widget.attrs.setdefault('required', 'required')
        self.fields['purchase_sheet_size'].widget.attrs.setdefault('required', 'required')
        self.fields['ups'].widget.attrs.setdefault('required', 'required')
        self.fields['purchase_sheet_ups'].widget.attrs.setdefault('required', 'required')
        self.fields['size_w_mm'].widget.attrs.setdefault('required', 'required')
        self.fields['size_h_mm'].widget.attrs.setdefault('required', 'required')
        self.fields['material'].label = 'Material Type'

        self.fields['color_spec'].widget.attrs.setdefault('placeholder', 'e.g. 4 color or 1+1')
        if 'die_cutting' in self.fields:
            self.fields['die_cutting'].widget.attrs.setdefault('required', 'required')

    def clean_color_spec(self):
        value = _normalize_color_spec_value(self.cleaned_data.get('color_spec'))
        if not value:
            return ''

        plus_match = _COLOR_PLUS_RE.fullmatch(value)
        if plus_match:
            return f"{int(plus_match.group(1))}+{int(plus_match.group(2))}"

        single_match = _COLOR_SINGLE_RE.fullmatch(value)
        if single_match:
            return f"{int(single_match.group(1))} color"

        raise forms.ValidationError('Use color format like 4 color or 1+1.')

    def clean_application(self):
        value = _normalize_application_value(self.cleaned_data.get('application'))
        if not value:
            return ''

        allowed = {'UV', 'Lamination Gloss', 'Lamination Matt', 'NO'}
        if value in allowed:
            return value
        raise forms.ValidationError('Select Application as UV, Lamination Gloss, Lamination Matt, or NO.')

    def _normalize_decimal_field(self, value):
        if value is None:
            return None
        if isinstance(value, Decimal) and value == value.to_integral_value():
            return value.quantize(Decimal('1'))
        return value

    def clean_size_w_mm(self):
        return self._normalize_decimal_field(self.cleaned_data.get('size_w_mm'))

    def clean_size_h_mm(self):
        return self._normalize_decimal_field(self.cleaned_data.get('size_h_mm'))


class JobCardLayoutForm(forms.ModelForm):
    class Meta:
        model = JobCardLayout
        fields = ['name', 'layout', 'is_active']
        widgets = {
            'layout': forms.HiddenInput(),
        }
