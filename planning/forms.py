import re
from decimal import Decimal

from django import forms

from .models import PlanningJob, SkuRecipe


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


PURCHASE_MATERIAL_ORIGIN_CHOICES = [
    ('', 'Select Origin'),
    ('Local', 'Local'),
    ('Imported', 'Imported'),
]

APPLICATION_CHOICES = [
    ('', 'Select Application'),
    ('UV', 'UV'),
    ('Lamination Gloss', 'Lamination Gloss'),
    ('Lamination Matt', 'Lamination Matt'),
    ('NO', 'NO'),
]


class PlanningJobEditForm(forms.ModelForm):
    class Meta:
        model = PlanningJob
        fields = [
            'plan_date',
            'po_number',
            'sku',
            'job_name',
            'material',
            'color_spec',
            'application',
            'order_qty',
            'print_sheets',
            'machine_name',
            'department',
            'destination',
            'unit_cost',
            'daily_demand',
            'remarks',
            'requirement',
            'status',
        ]
        widgets = {
            'plan_date': forms.DateInput(attrs={'type': 'date'}),
            'remarks': forms.Textarea(attrs={'rows': 3}),
            'requirement': forms.Textarea(attrs={'rows': 3}),
        }


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
            'purchase_material',
            'machine_name',
            'department',
            'default_unit_cost',
            'daily_demand',
            'awc_no',
            'plate_set_no',
            'die_cutting',
            'notes',
        ]
        widgets = {
            'notes': forms.Textarea(attrs={'rows': 3}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        field = self.fields['purchase_material']
        field.widget = forms.Select(choices=PURCHASE_MATERIAL_ORIGIN_CHOICES)
        field.required = True
        field.widget.attrs.setdefault('required', 'required')

        app_field = self.fields['application']
        app_field.widget = forms.Select(choices=APPLICATION_CHOICES)
        app_field.required = True
        app_field.widget.attrs.setdefault('required', 'required')

        self.fields['job_name'].widget.attrs.setdefault('required', 'required')
        self.fields['material'].widget.attrs.setdefault('required', 'required')
        self.fields['color_spec'].widget.attrs.setdefault('required', 'required')
        self.fields['application'].widget.attrs.setdefault('required', 'required')
        self.fields['machine_name'].widget.attrs.setdefault('required', 'required')
        self.fields['print_sheet_size'].widget.attrs.setdefault('required', 'required')
        self.fields['purchase_sheet_size'].widget.attrs.setdefault('required', 'required')
        self.fields['ups'].widget.attrs.setdefault('required', 'required')
        self.fields['purchase_sheet_ups'].widget.attrs.setdefault('required', 'required')
        self.fields['purchase_material'].widget.attrs.setdefault('required', 'required')
        self.fields['size_w_mm'].widget.attrs.setdefault('required', 'required')
        self.fields['size_h_mm'].widget.attrs.setdefault('required', 'required')

        self.fields['department'].required = True
        self.fields['department'].widget.attrs.setdefault('required', 'required')
        self.fields['color_spec'].widget.attrs.setdefault('placeholder', 'e.g. 4 color or 1+1')

    def clean_purchase_material(self):
        value = (self.cleaned_data.get('purchase_material') or '').strip()
        if not value:
            return ''

        lowered = value.lower()
        if lowered in {'local', 'imported'}:
            return value.title()
        raise forms.ValidationError('Select Purchase Material Origin as Local or Imported.')

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
