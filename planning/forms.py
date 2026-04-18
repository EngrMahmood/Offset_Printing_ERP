import re

from django import forms

from .models import PlanningJob, SkuRecipe


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

_COLOR_PLUS_RE = re.compile(r'^(\d+)\s*\+\s*(\d+)$')
_COLOR_SINGLE_RE = re.compile(r'^(\d+)\s*(?:colou?r(?:s)?)?$', re.IGNORECASE)


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
        field.required = False

        app_field = self.fields['application']
        app_field.widget = forms.Select(choices=APPLICATION_CHOICES)
        app_field.required = False

        self.fields['department'].required = False
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
        value = (self.cleaned_data.get('color_spec') or '').strip()
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
        value = (self.cleaned_data.get('application') or '').strip()
        if not value:
            return ''

        allowed = {'UV', 'Lamination Gloss', 'Lamination Matt', 'NO'}
        if value in allowed:
            return value
        raise forms.ValidationError('Select Application as UV, Lamination Gloss, Lamination Matt, or NO.')
