import re
from decimal import Decimal

from django import forms

from .models import JobCardLayout, PLANNING_QC_GATE_STATUSES, PlanningJob, PURCHASE_MATERIAL_ORIGIN_CHOICES, SkuRecipe


_COLOR_PLUS_RE = re.compile(r'^(\d+)\s*\+\s*(\d+)$')
_COLOR_SINGLE_RE = re.compile(r'^(\d+)\s*(?:colou?r(?:s)?)?$', re.IGNORECASE)


def _normalize_color_spec_value(raw_value):
    from core.print_colors import resolve_print_color_name

    raw_text = str(raw_value or '').strip()
    if not raw_text:
        return ''

    resolved = resolve_print_color_name(raw_text)
    if resolved:
        return resolved

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
        candidate = f"{int(plus_match.group(1))}+{int(plus_match.group(2))}"
        return resolve_print_color_name(candidate) or candidate

    single_match = _COLOR_SINGLE_RE.fullmatch(normalized)
    if single_match:
        candidate = str(int(single_match.group(1)))
        return resolve_print_color_name(candidate) or candidate

    numbers = re.findall(r'[0-9]+', normalized)
    if len(numbers) == 1:
        candidate = str(int(numbers[0]))
        return resolve_print_color_name(candidate) or candidate
    if len(numbers) == 2:
        candidate = f"{int(numbers[0])}+{int(numbers[1])}"
        return resolve_print_color_name(candidate) or candidate

    return raw_text


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

PRINT_PASS_CHOICES = [
    ('', 'Select Passes'),
    (1, '1 Pass'),
    (2, '2 Passes'),
    (3, '3 Passes'),
]


class PlanningJobFinalizationForm(forms.ModelForm):
    """Form to collect layout specs, purchase material, and release constraints
    during job finalization stage.
    """

    class Meta:
        model = PlanningJob
        fields = [
            'delivery_date',
            'wastage_sheets',
            'purchase_material_origin',
            'destination',
            'remarks',
            'requirement',
        ]
        widgets = {
            'delivery_date': forms.DateInput(attrs={'type': 'date'}),
            'remarks': forms.Textarea(attrs={'rows': 3}),
            'requirement': forms.Textarea(attrs={'rows': 3}),
        }
        labels = {
            'requirement': 'Special Instructions',
            'purchase_material_origin': 'Purchase Material Origin',
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        required_fields = [
            'delivery_date',
            'wastage_sheets',
            'purchase_material_origin',
            'destination',
            'requirement',
        ]
        for field_name in required_fields:
            if field_name in self.fields:
                self.fields[field_name].required = True
                self.fields[field_name].widget.attrs.setdefault('required', 'required')

        if 'purchase_material_origin' in self.fields:
            self.fields['purchase_material_origin'].widget = forms.Select(
                choices=PURCHASE_MATERIAL_ORIGIN_CHOICES,
                attrs={'class': 'erp-select'},
            )

    def _is_cut_and_pack(self):
        return self.instance.is_cut_and_pack() if self.instance and self.instance.pk else False

    def clean(self):
        cleaned = super().clean()
        cut_and_pack = self._is_cut_and_pack()

        required_messages = {
            'delivery_date': 'Delivery Date is required.',
            'wastage_sheets': 'Wastage Sheets is required.',
            'purchase_material_origin': 'Purchase Material Origin is required.',
            'destination': 'Destination is required.',
            'requirement': 'Special Instructions is required.',
        }

        for field_name, message in required_messages.items():
            value = cleaned.get(field_name)
            if field_name == 'wastage_sheets':
                if value is None:
                    self.add_error(field_name, message)
            else:
                if not str(value or '').strip():
                    self.add_error(field_name, message)

        if cut_and_pack:
            cleaned['planned_total_impressions'] = None
            return cleaned

        preview_job = self.instance
        preview_wastage = cleaned.get('wastage_sheets')
        if preview_wastage is not None:
            preview_job.wastage_sheets = preview_wastage
        preview_job.sync_print_passes_from_sku_master()
        passes = preview_job.effective_print_passes
        if passes and not self.errors.get('wastage_sheets'):
            if preview_job.calculated_sheets_required is None:
                self.add_error(
                    None,
                    'Cannot calculate impressions until print sheets are available (order qty, UPS, and wastage).',
                )
            else:
                cleaned['planned_total_impressions'] = preview_job.calculated_planned_total_impressions

        status = (cleaned.get('status') or self.instance.status or '').strip().lower()
        if status in PLANNING_QC_GATE_STATUSES and not cut_and_pack:
            if not getattr(self.instance, 'plate_set_no', '').strip():
                self.add_error(None, 'Plate Set is required before QC approval.')

            qc_required_messages = {
                'wastage_sheets': 'Wastage is required before QC approval.',
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

            if not passes:
                self.add_error(
                    None,
                    'No. of Passes is required on SKU master before QC approval.',
                )

        return cleaned

    def save(self, commit=True):
        job = super().save(commit=False)
        job.sync_job_process_type_from_sku_master()
        job.sync_print_passes_from_sku_master()
        job.sync_planned_total_impressions()
        if commit:
            job.save()
        return job


# Backward-compatible alias so existing view imports keep working
PlanningJobEditForm = PlanningJobFinalizationForm


class SkuRecipeForm(forms.ModelForm):
    class Meta:
        model = SkuRecipe
        fields = [
            'sku',
            'job_name',
            'job_process_type',
            'print_passes',
            'material',
            'color_spec',
            'application',
            'product_type',
            'machine_name',
            'plate_set_no',
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
            'remarks',
        ]
        widgets = {
            'notes': forms.Textarea(attrs={'rows': 3}),
            'remarks': forms.Textarea(attrs={'rows': 3}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        app_field = self.fields['application']
        app_field.widget = forms.Select(choices=APPLICATION_CHOICES, attrs={'class': 'erp-select'})
        app_field.required = True
        app_field.widget.attrs.setdefault('required', 'required')

        from core.models import Machine, Material, ProductType
        from core.print_colors import get_print_color_choices

        def _current_field_value(field_name):
            if self.data and field_name in self.data:
                return (self.data.get(field_name) or '').strip()
            if self.initial and self.initial.get(field_name) not in (None, ''):
                return self.initial.get(field_name)
            if self.instance is not None:
                return getattr(self.instance, field_name, None)
            return None

        product_types = ProductType.objects.all().order_by('name')
        product_type_choices = [('', 'Select Product Type')] + [(item.name, item.name) for item in product_types]
        current_product_type = _current_field_value('product_type')
        if current_product_type and current_product_type not in [item.name for item in product_types]:
            product_type_choices.append((current_product_type, current_product_type))
        self.fields['product_type'].widget = forms.Select(
            choices=product_type_choices,
            attrs={'class': 'erp-select', 'style': 'flex: 1;'},
        )
        self.fields['product_type'].required = True
        self.fields['product_type'].widget.attrs.setdefault('required', 'required')

        materials = Material.objects.all().order_by('name')
        material_choices = [('', 'Select Material Type')] + [(item.name, item.name) for item in materials]
        current_material = _current_field_value('material')
        if current_material and current_material not in [item.name for item in materials]:
            material_choices.append((current_material, current_material))
        self.fields['material'].widget = forms.Select(
            choices=material_choices,
            attrs={'class': 'erp-select', 'style': 'flex: 1;'},
        )

        machines = Machine.objects.filter(is_active=True).order_by('name')
        machine_choices = [('', 'Select Machine')] + [(m.name, m.name) for m in machines]
        current_value = _current_field_value('machine_name')
        if current_value and current_value not in [m.name for m in machines]:
            machine_choices.append((current_value, current_value))
        self.fields['machine_name'].widget = forms.Select(choices=machine_choices, attrs={'class': 'erp-select', 'style': 'flex: 1;'})

        current_color = _current_field_value('color_spec')
        self.fields['color_spec'].widget = forms.Select(
            choices=get_print_color_choices(include_legacy=current_color),
            attrs={'class': 'erp-select', 'style': 'flex: 1;'},
        )
        self.fields['color_spec'].label = 'Print Color'
        self.fields['color_spec'].help_text = 'Production colour pattern (1, 2, 4, 1+1, …). Admin manages the list.'

        self.fields['job_name'].widget.attrs.setdefault('required', 'required')
        self.fields['material'].widget.attrs.setdefault('required', 'required')
        self.fields['application'].widget.attrs.setdefault('required', 'required')
        self.fields['machine_name'].widget.attrs.setdefault('required', 'required')
        self.fields['machine_name'].required = True
        self.fields['print_sheet_size'].widget.attrs.setdefault('required', 'required')
        self.fields['purchase_sheet_size'].widget.attrs.setdefault('required', 'required')
        self.fields['ups'].widget.attrs.setdefault('required', 'required')
        self.fields['purchase_sheet_ups'].widget.attrs.setdefault('required', 'required')
        self.fields['size_w_mm'].widget.attrs.setdefault('required', 'required')
        self.fields['size_h_mm'].widget.attrs.setdefault('required', 'required')
        self.fields['material'].label = 'Material Type'
        self.fields['awc_no'].label = 'AWC #'
        self.fields['awc_no'].help_text = (
            'Artwork code (letters and/or numbers). Unique per SKU/design. '
            'Stored as text, not a decimal.'
        )
        self.fields['awc_no'].widget = forms.TextInput(attrs={
            'inputmode': 'text',
            'autocomplete': 'off',
        })

        from planning.models import SkuRecipe as SkuRecipeModel
        self.fields['job_process_type'].widget = forms.Select(
            choices=SkuRecipeModel.JOB_PROCESS_TYPE_CHOICES,
            attrs={'class': 'erp-select', 'id': 'id_job_process_type'},
        )
        self.fields['job_process_type'].label = 'Job Process'
        self.fields['job_process_type'].help_text = (
            'Print + Pack needs print color and plates. Cut & Pack has no printing. '
            'To change after approval, use Reopen SKU (change management).'
        )
        self.fields['job_process_type'].required = True

        if 'print_passes' in self.fields:
            self.fields['print_passes'].widget = forms.Select(
                choices=PRINT_PASS_CHOICES,
                attrs={'class': 'erp-select', 'id': 'id_print_passes'},
            )
            self.fields['print_passes'].label = 'No. of Passes'
            self.fields['print_passes'].help_text = (
                'Press passes for this SKU (1, 2, or 3). Not used for Cut & Pack.'
            )

        # Print Color required only for Print + Pack (validated in clean()).
        process_value = ''
        if self.data:
            process_value = (self.data.get('job_process_type') or '').strip()
        elif self.instance and self.instance.pk:
            process_value = (self.instance.job_process_type or '').strip()
        cut_and_pack = process_value == 'cut_and_pack'
        if not cut_and_pack:
            self.fields['color_spec'].required = True
            self.fields['color_spec'].widget.attrs.setdefault('required', 'required')
            if 'print_passes' in self.fields:
                self.fields['print_passes'].required = True
                self.fields['print_passes'].widget.attrs.setdefault('required', 'required')
        else:
            self.fields['color_spec'].required = False
            self.fields['color_spec'].widget.attrs.pop('required', None)
            if 'print_passes' in self.fields:
                self.fields['print_passes'].required = False
                self.fields['print_passes'].widget = forms.HiddenInput()
                self.fields['print_passes'].widget.attrs.pop('required', None)

        if 'die_cutting' in self.fields:
            self.fields['die_cutting'].widget.attrs.setdefault('required', 'required')

    def clean(self):
        cleaned = super().clean()
        process = (cleaned.get('job_process_type') or 'print_and_pack').strip()
        # Respect field.required (early plate-making saves relax designer fields).
        if (
            process != 'cut_and_pack'
            and self.fields.get('color_spec')
            and self.fields['color_spec'].required
            and not (cleaned.get('color_spec') or '').strip()
        ):
            self.add_error('color_spec', 'Print Color is required for Print + Pack jobs.')
        if process == 'cut_and_pack':
            cleaned['print_passes'] = None
        elif self.fields.get('print_passes') and self.fields['print_passes'].required:
            if cleaned.get('print_passes') in (None, ''):
                self.add_error('print_passes', 'No. of Passes is required for Print + Pack jobs.')
        return cleaned

    def clean_print_passes(self):
        value = self.cleaned_data.get('print_passes')
        if value in (None, ''):
            return None
        passes = int(value)
        if passes not in {1, 2, 3}:
            raise forms.ValidationError('Select 1, 2, or 3 passes.')
        return passes

    def clean_color_spec(self):
        from core.print_colors import resolve_print_color_name

        value = str(self.cleaned_data.get('color_spec') or '').strip()
        if not value:
            return ''

        resolved = resolve_print_color_name(value)
        if resolved:
            return resolved

        current = (self.instance.color_spec if self.instance and self.instance.pk else '') or ''
        if current and value == current:
            return value

        raise forms.ValidationError('Select a print color from the master list (admin can add new values).')

    def clean_application(self):
        value = _normalize_application_value(self.cleaned_data.get('application'))
        if not value:
            return ''

        allowed = {'UV', 'Lamination Gloss', 'Lamination Matt', 'NO'}
        if value in allowed:
            return value
        raise forms.ValidationError('Select Application as UV, Lamination Gloss, Lamination Matt, or NO.')

    def clean_awc_no(self):
        from planning.services import get_awc_conflict_message, normalize_awc_no

        value = normalize_awc_no(self.cleaned_data.get('awc_no'))
        if not value:
            return ''

        sku = str(self.cleaned_data.get('sku') or getattr(self.instance, 'sku', '') or '').strip()
        conflict = get_awc_conflict_message(
            value,
            sku=sku,
            exclude_recipe_id=self.instance.pk if self.instance and self.instance.pk else None,
        )
        if conflict:
            raise forms.ValidationError(conflict)
        return value

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

    def save(self, commit=True):
        from core.print_colors import apply_print_color_to_sku_recipe

        recipe = super().save(commit=False)
        apply_print_color_to_sku_recipe(recipe)
        if commit:
            recipe.save()
        return recipe


class JobCardLayoutForm(forms.ModelForm):
    class Meta:
        model = JobCardLayout
        fields = ['name', 'layout', 'is_active']
        widgets = {
            'layout': forms.HiddenInput(),
        }
