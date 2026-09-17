from django import forms
from django.forms import inlineformset_factory

from .formula import FormulaError, validate_expression
from .models import (
    BomTemplate, BomTemplateCondition, BomTemplateLine, RawItem, SpecAttribute,
)


class RawItemForm(forms.ModelForm):
    gsm = forms.DecimalField(required=False, label='GSM', widget=forms.NumberInput(attrs={'class': 'erp-input'}))
    sheet_size = forms.CharField(required=False, label='Sheet Size', widget=forms.TextInput(attrs={'class': 'erp-input'}))

    class Meta:
        model = RawItem
        fields = [
            'item_code', 'name', 'category', 'uom', 'specification', 'brand_grade',
            'production_line', 'item_role', 'unit_cost', 'default_vendor',
            'pack_size', 'moq', 'safety_stock', 'max_stock_level', 'lead_time_days',
        ]
        widgets = {
            'item_code': forms.TextInput(attrs={'class': 'erp-input'}),
            'name': forms.TextInput(attrs={'class': 'erp-input'}),
            'category': forms.Select(attrs={'class': 'erp-select'}),
            'uom': forms.Select(attrs={'class': 'erp-select'}),
            'specification': forms.TextInput(attrs={'class': 'erp-input'}),
            'brand_grade': forms.TextInput(attrs={'class': 'erp-input'}),
            'production_line': forms.Select(attrs={'class': 'erp-select'}),
            'item_role': forms.TextInput(attrs={'class': 'erp-input'}),
            'unit_cost': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.000001'}),
            'default_vendor': forms.Select(attrs={'class': 'erp-select'}),
            'pack_size': forms.NumberInput(attrs={'class': 'erp-input'}),
            'moq': forms.NumberInput(attrs={'class': 'erp-input'}),
            'safety_stock': forms.NumberInput(attrs={'class': 'erp-input'}),
            'max_stock_level': forms.NumberInput(attrs={'class': 'erp-input'}),
            'lead_time_days': forms.NumberInput(attrs={'class': 'erp-input'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if self.instance and self.instance.pk:
            self.fields['gsm'].initial = self.instance.attributes.get('gsm')
            self.fields['sheet_size'].initial = self.instance.attributes.get('sheet_size')

    def save(self, commit=True):
        instance = super().save(commit=False)
        attributes = dict(instance.attributes or {})
        if self.cleaned_data.get('gsm') is not None:
            attributes['gsm'] = float(self.cleaned_data['gsm'])
        if self.cleaned_data.get('sheet_size'):
            attributes['sheet_size'] = self.cleaned_data['sheet_size']
        instance.attributes = attributes
        if commit:
            instance.save()
        return instance


class SpecAttributeForm(forms.ModelForm):
    class Meta:
        model = SpecAttribute
        fields = [
            'code', 'label', 'sort_order', 'is_active', 'is_match_key', 'is_wizard_step',
            'source_kind', 'source_field', 'pattern', 'capture_group', 'derived_key',
            'value_type', 'normalizer',
        ]
        widgets = {
            'code': forms.TextInput(attrs={'class': 'erp-input'}),
            'label': forms.TextInput(attrs={'class': 'erp-input'}),
            'sort_order': forms.NumberInput(attrs={'class': 'erp-input'}),
            'source_kind': forms.Select(attrs={'class': 'erp-select'}),
            'source_field': forms.Select(attrs={'class': 'erp-select'}),
            'pattern': forms.TextInput(attrs={'class': 'erp-input'}),
            'capture_group': forms.NumberInput(attrs={'class': 'erp-input'}),
            'derived_key': forms.Select(attrs={'class': 'erp-select'}),
            'value_type': forms.Select(attrs={'class': 'erp-select'}),
            'normalizer': forms.Select(attrs={'class': 'erp-select'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        from . import spec as spec_module
        self.fields['source_field'] = forms.ChoiceField(
            choices=[('', '---')] + [(f, f) for f in spec_module.available_sku_fields()],
            required=False, widget=forms.Select(attrs={'class': 'erp-select'}),
        )
        self.fields['derived_key'] = forms.ChoiceField(
            choices=[('', '---')] + [(k, k) for k in spec_module.DERIVED_RESOLVERS],
            required=False, widget=forms.Select(attrs={'class': 'erp-select'}),
        )


class BomTemplateForm(forms.ModelForm):
    class Meta:
        model = BomTemplate
        fields = ['code', 'name', 'description', 'production_line', 'priority']
        widgets = {
            'code': forms.TextInput(attrs={'class': 'erp-input'}),
            'name': forms.TextInput(attrs={'class': 'erp-input'}),
            'description': forms.Textarea(attrs={'class': 'erp-input', 'rows': 2}),
            'production_line': forms.Select(attrs={'class': 'erp-select'}),
            'priority': forms.NumberInput(attrs={'class': 'erp-input'}),
        }


class BomTemplateConditionForm(forms.ModelForm):
    class Meta:
        model = BomTemplateCondition
        fields = ['attribute', 'operator', 'value_text', 'value_min', 'value_max', 'value_list']
        widgets = {
            'attribute': forms.Select(attrs={'class': 'erp-select'}),
            'operator': forms.Select(attrs={'class': 'erp-select'}),
            'value_text': forms.TextInput(attrs={'class': 'erp-input'}),
            'value_min': forms.TextInput(attrs={'class': 'erp-input'}),
            'value_max': forms.TextInput(attrs={'class': 'erp-input'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['attribute'].queryset = SpecAttribute.objects.filter(is_active=True, is_match_key=True)
        self.fields['value_list'].required = False


BomTemplateConditionFormSet = inlineformset_factory(
    BomTemplate, BomTemplateCondition, form=BomTemplateConditionForm, extra=1, can_delete=True,
)


class BomTemplateLineForm(forms.ModelForm):
    class Meta:
        model = BomTemplateLine
        fields = [
            'line_code', 'sequence', 'process_step', 'bom_item_type', 'resolution_mode',
            'raw_item', 'item_role', 'selector', 'entry_mode', 'fixed_qty',
            'fixed_wastage_percent', 'quantity_basis', 'qty_expression',
            'driver_line', 'scale_by_colors', 'scale_by_passes', 'wastage_expression',
            'is_optional', 'notes',
        ]
        widgets = {
            'line_code': forms.TextInput(attrs={'class': 'erp-input'}),
            'sequence': forms.NumberInput(attrs={'class': 'erp-input'}),
            'process_step': forms.Select(attrs={'class': 'erp-select'}),
            'bom_item_type': forms.Select(attrs={'class': 'erp-select'}),
            'resolution_mode': forms.Select(attrs={'class': 'erp-select'}),
            'raw_item': forms.Select(attrs={'class': 'erp-select'}),
            'item_role': forms.TextInput(attrs={'class': 'erp-input'}),
            'entry_mode': forms.Select(attrs={'class': 'erp-select', 'data-entry-mode-toggle': '1'}),
            'fixed_qty': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.00000001', 'placeholder': 'e.g. 0.018 (kg per piece)'}),
            'fixed_wastage_percent': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.0001', 'placeholder': 'e.g. 0.05'}),
            'quantity_basis': forms.Select(attrs={'class': 'erp-select'}),
            'qty_expression': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'e.g. piece_area_sqm * gsm / 1000'}),
            'driver_line': forms.Select(attrs={'class': 'erp-select'}),
            'wastage_expression': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'e.g. 0.05'}),
            'notes': forms.TextInput(attrs={'class': 'erp-input'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if not self.instance.pk and not self.initial.get('quantity_basis'):
            from .models import QuantityBasis
            default_basis = QuantityBasis.objects.filter(code='per_1000_pcs', is_active=True).first()
            if default_basis:
                self.fields['quantity_basis'].initial = default_basis.pk

    def clean_qty_expression(self):
        expr = self.cleaned_data.get('qty_expression', '')
        if self.cleaned_data.get('entry_mode') != 'formula':
            return expr
        try:
            validate_expression(expr)
        except FormulaError as exc:
            raise forms.ValidationError(str(exc))
        return expr

    def clean_wastage_expression(self):
        expr = self.cleaned_data.get('wastage_expression') or '0'
        if self.cleaned_data.get('entry_mode') != 'formula':
            return expr
        try:
            validate_expression(expr)
        except FormulaError as exc:
            raise forms.ValidationError(str(exc))
        return expr

    def clean(self):
        cleaned = super().clean()
        entry_mode = cleaned.get('entry_mode')
        if entry_mode == 'simple' and cleaned.get('fixed_qty') is None:
            self.add_error('fixed_qty', 'Required when entry mode is "Simple".')
        if entry_mode == 'formula' and not cleaned.get('qty_expression'):
            self.add_error('qty_expression', 'Required when entry mode is "Formula".')
        return cleaned


BomTemplateLineFormSet = inlineformset_factory(
    BomTemplate, BomTemplateLine, form=BomTemplateLineForm, fk_name='template', extra=1, can_delete=True,
)


class TestMatchForm(forms.Form):
    sku = forms.CharField(
        label='SKU', widget=forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'Enter or pick a SKU code'}),
    )


class RaiseItemRequestForm(forms.Form):
    request_type = forms.ModelChoiceField(
        queryset=None, widget=forms.Select(attrs={'class': 'erp-select'}),
    )
    department = forms.ModelChoiceField(
        queryset=None, widget=forms.Select(attrs={'class': 'erp-select'}),
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        from supply_chain.models import ItemRequestDepartment, ItemRequestType
        self.fields['request_type'].queryset = ItemRequestType.objects.filter(is_active=True)
        self.fields['department'].queryset = ItemRequestDepartment.objects.filter(is_active=True)


class RawItemImportForm(forms.Form):
    upload = forms.FileField(widget=forms.ClearableFileInput(attrs={'class': 'erp-input', 'accept': '.xlsx'}))
