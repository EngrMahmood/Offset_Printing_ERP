from django import forms
from django.forms import formset_factory, inlineformset_factory

from .models import (
    CANCEL_REASON_CHOICES, PRIORITY_CHOICES, FlexoJobCard, FlexoPlanningDispatchRun,
    FlexoPlanningJob, FlexoPrintingLogEntry, FlexoSkuRecipe, FlexoSlittingCuttingLogEntry,
)


class ErpFormMixin:
    """Apply the shared erp-* CSS classes and sensible number steps to every widget."""

    def _style_widgets(self):
        for name, field in self.fields.items():
            widget = field.widget
            if isinstance(widget, (forms.CheckboxInput,)):
                continue
            css = 'erp-select' if isinstance(widget, forms.Select) else 'erp-input'
            widget.attrs.setdefault('class', css)
            if isinstance(widget, forms.DateInput):
                widget.input_type = 'date'
                widget.format = '%Y-%m-%d'
            if isinstance(widget, forms.NumberInput) and isinstance(field, forms.DecimalField):
                places = field.decimal_places or 0
                widget.attrs.setdefault('step', '1' if places == 0 else str(10 ** -places).rstrip('0') if places < 7 else 'any')

    def _limit_machines(self):
        """Offer Flexo/cutting/generic machines only (never offset or digital presses),
        keeping an already-saved machine selectable so old records still render."""
        from django.db.models import Q

        from core.machine_scope import flexo_machines

        instance = getattr(self, 'instance', None)
        for name in ('machine', 'machine_name'):
            field = self.fields.get(name)
            if field is None or not hasattr(field, 'queryset'):
                continue
            allowed = flexo_machines(field.queryset.filter(is_active=True))
            current = getattr(instance, f'{name}_id', None) if instance is not None else None
            if current:
                allowed = field.queryset.filter(Q(pk__in=allowed.values('pk')) | Q(pk=current))
            field.queryset = allowed.order_by('name')

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._style_widgets()
        self._limit_machines()


class FlexoSkuRecipeForm(ErpFormMixin, forms.ModelForm):
    class Meta:
        model = FlexoSkuRecipe
        fields = [
            'sku', 'job_name', 'material', 'color_spec',
            'size_w_mm', 'size_h_mm', 'ups', 'roll_size', 'meter_per_roll', 'pcs_per_roll',
            'wastage_pct', 'unit_type', 'die_code', 'teeth',
            'machine', 'department', 'purchase_material_origin', 'default_unit_cost',
            'daily_demand', 'remarks',
        ]

    def clean_sku(self):
        sku = (self.cleaned_data.get('sku') or '').strip()
        clash = FlexoSkuRecipe.objects.filter(sku__iexact=sku)
        if self.instance.pk:
            clash = clash.exclude(pk=self.instance.pk)
        if clash.exists():
            raise forms.ValidationError('A Flexo SKU master with this SKU already exists.')
        return sku


class FlexoPlanningJobEditForm(ErpFormMixin, forms.ModelForm):
    """Operational fields only — spec fields (size, ups, die code ...) come from
    the approved SKU master and are changed there (Reopen SKU)."""

    class Meta:
        model = FlexoPlanningJob
        fields = [
            'po_number', 'po_received_date', 'delivery_date', 'unit_type', 'order_qty', 'stock_qty',
            'wastage_pct', 'machine_name', 'department', 'destination', 'priority',
            'purchase_material_origin', 'daily_demand', 'unit_cost', 'amount',
            'rejected_qty', 'roll_qty_received', 'actual_wastage_mtr', 'received_qty_mtr',
            'remarks',
        ]


class FlexoPlanningDispatchRunForm(ErpFormMixin, forms.ModelForm):
    class Meta:
        model = FlexoPlanningDispatchRun
        fields = ['dispatch_index', 'delivery_date', 'dc_no', 'delivered_qty', 'packing', 'site']


FlexoPlanningDispatchRunFormSet = inlineformset_factory(
    FlexoPlanningJob, FlexoPlanningDispatchRun, form=FlexoPlanningDispatchRunForm,
    extra=1, can_delete=True,
)


class ManualPoHeaderForm(ErpFormMixin, forms.Form):
    po_number = forms.CharField(max_length=80, label='PO / WO Number')
    po_received_date = forms.DateField(label='PO Received Date', widget=forms.DateInput())
    destination = forms.CharField(max_length=100, required=False, label='Delivery Location')
    priority = forms.ChoiceField(choices=PRIORITY_CHOICES, initial=1)


class ManualPoLineForm(ErpFormMixin, forms.Form):
    sku = forms.ModelChoiceField(
        queryset=FlexoSkuRecipe.objects.none(), required=False, label='SKU (approved)',
    )
    order_qty = forms.IntegerField(min_value=1, required=False, label='Order Qty (SKU master unit)')
    stock_qty = forms.IntegerField(min_value=0, required=False, label='Stock Qty (same unit)')
    delivery_date = forms.DateField(required=False, widget=forms.DateInput())

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['sku'].queryset = FlexoSkuRecipe.objects.filter(is_active=True, master_data_status='approved')

    def clean(self):
        data = super().clean()
        if data.get('sku') and not data.get('order_qty'):
            self.add_error('order_qty', 'Order qty (pcs) is required for this SKU.')
        if data.get('order_qty') and not data.get('sku'):
            self.add_error('sku', 'Select the SKU for this qty.')
        stock, qty = data.get('stock_qty') or 0, data.get('order_qty') or 0
        if stock > qty and qty:
            self.add_error('stock_qty', 'Stock qty cannot exceed the order qty.')
        return data


ManualPoLineFormSet = formset_factory(ManualPoLineForm, extra=5)


class CancelForm(ErpFormMixin, forms.Form):
    cancel_reason_code = forms.ChoiceField(choices=CANCEL_REASON_CHOICES)
    reason = forms.CharField(widget=forms.Textarea(attrs={'rows': 2}))


class FlexoJobCardForm(ErpFormMixin, forms.ModelForm):
    class Meta:
        model = FlexoJobCard
        fields = [
            'delivery_location', 'department', 'days_plan', 'daily_demand',
            'material_type', 'size_w_mm', 'size_h_mm',
            'print_roll_size', 'roll_meter_qty', 'ups', 'machine', 'wastage',
            'color', 'special_instructions', 'mtr_per_roll', 'die_code',
            'prepared_by', 'approved_by', 'stock_check', 'dispatched_by',
            'extra_production', 'actual_waste', 'inventory_controller', 'receiving_signature',
        ]
        widgets = {'special_instructions': forms.Textarea(attrs={'rows': 2})}


class FlexoPrintingLogEntryForm(ErpFormMixin, forms.ModelForm):
    class Meta:
        model = FlexoPrintingLogEntry
        fields = [
            'date', 'machine', 'operator', 'shift',
            'given_meter_qty', 'wastage_meters', 'finished_meter_qty', 'signature_name',
        ]


class FlexoSlittingCuttingLogEntryForm(ErpFormMixin, forms.ModelForm):
    class Meta:
        model = FlexoSlittingCuttingLogEntry
        fields = [
            'date', 'machine', 'operator', 'shift',
            'given_meter_qty', 'wastage_pcs', 'finished_pcs', 'signature_name',
        ]
