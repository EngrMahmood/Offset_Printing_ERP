from django import forms

from core.models import Material

from .models import PhysicalStockCount, RawMaterialSku, StockDemand, StockTransaction


class RawMaterialSkuForm(forms.ModelForm):
    class Meta:
        model = RawMaterialSku
        fields = [
            'sku',
            'material',
            'purchase_sheet_size',
            'uom',
            'sheet_packing_pcs',
            'unit_cost',
            'safety_stock',
            'max_stock_level',
            'lead_time_days',
            'is_active',
        ]
        widgets = {
            'sku': forms.TextInput(attrs={'class': 'erp-input'}),
            'material': forms.Select(attrs={'class': 'erp-select'}),
            'purchase_sheet_size': forms.TextInput(attrs={'class': 'erp-input'}),
            'uom': forms.TextInput(attrs={'class': 'erp-input'}),
            'sheet_packing_pcs': forms.NumberInput(attrs={'class': 'erp-input'}),
            'unit_cost': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.01'}),
            'safety_stock': forms.NumberInput(attrs={'class': 'erp-input'}),
            'max_stock_level': forms.NumberInput(attrs={'class': 'erp-input'}),
            'lead_time_days': forms.NumberInput(attrs={'class': 'erp-input'}),
            'is_active': forms.CheckboxInput(attrs={'class': 'erp-checkbox'}),
        }


class SupplyChainItemForm(RawMaterialSkuForm):
    """Backward-compatible alias."""


class StockTransactionForm(forms.ModelForm):
    class Meta:
        model = StockTransaction
        fields = [
            'raw_material_sku',
            'month_str',
            'date',
            'gin_jc',
            'sheet_qty_pcs',
            'pkt_rim_qty',
        ]
        widgets = {
            'raw_material_sku': forms.Select(attrs={'class': 'erp-select'}),
            'month_str': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'e.g. June 2026'}),
            'date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'gin_jc': forms.TextInput(attrs={'class': 'erp-input'}),
            'sheet_qty_pcs': forms.NumberInput(attrs={'class': 'erp-input'}),
            'pkt_rim_qty': forms.NumberInput(attrs={'class': 'erp-input'}),
        }


class StockDemandForm(forms.ModelForm):
    class Meta:
        model = StockDemand
        fields = ['raw_material_sku', 'month_str', 'sheet_qty_pcs', 'pkt_rim_qty']
        widgets = {
            'raw_material_sku': forms.Select(attrs={'class': 'erp-select'}),
            'month_str': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'e.g. June 2026'}),
            'sheet_qty_pcs': forms.NumberInput(attrs={'class': 'erp-input'}),
            'pkt_rim_qty': forms.NumberInput(attrs={'class': 'erp-input'}),
        }


class ExcelUploadForm(forms.Form):
    upload_file = forms.FileField(
        label='Excel file',
        widget=forms.ClearableFileInput(attrs={'class': 'erp-input', 'accept': '.xlsx,.csv'}),
    )


class PhysicalStockCountForm(forms.ModelForm):
    class Meta:
        model = PhysicalStockCount
        fields = ['raw_material_sku', 'count_date', 'physical_sheet_qty', 'physical_pkt_rim_qty', 'notes']
        widgets = {
            'raw_material_sku': forms.Select(attrs={'class': 'erp-select'}),
            'count_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'physical_sheet_qty': forms.NumberInput(attrs={'class': 'erp-input'}),
            'physical_pkt_rim_qty': forms.NumberInput(attrs={'class': 'erp-input'}),
            'notes': forms.TextInput(attrs={'class': 'erp-input'}),
        }


class BulkPhysicalCountForm(forms.Form):
    count_date = forms.DateField(
        widget=forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
    )


class QuickRawMaterialSkuForm(forms.Form):
    sku = forms.CharField(max_length=50, widget=forms.TextInput(attrs={'class': 'erp-input'}))
    material_name = forms.CharField(max_length=100, widget=forms.TextInput(attrs={'class': 'erp-input'}))
    purchase_sheet_size = forms.CharField(max_length=80, widget=forms.TextInput(attrs={'class': 'erp-input'}))

    def clean_material_name(self):
        name = (self.cleaned_data.get('material_name') or '').strip()
        if not name:
            raise forms.ValidationError('Material name is required.')
        return name
