from django import forms

from .models import PhysicalStockCount, StockDemand, StockTransaction, SupplyChainItem


class SupplyChainItemForm(forms.ModelForm):
    class Meta:
        model = SupplyChainItem
        fields = [
            'item_id',
            'uom',
            'sheet_packing_pcs',
            'unit_cost',
            'safety_stock',
            'max_stock_level',
            'lead_time_days',
        ]
        widgets = {
            'item_id': forms.TextInput(attrs={'class': 'erp-input'}),
            'uom': forms.TextInput(attrs={'class': 'erp-input'}),
            'sheet_packing_pcs': forms.NumberInput(attrs={'class': 'erp-input'}),
            'unit_cost': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.01'}),
            'safety_stock': forms.NumberInput(attrs={'class': 'erp-input'}),
            'max_stock_level': forms.NumberInput(attrs={'class': 'erp-input'}),
            'lead_time_days': forms.NumberInput(attrs={'class': 'erp-input'}),
        }


class StockTransactionForm(forms.ModelForm):
    class Meta:
        model = StockTransaction
        fields = [
            'item',
            'month_str',
            'date',
            'gin_jc',
            'sheet_qty_pcs',
            'pkt_rim_qty',
        ]
        widgets = {
            'item': forms.Select(attrs={'class': 'erp-select'}),
            'month_str': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'e.g. June 2026'}),
            'date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'gin_jc': forms.TextInput(attrs={'class': 'erp-input'}),
            'sheet_qty_pcs': forms.NumberInput(attrs={'class': 'erp-input'}),
            'pkt_rim_qty': forms.NumberInput(attrs={'class': 'erp-input'}),
        }


class StockDemandForm(forms.ModelForm):
    class Meta:
        model = StockDemand
        fields = ['item', 'month_str', 'sheet_qty_pcs', 'pkt_rim_qty']
        widgets = {
            'item': forms.Select(attrs={'class': 'erp-select'}),
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
        fields = ['item', 'count_date', 'physical_sheet_qty', 'physical_pkt_rim_qty', 'notes']
        widgets = {
            'item': forms.Select(attrs={'class': 'erp-select'}),
            'count_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'physical_sheet_qty': forms.NumberInput(attrs={'class': 'erp-input'}),
            'physical_pkt_rim_qty': forms.NumberInput(attrs={'class': 'erp-input'}),
            'notes': forms.TextInput(attrs={'class': 'erp-input'}),
        }


class BulkPhysicalCountForm(forms.Form):
    count_date = forms.DateField(
        widget=forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
    )
