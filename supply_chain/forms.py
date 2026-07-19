import io

from django import forms
from django.core.files.base import ContentFile
from django.core.files.uploadedfile import UploadedFile

from core.models import Machine, Material

from .models import (
    ItemProcurementTimeline,
    ItemRequest,
    ItemRequestDepartment,
    ItemRequestQuote,
    ItemRequestType,
    PhysicalStockCount,
    RawMaterialSku,
    StockDemand,
    StockTransaction,
)


def compress_image_upload(uploaded_file):
    """Compress image uploads (JPEG/PNG) to keep the media folder lean; other file types pass through untouched."""
    if not uploaded_file or not isinstance(uploaded_file, UploadedFile):
        return uploaded_file
    if not uploaded_file.content_type or not uploaded_file.content_type.startswith('image/'):
        return uploaded_file

    try:
        from PIL import Image
    except ImportError:
        return uploaded_file

    try:
        image = Image.open(uploaded_file)
        image = image.convert('RGB') if image.mode in ('RGBA', 'P') else image
        max_dim = 1600
        if image.width > max_dim or image.height > max_dim:
            image.thumbnail((max_dim, max_dim), Image.LANCZOS)

        buffer = io.BytesIO()
        image.save(buffer, format='JPEG', quality=70, optimize=True)
        buffer.seek(0)

        name = uploaded_file.name.rsplit('.', 1)[0] + '.jpg'
        return ContentFile(buffer.read(), name=name)
    except Exception:
        return uploaded_file


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


class ItemRequestForm(forms.ModelForm):
    class Meta:
        model = ItemRequest
        fields = [
            'request_type',
            'request_date',
            'item_title',
            'machine',
            'machine_other',
            'uom',
            'specifications',
            'description',
            'dimensions',
            'local_import',
            'part_number',
            'required_quantity',
            'department',
            'existing_sku',
            'estimated_unit_price',
            'cost_centre',
            'attachment',
        ]
        widgets = {
            'request_type': forms.Select(attrs={'class': 'erp-select'}),
            'request_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'item_title': forms.TextInput(attrs={'class': 'erp-input'}),
            'machine': forms.Select(attrs={'class': 'erp-select'}),
            'machine_other': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': "Fill only if the machine isn't listed"}),
            'uom': forms.TextInput(attrs={'class': 'erp-input'}),
            'specifications': forms.Textarea(attrs={'class': 'erp-input', 'rows': 3}),
            'description': forms.Textarea(attrs={'class': 'erp-input', 'rows': 3}),
            'dimensions': forms.TextInput(attrs={'class': 'erp-input'}),
            'local_import': forms.Select(attrs={'class': 'erp-select'}),
            'part_number': forms.TextInput(attrs={'class': 'erp-input'}),
            'required_quantity': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.01'}),
            'department': forms.Select(attrs={'class': 'erp-select'}),
            'existing_sku': forms.Select(attrs={'class': 'erp-select'}),
            'estimated_unit_price': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.01'}),
            'cost_centre': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'e.g. CC-1042'}),
            'attachment': forms.ClearableFileInput(attrs={'class': 'erp-input'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['request_type'].queryset = ItemRequestType.objects.filter(is_active=True)
        self.fields['machine'].queryset = Machine.objects.filter(is_active=True) if hasattr(Machine, 'is_active') else Machine.objects.all()
        self.fields['machine'].required = False
        self.fields['machine_other'].required = False
        self.fields['department'].queryset = ItemRequestDepartment.objects.filter(is_active=True)
        self.fields['existing_sku'].queryset = RawMaterialSku.objects.filter(is_active=True).select_related('material')
        self.fields['existing_sku'].required = False
        self.fields['existing_sku'].widget.attrs['data-sku-search'] = '1'
        self.fields['estimated_unit_price'].required = False
        self.fields['attachment'].required = False
        self.fields['description'].required = False
        self.fields['dimensions'].required = False
        self.fields['local_import'].required = False
        self.fields['part_number'].required = False
        self.fields['cost_centre'].required = False

    def clean_attachment(self):
        return compress_image_upload(self.cleaned_data.get('attachment'))


class ItemRequestReviewForm(forms.Form):
    ACTION_CHOICES = [
        ('APPROVE', 'Approve'),
        ('REJECT', 'Reject'),
        ('REVISE', 'Needs Revision'),
    ]
    action = forms.ChoiceField(choices=ACTION_CHOICES, widget=forms.RadioSelect)
    comment = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'erp-input', 'rows': 2}))

    def clean(self):
        cleaned = super().clean()
        if cleaned.get('action') in ('REJECT', 'REVISE') and not cleaned.get('comment'):
            raise forms.ValidationError('A comment is required for this action.')
        return cleaned


class ItemRequestTypeQuickForm(forms.ModelForm):
    class Meta:
        model = ItemRequestType
        fields = ['name', 'code']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'erp-input'}),
            'code': forms.TextInput(attrs={'class': 'erp-input', 'maxlength': 4}),
        }


class DepartmentQuickForm(forms.ModelForm):
    class Meta:
        model = ItemRequestDepartment
        fields = ['name']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'erp-input'}),
        }


class ItemProcurementTimelineForm(forms.ModelForm):
    class Meta:
        model = ItemProcurementTimeline
        fields = [
            'item_code',
            'code_opened_date',
            'indent_pr_no',
            'pr_date',
            'po_no',
            'po_date',
            'supplier',
            'received_date',
            'unit_price',
            'received_qty',
            'remarks',
        ]
        widgets = {
            'item_code': forms.TextInput(attrs={'class': 'erp-input'}),
            'code_opened_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'indent_pr_no': forms.TextInput(attrs={'class': 'erp-input'}),
            'pr_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'po_no': forms.TextInput(attrs={'class': 'erp-input'}),
            'po_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'supplier': forms.TextInput(attrs={'class': 'erp-input'}),
            'received_date': forms.DateInput(attrs={'class': 'erp-input', 'type': 'date'}),
            'unit_price': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.01'}),
            'received_qty': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.01'}),
            'remarks': forms.Textarea(attrs={'class': 'erp-input', 'rows': 2}),
        }


class ItemRequestQuoteForm(forms.ModelForm):
    class Meta:
        model = ItemRequestQuote
        fields = ['supplier', 'quoted_price', 'file', 'notes']
        widgets = {
            'supplier': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'Supplier name'}),
            'quoted_price': forms.NumberInput(attrs={'class': 'erp-input', 'step': '0.01'}),
            'file': forms.ClearableFileInput(attrs={'class': 'erp-input'}),
            'notes': forms.TextInput(attrs={'class': 'erp-input'}),
        }

    def clean_file(self):
        return compress_image_upload(self.cleaned_data.get('file'))
