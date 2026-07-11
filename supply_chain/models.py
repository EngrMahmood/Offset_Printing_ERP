import re

from django.db import models
from django.utils import timezone


# Width × height: "10.5 x 15", "10.5X15", "10.5*15", "10.5 × 15" → "10.5*15"
_PURCHASE_SIZE_PAIR_RE = re.compile(
    r'^\s*(\d+(?:\.\d+)?)\s*[xX×*]\s*(\d+(?:\.\d+)?)\s*$'
)


def normalize_purchase_sheet_size(value):
    """Canonical purchase sheet size for matching and rollup.

    Dimension pairs become ``width*height`` (no spaces). Other labels keep
    collapsed whitespace only (e.g. ``A4``).
    """
    text = ' '.join(str(value or '').strip().split())
    if not text:
        return ''
    match = _PURCHASE_SIZE_PAIR_RE.match(text)
    if match:
        return f'{match.group(1)}*{match.group(2)}'
    return text


def normalize_material_name(value):
    """Case-insensitive material key: collapse whitespace and lowercase."""
    return ' '.join(str(value or '').strip().split()).lower()


def display_material_name(value):
    """Stable display label for free-text material names."""
    text = ' '.join(str(value or '').strip().split())
    if not text:
        return ''
    if text.isupper() or text.islower():
        return text.title()
    return text


class RawMaterialSku(models.Model):
    """Procurement + inventory identity: material type and purchase sheet size."""

    sku = models.CharField(max_length=50, unique=True, verbose_name='Raw Material SKU')
    material = models.ForeignKey(
        'core.Material',
        on_delete=models.PROTECT,
        related_name='raw_material_skus',
    )
    purchase_sheet_size = models.CharField(max_length=80, verbose_name='Purchase Sheet Size')
    uom = models.CharField(max_length=20, verbose_name='UOM', default='Sheets')
    sheet_packing_pcs = models.IntegerField(verbose_name='Sheet Packing/Pcs', default=1)
    unit_cost = models.DecimalField(max_digits=12, decimal_places=2, default=0.0, verbose_name='Unit Cost')
    safety_stock = models.IntegerField(default=0, verbose_name='Safety Stock')
    max_stock_level = models.IntegerField(default=10000, verbose_name='Maximum Stock Level')
    lead_time_days = models.IntegerField(default=1, verbose_name='Lead Time (Days)')
    is_active = models.BooleanField(default=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['material__name', 'purchase_sheet_size', 'sku']
        unique_together = ('material', 'purchase_sheet_size')
        verbose_name = 'Raw Material SKU'
        verbose_name_plural = 'Raw Material SKUs'

    def __str__(self):
        return self.display_label

    @property
    def display_label(self):
        return f'{self.material.name} — {self.purchase_sheet_size} ({self.sku})'

    @property
    def item_id(self):
        return self.sku


class SupplyChainItem(models.Model):
    """Legacy inventory row — superseded by RawMaterialSku."""

    material = models.OneToOneField(
        'core.Material',
        on_delete=models.CASCADE,
        related_name='supply_chain_details',
    )
    item_id = models.CharField(max_length=50, unique=True, verbose_name='Item ID', blank=True, null=True)
    uom = models.CharField(max_length=20, verbose_name='UOM', default='Sheets')
    sheet_packing_pcs = models.IntegerField(verbose_name='Sheet Packing/Pcs', default=1)
    unit_cost = models.DecimalField(max_digits=12, decimal_places=2, default=0.0, verbose_name='Unit Cost')
    safety_stock = models.IntegerField(default=0, verbose_name='Safety Stock')
    max_stock_level = models.IntegerField(default=10000, verbose_name='Maximum Stock Level')
    lead_time_days = models.IntegerField(default=1, verbose_name='Lead Time (Days)')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f'{self.item_id or "N/A"} - {self.material.name}'


class StockTransaction(models.Model):
    TRANSACTION_TYPES = [
        ('OPENING', 'Opening'),
        ('RECEIVING', 'Receiving'),
        ('ISSUANCE', 'Issuance'),
        ('ADJUSTMENT', 'Adjustment'),
    ]
    SOURCE_CHOICES = [
        ('MANUAL', 'Manual'),
        ('JOB_CARD', 'Job Card'),
    ]

    raw_material_sku = models.ForeignKey(
        RawMaterialSku,
        on_delete=models.CASCADE,
        related_name='transactions',
        null=True,
        blank=True,
    )
    item = models.ForeignKey(
        SupplyChainItem,
        on_delete=models.CASCADE,
        related_name='legacy_transactions',
        null=True,
        blank=True,
    )
    transaction_type = models.CharField(max_length=20, choices=TRANSACTION_TYPES)
    source = models.CharField(max_length=20, choices=SOURCE_CHOICES, default='MANUAL')
    job_card = models.ForeignKey(
        'core.JobCard',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='stock_transactions',
    )
    production = models.OneToOneField(
        'core.Production',
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='stock_issuance',
    )
    month_str = models.CharField(max_length=20, blank=True, null=True, verbose_name='Month')
    date = models.DateField(default=timezone.now, verbose_name='Date')
    gin_jc = models.CharField(max_length=50, blank=True, null=True, verbose_name='GIN / JC')
    sheet_qty_pcs = models.IntegerField(default=0, verbose_name='Sheet Qty/Pcs')
    pkt_rim_qty = models.IntegerField(default=0, verbose_name='Pkt/Rim Qty')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        sku = self.raw_material_sku.sku if self.raw_material_sku_id else 'N/A'
        return f'{self.get_transaction_type_display()} - {sku} - {self.date}'


class StockDemand(models.Model):
    raw_material_sku = models.ForeignKey(
        RawMaterialSku,
        on_delete=models.CASCADE,
        related_name='demands',
        null=True,
        blank=True,
    )
    item = models.ForeignKey(
        SupplyChainItem,
        on_delete=models.CASCADE,
        related_name='legacy_demands',
        null=True,
        blank=True,
    )
    month_str = models.CharField(max_length=20, blank=True, null=True, verbose_name='Month')
    sheet_qty_pcs = models.IntegerField(default=0, verbose_name='Sheet Qty/Pcs')
    pkt_rim_qty = models.IntegerField(default=0, verbose_name='Pkt/Rim Qty')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        sku = self.raw_material_sku.sku if self.raw_material_sku_id else 'N/A'
        return f'Demand - {sku} - {self.month_str}'


class PhysicalStockCount(models.Model):
    raw_material_sku = models.ForeignKey(
        RawMaterialSku,
        on_delete=models.CASCADE,
        related_name='physical_counts',
        null=True,
        blank=True,
    )
    item = models.ForeignKey(
        SupplyChainItem,
        on_delete=models.CASCADE,
        related_name='legacy_physical_counts',
        null=True,
        blank=True,
    )
    count_date = models.DateField(default=timezone.now, verbose_name='Count Date')
    physical_sheet_qty = models.IntegerField(default=0, verbose_name='Physical Sheet Qty/Pcs')
    physical_pkt_rim_qty = models.IntegerField(default=0, verbose_name='Physical Pkt/Rim Qty')
    system_sheet_qty = models.IntegerField(default=0, verbose_name='System Sheet Qty/Pcs')
    system_pkt_rim_qty = models.IntegerField(default=0, verbose_name='System Pkt/Rim Qty')
    accuracy_percent = models.DecimalField(
        max_digits=6,
        decimal_places=2,
        default=0,
        verbose_name='Inventory Accuracy %',
    )
    notes = models.CharField(max_length=255, blank=True, default='')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-count_date', '-id']

    def __str__(self):
        sku = self.raw_material_sku.sku if self.raw_material_sku_id else 'N/A'
        return f'Count {sku} @ {self.count_date} — {self.accuracy_percent}%'

    @property
    def variance(self):
        return self.physical_sheet_qty - self.system_sheet_qty
