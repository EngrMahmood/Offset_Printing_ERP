from django.db import models
from django.utils import timezone

class SupplyChainItem(models.Model):
    material = models.OneToOneField('core.Material', on_delete=models.CASCADE, related_name='supply_chain_details')
    item_id = models.CharField(max_length=50, unique=True, verbose_name="Item ID", blank=True, null=True)
    uom = models.CharField(max_length=20, verbose_name="UOM", default="Sheets")
    sheet_packing_pcs = models.IntegerField(verbose_name="Sheet Packing/Pcs", default=1)
    
    # Required for KPIs
    unit_cost = models.DecimalField(max_digits=12, decimal_places=2, default=0.0, verbose_name="Unit Cost")
    safety_stock = models.IntegerField(default=0, verbose_name="Safety Stock")
    max_stock_level = models.IntegerField(default=10000, verbose_name="Maximum Stock Level")
    lead_time_days = models.IntegerField(default=1, verbose_name="Lead Time (Days)")

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.item_id or 'N/A'} - {self.material.name}"

class StockTransaction(models.Model):
    TRANSACTION_TYPES = [
        ('OPENING', 'Opening'),
        ('RECEIVING', 'Receiving'),
        ('ISSUANCE', 'Issuance'),
        ('ADJUSTMENT', 'Adjustment'),
    ]

    item = models.ForeignKey(SupplyChainItem, on_delete=models.CASCADE, related_name="transactions")
    transaction_type = models.CharField(max_length=20, choices=TRANSACTION_TYPES)
    
    month_str = models.CharField(max_length=20, blank=True, null=True, verbose_name="Month")
    date = models.DateField(default=timezone.now, verbose_name="Date")
    gin_jc = models.CharField(max_length=50, blank=True, null=True, verbose_name="GIN / JC")
    
    sheet_qty_pcs = models.IntegerField(default=0, verbose_name="Sheet Qty/Pcs")
    pkt_rim_qty = models.IntegerField(default=0, verbose_name="Pkt/Rim Qty")

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.get_transaction_type_display()} - {self.item.item_id} - {self.date}"

class StockDemand(models.Model):
    item = models.ForeignKey(SupplyChainItem, on_delete=models.CASCADE, related_name="demands")
    month_str = models.CharField(max_length=20, blank=True, null=True, verbose_name="Month")
    
    sheet_qty_pcs = models.IntegerField(default=0, verbose_name="Sheet Qty/Pcs")
    pkt_rim_qty = models.IntegerField(default=0, verbose_name="Pkt/Rim Qty")

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"Demand - {self.item.item_id} - {self.month_str}"
