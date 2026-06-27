from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from django.db.models import Sum
from .models import SupplyChainItem, StockTransaction, StockDemand

@login_required
def dashboard(request):
    items = SupplyChainItem.objects.all()
    
    dashboard_data = []
    
    for item in items:
        # Aggregate transactions
        opening = item.transactions.filter(transaction_type='OPENING').aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0
        receiving = item.transactions.filter(transaction_type='RECEIVING').aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0
        issuance = item.transactions.filter(transaction_type='ISSUANCE').aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0
        adjustment = item.transactions.filter(transaction_type='ADJUSTMENT').aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0
        
        monthly_demand = item.demands.aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0
        
        closing = (opening + receiving) - issuance + adjustment
        
        # Alerts logic
        stockout = closing <= 0
        reorder = closing <= (item.safety_stock + 100) # Simple fallback for ROP
        overstock = closing > item.max_stock_level
        
        dashboard_data.append({
            'item': item,
            'opening': opening,
            'receiving': receiving,
            'issuance': issuance,
            'adjustment': adjustment,
            'closing': closing,
            'monthly_demand': monthly_demand,
            'stockout': stockout,
            'reorder': reorder,
            'overstock': overstock,
        })
        
    context = {
        'dashboard_data': dashboard_data
    }
    
    return render(request, 'supply_chain/dashboard.html', context)
