from django.contrib import messages
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse

from core.models import JobCard

from .decorators import supply_chain_required
from .excel_io import (
    export_demands,
    export_item_wise_consumption,
    export_items,
    export_kpi_dashboard,
    export_month_wise_consumption,
    export_physical_counts,
    export_transactions,
    import_demands,
    import_transactions,
)
from .forms import (
    BulkPhysicalCountForm,
    ExcelUploadForm,
    PhysicalStockCountForm,
    StockDemandForm,
    StockTransactionForm,
    SupplyChainItemForm,
)
from .models import PhysicalStockCount, SupplyChainItem
from .jc_sync import (
    build_job_card_link_rows,
    sync_all_job_card_issuances,
    sync_issuance_for_job_card,
)
from .kpis import build_kpi_dashboard_data
from .physical_count import (
    build_physical_count_rows,
    physical_count_history,
    save_physical_count,
)
from .reports import (
    build_item_wise_monthly_consumption,
    build_month_wise_item_consumption,
    parse_report_filters,
)
from .services import build_dashboard_data, demand_queryset, transaction_queryset


TRANSACTION_PAGES = {
    'opening': {
        'transaction_type': 'OPENING',
        'title': 'Stock Opening',
        'subtitle': 'Record opening balances at the start of each inventory period.',
        'export_name': 'stock_opening.xlsx',
    },
    'receiving': {
        'transaction_type': 'RECEIVING',
        'title': 'Stock Receiving',
        'subtitle': 'Log inward stock against GIN or job card references.',
        'export_name': 'stock_receiving.xlsx',
    },
    'issuance': {
        'transaction_type': 'ISSUANCE',
        'title': 'Stock Issuance',
        'subtitle': 'Track material issued to production via GIN or job card.',
        'export_name': 'stock_issuance.xlsx',
    },
    'adjustment': {
        'transaction_type': 'ADJUSTMENT',
        'title': 'Stock Adjustment',
        'subtitle': 'Correct stock levels after physical counts or write-offs.',
        'export_name': 'stock_adjustment.xlsx',
    },
}


@supply_chain_required
def dashboard(request):
    dashboard_data, kpi_summary = build_kpi_dashboard_data()
    return render(request, 'supply_chain/dashboard.html', {
        'dashboard_data': dashboard_data,
        'alert_counts': kpi_summary,
        'kpi_summary': kpi_summary,
    })


@supply_chain_required
def kpi_dashboard(request):
    kpi_rows, kpi_summary = build_kpi_dashboard_data()
    if request.GET.get('export') == 'xlsx':
        return export_kpi_dashboard(kpi_rows)
    return render(request, 'supply_chain/kpi_dashboard.html', {
        'kpi_rows': kpi_rows,
        'kpi_summary': kpi_summary,
    })


@supply_chain_required
def item_list(request):
    items = SupplyChainItem.objects.select_related('material').order_by('item_id', 'material__name')
    if request.method == 'POST' and request.POST.get('action') == 'import':
        return _handle_item_import(request)
    if request.GET.get('export') == 'xlsx':
        return export_items(items, 'supply_chain_items.xlsx')

    upload_form = ExcelUploadForm()
    return render(request, 'supply_chain/items.html', {
        'items': items,
        'upload_form': upload_form,
    })


@supply_chain_required
def item_edit(request, pk):
    item = get_object_or_404(SupplyChainItem.objects.select_related('material'), pk=pk)
    if request.method == 'POST':
        form = SupplyChainItemForm(request.POST, instance=item)
        if form.is_valid():
            form.save()
            messages.success(request, f'Updated {item.item_id or item.material.name}.')
            return redirect('supply_chain:items')
    else:
        form = SupplyChainItemForm(instance=item)

    return render(request, 'supply_chain/item_edit.html', {'item': item, 'form': form})


@supply_chain_required
def monthly_demand(request):
    month_filter = (request.GET.get('month') or '').strip()
    demands = demand_queryset(month_filter or None)

    if request.method == 'POST':
        if request.POST.get('action') == 'import':
            return _handle_demand_import(request, month_filter)
        form = StockDemandForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'Monthly demand entry saved.')
            return redirect('supply_chain:monthly_demand')
    else:
        form = StockDemandForm()

    if request.GET.get('export') == 'xlsx':
        return export_demands(demands, 'stock_monthly_demand.xlsx')

    upload_form = ExcelUploadForm()
    return render(request, 'supply_chain/monthly_demand.html', {
        'demands': demands,
        'form': form,
        'upload_form': upload_form,
        'month_filter': month_filter,
    })


@supply_chain_required
def transaction_page(request, page_key):
    config = TRANSACTION_PAGES.get(page_key)
    if not config:
        messages.error(request, 'Unknown stock screen.')
        return redirect('supply_chain:dashboard')

    transaction_type = config['transaction_type']
    month_filter = (request.GET.get('month') or '').strip()
    transactions = transaction_queryset(transaction_type, month_filter or None)

    if request.method == 'POST':
        if request.POST.get('action') == 'import':
            return _handle_transaction_import(request, page_key, transaction_type, month_filter)
        if request.POST.get('action') == 'sync_jc' and page_key == 'issuance':
            synced, skipped = sync_all_job_card_issuances()
            messages.success(request, f'Synced {synced} issuance row(s) from job cards. Skipped {skipped}.')
            return redirect('supply_chain:issuance')
        form = StockTransactionForm(request.POST)
        if form.is_valid():
            txn = form.save(commit=False)
            txn.transaction_type = transaction_type
            txn.save()
            messages.success(request, f'{config["title"]} entry saved.')
            return redirect('supply_chain:' + page_key)
    else:
        form = StockTransactionForm()

    if request.GET.get('export') == 'xlsx':
        return export_transactions(transaction_type, transactions, config['export_name'])

    upload_form = ExcelUploadForm()
    return render(request, 'supply_chain/transactions.html', {
        'page_key': page_key,
        'page_url': reverse('supply_chain:' + page_key),
        'config': config,
        'transactions': transactions,
        'form': form,
        'upload_form': upload_form,
        'month_filter': month_filter,
        'show_jc_sync': page_key == 'issuance',
    })


def _handle_transaction_import(request, page_key, transaction_type, month_filter):
    upload_form = ExcelUploadForm(request.POST, request.FILES)
    if not upload_form.is_valid():
        messages.error(request, 'Please choose a valid Excel or CSV file.')
        return redirect('supply_chain:' + page_key)

    try:
        created, skipped = import_transactions(transaction_type, upload_form.cleaned_data['upload_file'])
        messages.success(request, f'Imported {created} row(s). Skipped {skipped} row(s) without a matching Item ID.')
    except Exception as exc:
        messages.error(request, f'Import failed: {exc}')

    if month_filter:
        return redirect(f"{request.path}?month={month_filter}")
    return redirect('supply_chain:' + page_key)


def _handle_demand_import(request, month_filter):
    upload_form = ExcelUploadForm(request.POST, request.FILES)
    if not upload_form.is_valid():
        messages.error(request, 'Please choose a valid Excel or CSV file.')
        return redirect('supply_chain:monthly_demand')

    try:
        created, skipped = import_demands(upload_form.cleaned_data['upload_file'])
        messages.success(request, f'Imported {created} demand row(s). Skipped {skipped} row(s) without a matching Item ID.')
    except Exception as exc:
        messages.error(request, f'Import failed: {exc}')

    if month_filter:
        return redirect(f"{request.path}?month={month_filter}")
    return redirect('supply_chain:monthly_demand')


def _handle_item_import(request):
    messages.info(request, 'Item master import is available via export/edit workflow. Bulk item creation uses Material master sync.')
    return redirect('supply_chain:items')


@supply_chain_required
def consumption_reports(request):
    filters = parse_report_filters(request)
    item_wise_rows = build_item_wise_monthly_consumption(
        month_filter=filters['month_filter'],
        from_date=filters['from_date_parsed'],
        to_date=filters['to_date_parsed'],
    )
    month_wise_rows = build_month_wise_item_consumption(
        month_filter=filters['month_filter'],
        from_date=filters['from_date_parsed'],
        to_date=filters['to_date_parsed'],
    )

    export_type = (request.GET.get('export') or '').strip()
    if export_type == 'item_wise':
        return export_item_wise_consumption(item_wise_rows)
    if export_type == 'month_wise':
        return export_month_wise_consumption(month_wise_rows)

    totals = {
        'sheet_qty_pcs': sum(row['sheet_qty_pcs'] for row in item_wise_rows),
        'pkt_rim_qty': sum(row['pkt_rim_qty'] for row in item_wise_rows),
        'consumption_value': round(sum(row['consumption_value'] for row in item_wise_rows), 2),
    }

    return render(request, 'supply_chain/consumption_reports.html', {
        'filters': filters,
        'item_wise_rows': item_wise_rows,
        'month_wise_rows': month_wise_rows,
        'totals': totals,
    })


@supply_chain_required
def jc_links(request):
    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'sync_all':
            synced, skipped = sync_all_job_card_issuances()
            messages.success(request, f'Synced {synced} issuance row(s) from job cards. Skipped {skipped}.')
            return redirect('supply_chain:jc_links')
        if action == 'sync_one':
            job_card_id = request.POST.get('job_card_id')
            job_card = get_object_or_404(JobCard, pk=job_card_id)
            synced, skipped = sync_issuance_for_job_card(job_card)
            messages.success(
                request,
                f'Synced {synced} issuance row(s) for {job_card.job_card_no}. Skipped {skipped}.',
            )
            return redirect('supply_chain:jc_links')

    link_rows = build_job_card_link_rows()
    summary = {
        'total': len(link_rows),
        'linked': sum(1 for row in link_rows if row['is_linked']),
        'unlinked': sum(1 for row in link_rows if not row['is_linked']),
    }
    return render(request, 'supply_chain/jc_links.html', {
        'link_rows': link_rows,
        'summary': summary,
    })


@supply_chain_required
def physical_counts(request):
    count_rows = build_physical_count_rows()
    history = physical_count_history()

    if request.GET.get('export') == 'xlsx':
        return export_physical_counts(history)

    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'single':
            form = PhysicalStockCountForm(request.POST)
            if form.is_valid():
                item = form.cleaned_data['item']
                save_physical_count(
                    item=item,
                    count_date=form.cleaned_data['count_date'],
                    physical_sheet_qty=form.cleaned_data['physical_sheet_qty'],
                    physical_pkt_rim_qty=form.cleaned_data['physical_pkt_rim_qty'],
                    notes=form.cleaned_data['notes'],
                )
                messages.success(request, f'Physical count saved for {item.item_id or item.material.name}.')
                return redirect('supply_chain:physical_counts')
        elif action == 'bulk':
            bulk_form = BulkPhysicalCountForm(request.POST)
            if bulk_form.is_valid():
                count_date = bulk_form.cleaned_data['count_date']
                saved = 0
                for row in count_rows:
                    field_name = f'physical_{row["item"].pk}'
                    if field_name not in request.POST:
                        continue
                    raw_value = (request.POST.get(field_name) or '').strip()
                    if raw_value == '':
                        continue
                    save_physical_count(
                        item=row['item'],
                        count_date=count_date,
                        physical_sheet_qty=int(raw_value),
                    )
                    saved += 1
                messages.success(request, f'Saved {saved} physical count record(s).')
                return redirect('supply_chain:physical_counts')

    form = PhysicalStockCountForm()
    bulk_form = BulkPhysicalCountForm()
    accuracy_values = [
        float(row['latest_accuracy'])
        for row in count_rows
        if row['latest_accuracy'] is not None
    ]
    summary = {
        'avg_accuracy': round(sum(accuracy_values) / len(accuracy_values), 2) if accuracy_values else None,
        'items_counted': len(accuracy_values),
        'history_count': PhysicalStockCount.objects.count(),
    }
    return render(request, 'supply_chain/physical_counts.html', {
        'form': form,
        'bulk_form': bulk_form,
        'count_rows': count_rows,
        'history': history,
        'summary': summary,
    })
