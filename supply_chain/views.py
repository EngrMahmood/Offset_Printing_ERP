from django.contrib import messages
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from core.models import JobCard

from .decorators import supply_chain_required
from .demand_gap import (
    available_plan_months,
    build_demand_gap_report,
    parse_gap_filters,
)
from .excel_io import (
    export_demands,
    export_demand_gap_jobs,
    export_demand_gap_materials,
    export_item_wise_consumption,
    export_items,
    export_kpi_dashboard,
    export_month_wise_consumption,
    export_physical_counts,
    export_raw_material_sku_template,
    export_transactions,
    import_demands,
    import_raw_material_skus,
    import_transactions,
)
from .forms import (
    BulkPhysicalCountForm,
    ExcelUploadForm,
    PhysicalStockCountForm,
    QuickRawMaterialSkuForm,
    RawMaterialSkuForm,
    StockDemandForm,
    StockTransactionForm,
)
from .jc_sync import (
    build_job_card_link_rows,
    sync_all_job_card_issuances,
    sync_issuance_for_job_card,
)
from .kpis import build_kpi_dashboard_data
from .models import PhysicalStockCount, RawMaterialSku
from .physical_count import (
    build_physical_count_rows,
    physical_count_history,
    save_physical_count,
)
from .raw_material_sku import list_material_choices, upsert_raw_material_sku_row
from .reports import (
    build_item_wise_monthly_consumption,
    build_month_wise_item_consumption,
    parse_report_filters,
)
from .services import demand_queryset, transaction_queryset


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
    items = RawMaterialSku.objects.select_related('material').order_by('material__name', 'purchase_sheet_size', 'sku')

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'import':
            return _handle_raw_material_import(request)
        if action == 'quick_add':
            return _handle_quick_add_raw_material(request)

    if request.GET.get('export') == 'xlsx':
        return export_items(items, 'raw_material_skus.xlsx')
    if request.GET.get('export') == 'template':
        return export_raw_material_sku_template()

    return render(request, 'supply_chain/raw_material_skus.html', {
        'items': items,
        'upload_form': ExcelUploadForm(),
        'quick_form': QuickRawMaterialSkuForm(),
        'material_choices': list_material_choices(),
    })


@supply_chain_required
@require_POST
def quick_add_raw_material_sku(request):
    return _handle_quick_add_raw_material(request, as_json=True)


def _handle_quick_add_raw_material(request, as_json=False):
    form = QuickRawMaterialSkuForm(request.POST)
    if not form.is_valid():
        if as_json:
            return JsonResponse({'ok': False, 'error': form.errors.as_text()}, status=400)
        messages.error(request, form.errors.as_text())
        return redirect('supply_chain:items')

    obj, errors, _created = upsert_raw_material_sku_row({
        'sku': form.cleaned_data['sku'],
        'material_name': form.cleaned_data['material_name'],
        'purchase_sheet_size': form.cleaned_data['purchase_sheet_size'],
    })
    if errors:
        if as_json:
            return JsonResponse({'ok': False, 'error': '; '.join(errors)}, status=400)
        messages.error(request, '; '.join(errors))
        return redirect('supply_chain:items')

    if as_json:
        return JsonResponse({
            'ok': True,
            'id': obj.pk,
            'sku': obj.sku,
            'material_name': obj.material.name,
            'purchase_sheet_size': obj.purchase_sheet_size,
            'display_label': obj.display_label,
        })

    messages.success(request, f'Saved raw material SKU {obj.sku}.')
    return redirect('supply_chain:items')


def _handle_raw_material_import(request):
    upload_form = ExcelUploadForm(request.POST, request.FILES)
    if not upload_form.is_valid():
        messages.error(request, 'Please choose a valid Excel or CSV file.')
        return redirect('supply_chain:items')

    try:
        result = import_raw_material_skus(upload_form.cleaned_data['upload_file'])
        messages.success(
            request,
            f"Imported {result['created']} new and updated {result['updated']} raw material SKU(s). "
            f"Skipped {result['skipped']}.",
        )
        if result['errors']:
            messages.warning(request, ' '.join(result['errors'][:5]))
    except Exception as exc:
        messages.error(request, f'Import failed: {exc}')
    return redirect('supply_chain:items')


@supply_chain_required
def item_edit(request, pk):
    item = get_object_or_404(RawMaterialSku.objects.select_related('material'), pk=pk)
    if request.method == 'POST':
        form = RawMaterialSkuForm(request.POST, instance=item)
        if form.is_valid():
            form.save()
            messages.success(request, f'Updated {item.sku}.')
            return redirect('supply_chain:items')
    else:
        form = RawMaterialSkuForm(instance=item)

    return render(request, 'supply_chain/raw_material_sku_edit.html', {'item': item, 'form': form})


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

    return render(request, 'supply_chain/monthly_demand.html', {
        'demands': demands,
        'form': form,
        'upload_form': ExcelUploadForm(),
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

    return render(request, 'supply_chain/transactions.html', {
        'page_key': page_key,
        'page_url': reverse('supply_chain:' + page_key),
        'config': config,
        'transactions': transactions,
        'form': form,
        'upload_form': ExcelUploadForm(),
        'month_filter': month_filter,
        'show_jc_sync': page_key == 'issuance',
    })


def _handle_transaction_import(request, page_key, transaction_type, month_filter):
    upload_form = ExcelUploadForm(request.POST, request.FILES)
    if not upload_form.is_valid():
        messages.error(request, 'Please choose a valid Excel or CSV file.')
        return redirect('supply_chain:' + page_key)

    try:
        result = import_transactions(transaction_type, upload_form.cleaned_data['upload_file'])
        messages.success(
            request,
            f"Imported {result['created']} row(s). "
            f"Skipped {result['skipped_duplicates']} duplicate(s) and "
            f"{result['skipped_missing_sku']} without a matching raw material SKU.",
        )
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
        result = import_demands(upload_form.cleaned_data['upload_file'])
        messages.success(
            request,
            f"Imported {result['created']} demand row(s). "
            f"Skipped {result['skipped_duplicates']} duplicate(s) and "
            f"{result['skipped_missing_sku']} without a matching raw material SKU.",
        )
    except Exception as exc:
        messages.error(request, f'Import failed: {exc}')

    if month_filter:
        return redirect(f"{request.path}?month={month_filter}")
    return redirect('supply_chain:monthly_demand')


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
                item = form.cleaned_data['raw_material_sku']
                save_physical_count(
                    item=item,
                    count_date=form.cleaned_data['count_date'],
                    physical_sheet_qty=form.cleaned_data['physical_sheet_qty'],
                    physical_pkt_rim_qty=form.cleaned_data['physical_pkt_rim_qty'],
                    notes=form.cleaned_data['notes'],
                )
                messages.success(request, f'Physical count saved for {item.sku}.')
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
        'form': PhysicalStockCountForm(),
        'bulk_form': BulkPhysicalCountForm(),
        'count_rows': count_rows,
        'history': history,
        'summary': summary,
    })


@supply_chain_required
def demand_gap(request):
    filters = parse_gap_filters(request.GET)
    report = build_demand_gap_report(filters)

    export_type = request.GET.get('export')
    if export_type == 'materials':
        return export_demand_gap_materials(report['material_rows'])
    if export_type == 'jobs':
        return export_demand_gap_jobs(report['job_rows'])

    from planning.models import PLANNING_STATUS_CHOICES

    return render(request, 'supply_chain/demand_gap.html', {
        'material_rows': report['material_rows'],
        'job_rows': report['job_rows'],
        'summary': report['summary'],
        'filters': filters,
        'plan_months': available_plan_months(),
        'status_choices': PLANNING_STATUS_CHOICES,
        'process_type_choices': [
            ('print_and_pack', 'Print + Pack'),
            ('cut_and_pack', 'Cut & Pack'),
        ],
    })
