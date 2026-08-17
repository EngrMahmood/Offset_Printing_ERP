from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST
from django.utils import timezone


from core.models import JobCard, Machine

from .decorators import item_request_access_required, supply_chain_required
from .demand_gap import (
    available_plan_months,
    build_demand_gap_report,
    parse_gap_filters,
)
from .excel_io import (
    export_demands,
    export_demand_gap_jobs,
    export_demand_gap_materials,
    export_item_requests,
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
    DepartmentQuickForm,
    ExcelUploadForm,
    ItemProcurementTimelineForm,
    ItemRequestForm,
    ItemRequestQuoteForm,
    ItemRequestReviewForm,
    ItemRequestTypeQuickForm,
    PhysicalStockCountForm,
    QuickRawMaterialSkuForm,
    RawMaterialSkuForm,
    StockDemandForm,
    StockTransactionForm,
)
from .item_request_kpis import build_item_request_kpi_data
from .item_request_service import resubmit_request, review_request, submit_request
from .item_request_sla import sla_status
from .jc_sync import (
    build_job_card_link_rows,
    sync_all_job_card_issuances,
    sync_issuance_for_job_card,
)
from .kpis import build_kpi_dashboard_data
from .models import (
    ChangeRequest,
    ItemRequest,
    ItemRequestApproval,
    ItemRequestDepartment,
    ItemRequestQuote,
    ItemRequestType,
    PhysicalStockCount,
    RawMaterialSku,
    StockDemand,
    StockTransaction,
)
from .physical_count import (
    build_physical_count_rows,
    physical_count_history,
    save_physical_count,
    compute_inventory_accuracy,
)
from .raw_material_sku import list_material_choices, upsert_raw_material_sku_row
from .reports import (
    build_item_wise_monthly_consumption,
    build_month_wise_item_consumption,
    parse_report_filters,
)
from .services import apply_transaction_search, demand_queryset, transaction_queryset


def is_supply_chain_admin(user):
    if not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    from core.permissions import user_has_permission
    return user_has_permission(user, 'action.supply_chain_admin')


def serialize_cleaned_data(cleaned_data):
    from django.db.models import Model
    import decimal
    import datetime

    serialized = {}
    for k, v in cleaned_data.items():
        if isinstance(v, Model):
            serialized[f"{k}_id"] = v.pk
        elif isinstance(v, decimal.Decimal):
            serialized[k] = float(v)
        elif isinstance(v, (datetime.date, datetime.datetime)):
            serialized[k] = v.isoformat()
        else:
            serialized[k] = v
    return serialized


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
def stock_balance(request):
    from .services import build_dashboard_data
    rows = build_dashboard_data()
    for row in rows:
        row['below_safety'] = row['closing'] <= row['item'].safety_stock
    below_safety_count = sum(1 for row in rows if row['below_safety'])
    stockout_count = sum(1 for row in rows if row['stockout'])
    return render(request, 'supply_chain/stock_balance.html', {
        'rows': rows,
        'below_safety_count': below_safety_count,
        'stockout_count': stockout_count,
    })


@supply_chain_required
def item_list(request):
    show_archived = request.GET.get('show') == 'archived'
    items = (
        RawMaterialSku.objects
        .filter(is_active=not show_archived)
        .select_related('material', 'sku_type')
        .order_by('material__name', 'purchase_sheet_size', 'sku')
    )

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'import':
            if not is_supply_chain_admin(request.user):
                messages.error(request, 'Only administrators can import raw material SKUs.')
                return redirect('supply_chain:items')
            return _handle_raw_material_import(request)
        if action == 'quick_add':
            return _handle_quick_add_raw_material(request)

    if request.GET.get('export') == 'xlsx':
        return export_items(items, 'raw_material_skus.xlsx')
    if request.GET.get('export') == 'template':
        return export_raw_material_sku_template()

    # One tab per SKU type, plus an "Unclassified" bucket for rows with no type.
    items = list(items)
    type_tabs = []
    for sku_type in ItemRequestType.objects.filter(is_active=True):
        rows = [i for i in items if i.sku_type_id == sku_type.id]
        if rows:
            type_tabs.append({'key': f'type-{sku_type.id}', 'label': sku_type.name, 'items': rows})
    untyped = [i for i in items if i.sku_type_id is None]
    if untyped:
        type_tabs.append({'key': 'type-none', 'label': 'Unclassified', 'items': untyped})

    return render(request, 'supply_chain/raw_material_skus.html', {
        'items': items,
        'type_tabs': type_tabs,
        'upload_form': ExcelUploadForm(),
        'quick_form': QuickRawMaterialSkuForm(),
        'material_choices': list_material_choices(),
        'is_admin': is_supply_chain_admin(request.user),
        'show_archived': show_archived,
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

    proposed_data = {
        'sku': form.cleaned_data['sku'],
        'material_name': form.cleaned_data['material_name'],
        'purchase_sheet_size': form.cleaned_data['purchase_sheet_size'],
    }

    if is_supply_chain_admin(request.user):
        obj, errors, _created = upsert_raw_material_sku_row(proposed_data)
        if errors:
            if as_json:
                return JsonResponse({'ok': False, 'error': '; '.join(errors)}, status=400)
            messages.error(request, '; '.join(errors))
            return redirect('supply_chain:items')

        # Log approved request for audit
        ChangeRequest.objects.create(
            model_name='RawMaterialSku',
            action='CREATE',
            target_id=obj.pk,
            proposed_data=proposed_data,
            requested_by=request.user,
            status='APPROVED',
            reviewed_by=request.user,
            reviewed_at=timezone.now(),
        )

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
    else:
        # Create a pending change request
        ChangeRequest.objects.create(
            model_name='RawMaterialSku',
            action='CREATE',
            proposed_data=proposed_data,
            requested_by=request.user,
            status='PENDING',
        )
        if as_json:
            return JsonResponse({
                'ok': True,
                'message': 'Creation request submitted for admin review.'
            })
        messages.success(request, 'Creation request submitted for admin review.')
    return redirect('supply_chain:items')


def _handle_raw_material_import(request):
    if not is_supply_chain_admin(request.user):
        messages.error(request, 'Only administrators can import raw material SKUs.')
        return redirect('supply_chain:items')

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
    item = get_object_or_404(RawMaterialSku.objects.filter(is_active=True).select_related('material'), pk=pk)
    if request.method == 'POST':
        form = RawMaterialSkuForm(request.POST, instance=item)
        if form.is_valid():
            if is_supply_chain_admin(request.user):
                form.save()
                # Log approved request for audit
                ChangeRequest.objects.create(
                    model_name='RawMaterialSku',
                    action='UPDATE',
                    target_id=item.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='APPROVED',
                    reviewed_by=request.user,
                    reviewed_at=timezone.now(),
                )
                messages.success(request, f'Updated {item.sku}.')
            else:
                ChangeRequest.objects.create(
                    model_name='RawMaterialSku',
                    action='UPDATE',
                    target_id=item.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='PENDING',
                )
                messages.success(request, 'Update request submitted for admin review.')
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
            if not is_supply_chain_admin(request.user):
                messages.error(request, 'Only administrators can import demands.')
                return redirect('supply_chain:monthly_demand')
            return _handle_demand_import(request, month_filter)
        form = StockDemandForm(request.POST)
        if form.is_valid():
            if is_supply_chain_admin(request.user):
                instance = form.save()
                ChangeRequest.objects.create(
                    model_name='StockDemand',
                    action='CREATE',
                    target_id=instance.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='APPROVED',
                    reviewed_by=request.user,
                    reviewed_at=timezone.now(),
                )
                messages.success(request, 'Monthly demand entry saved.')
            else:
                ChangeRequest.objects.create(
                    model_name='StockDemand',
                    action='CREATE',
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='PENDING',
                )
                messages.success(request, 'Creation request submitted for admin review.')
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
        'is_admin': is_supply_chain_admin(request.user),
    })


@supply_chain_required
def transaction_page(request, page_key):
    config = TRANSACTION_PAGES.get(page_key)
    if not config:
        messages.error(request, 'Unknown stock screen.')
        return redirect('supply_chain:dashboard')

    transaction_type = config['transaction_type']
    month_filter = (request.GET.get('month') or '').strip()
    search = (request.GET.get('search') or '').strip()
    transactions = transaction_queryset(transaction_type, month_filter or None, search or None)

    if request.method == 'POST':
        if request.POST.get('action') == 'import':
            if not is_supply_chain_admin(request.user):
                messages.error(request, 'Only administrators can import transactions.')
                return redirect('supply_chain:' + page_key)
            return _handle_transaction_import(request, page_key, transaction_type, month_filter)
        if request.POST.get('action') == 'sync_jc' and page_key == 'issuance':
            if not is_supply_chain_admin(request.user):
                messages.error(request, 'Only administrators can sync job cards directly.')
                return redirect('supply_chain:issuance')
            synced, skipped = sync_all_job_card_issuances()
            messages.success(request, f'Synced {synced} issuance row(s) from job cards. Skipped {skipped}.')
            return redirect('supply_chain:issuance')
        form = StockTransactionForm(request.POST)
        if form.is_valid():
            if is_supply_chain_admin(request.user):
                txn = form.save(commit=False)
                txn.transaction_type = transaction_type
                txn.save()
                ChangeRequest.objects.create(
                    model_name='StockTransaction',
                    action='CREATE',
                    target_id=txn.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='APPROVED',
                    reviewed_by=request.user,
                    reviewed_at=timezone.now(),
                )
                messages.success(request, f'{config["title"]} entry saved.')
            else:
                proposed_data = serialize_cleaned_data(form.cleaned_data)
                proposed_data['transaction_type'] = transaction_type
                ChangeRequest.objects.create(
                    model_name='StockTransaction',
                    action='CREATE',
                    proposed_data=proposed_data,
                    requested_by=request.user,
                    status='PENDING',
                )
                messages.success(request, 'Creation request submitted for admin review.')
            return redirect('supply_chain:' + page_key)
    else:
        form = StockTransactionForm()

    pending_transactions = []
    if page_key == 'issuance':
        pending_transactions = (
            StockTransaction.objects
            .filter(is_active=True, is_approved=False, transaction_type='ISSUANCE')
            .select_related('raw_material_sku', 'raw_material_sku__material', 'job_card', 'production')
            .order_by('-date', '-id')
        )
        pending_transactions = apply_transaction_search(pending_transactions, search)

    if request.GET.get('export') == 'xlsx':
        return export_transactions(transaction_type, transactions, config['export_name'])

    return render(request, 'supply_chain/transactions.html', {
        'page_key': page_key,
        'page_url': reverse('supply_chain:' + page_key),
        'config': config,
        'transactions': transactions,
        'pending_transactions': pending_transactions,
        'form': form,
        'upload_form': ExcelUploadForm(),
        'month_filter': month_filter,
        'search': search,
        'show_jc_sync': page_key == 'issuance',
        'is_admin': is_supply_chain_admin(request.user),
    })


def _handle_transaction_import(request, page_key, transaction_type, month_filter):
    if not is_supply_chain_admin(request.user):
        messages.error(request, 'Only administrators can import transactions.')
        return redirect('supply_chain:' + page_key)

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
    if not is_supply_chain_admin(request.user):
        messages.error(request, 'Only administrators can import demands.')
        return redirect('supply_chain:monthly_demand')

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
            txn = sync_issuance_for_job_card(job_card)
            if txn:
                messages.success(request, f'Synced issuance row for {job_card.job_card_no}.')
            else:
                messages.info(
                    request,
                    f'No issuance row for {job_card.job_card_no} '
                    f'(needs an active production run, a linked SKU, and planned sheets).',
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
                if is_supply_chain_admin(request.user):
                    pc = save_physical_count(
                        item=item,
                        count_date=form.cleaned_data['count_date'],
                        physical_sheet_qty=form.cleaned_data['physical_sheet_qty'],
                        physical_pkt_rim_qty=form.cleaned_data['physical_pkt_rim_qty'],
                        notes=form.cleaned_data['notes'],
                    )
                    ChangeRequest.objects.create(
                        model_name='PhysicalStockCount',
                        action='CREATE',
                        target_id=pc.pk,
                        proposed_data=serialize_cleaned_data(form.cleaned_data),
                        requested_by=request.user,
                        status='APPROVED',
                        reviewed_by=request.user,
                        reviewed_at=timezone.now(),
                    )
                    messages.success(request, f'Physical count saved for {item.sku}.')
                else:
                    ChangeRequest.objects.create(
                        model_name='PhysicalStockCount',
                        action='CREATE',
                        proposed_data=serialize_cleaned_data(form.cleaned_data),
                        requested_by=request.user,
                        status='PENDING',
                    )
                    messages.success(request, 'Creation request submitted for admin review.')
                return redirect('supply_chain:physical_counts')
        elif action == 'bulk':
            if not is_supply_chain_admin(request.user):
                messages.error(request, 'Only administrators can log bulk physical counts.')
                return redirect('supply_chain:physical_counts')
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
                    pc = save_physical_count(
                        item=row['item'],
                        count_date=count_date,
                        physical_sheet_qty=int(raw_value),
                    )
                    ChangeRequest.objects.create(
                        model_name='PhysicalStockCount',
                        action='CREATE',
                        target_id=pc.pk,
                        proposed_data={
                            'raw_material_sku_id': row['item'].pk,
                            'count_date': count_date.isoformat(),
                            'physical_sheet_qty': int(raw_value),
                        },
                        requested_by=request.user,
                        status='APPROVED',
                        reviewed_by=request.user,
                        reviewed_at=timezone.now(),
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
        'history_count': PhysicalStockCount.objects.filter(is_active=True).count(),
    }
    return render(request, 'supply_chain/physical_counts.html', {
        'form': PhysicalStockCountForm(),
        'bulk_form': BulkPhysicalCountForm(),
        'count_rows': count_rows,
        'history': history,
        'summary': summary,
        'is_admin': is_supply_chain_admin(request.user),
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


# Change Request Views
@supply_chain_required
def change_requests_list(request):
    status_filter = request.GET.get('status', 'PENDING').upper()
    if status_filter not in ('PENDING', 'APPROVED', 'REJECTED'):
        status_filter = 'PENDING'

    if is_supply_chain_admin(request.user):
        requests = ChangeRequest.objects.filter(status=status_filter).select_related('requested_by', 'reviewed_by')
    else:
        requests = ChangeRequest.objects.filter(status=status_filter, requested_by=request.user).select_related('requested_by', 'reviewed_by')

    return render(request, 'supply_chain/change_requests.html', {
        'requests': requests,
        'status_filter': status_filter,
        'is_admin': is_supply_chain_admin(request.user),
    })


@supply_chain_required
def change_request_detail(request, pk):
    if is_supply_chain_admin(request.user):
        req = get_object_or_404(ChangeRequest, pk=pk)
    else:
        req = get_object_or_404(ChangeRequest, pk=pk, requested_by=request.user)

    return render(request, 'supply_chain/change_request_detail.html', {
        'req': req,
        'old_and_new': req.get_old_and_new_values(),
        'proposed_fields': req.get_proposed_fields(),
        'is_admin': is_supply_chain_admin(request.user),
    })


@supply_chain_required
@require_POST
def change_request_approve(request, pk):
    if not is_supply_chain_admin(request.user):
        messages.error(request, 'Only administrators can approve change requests.')
        return redirect('supply_chain:change_requests')

    req = get_object_or_404(ChangeRequest, pk=pk, status='PENDING')
    try:
        req.apply(request.user)
        messages.success(request, f'Change request #{req.id} approved and applied successfully.')
    except Exception as e:
        messages.error(request, f'Error applying change: {str(e)}')

    return redirect('supply_chain:change_requests')


@supply_chain_required
@require_POST
def change_request_reject(request, pk):
    if not is_supply_chain_admin(request.user):
        messages.error(request, 'Only administrators can reject change requests.')
        return redirect('supply_chain:change_requests')

    req = get_object_or_404(ChangeRequest, pk=pk, status='PENDING')
    reason = request.POST.get('rejection_reason', '').strip()
    if not reason:
        messages.error(request, 'A rejection reason is required.')
        return redirect('supply_chain:change_request_detail', pk=req.pk)

    req.status = 'REJECTED'
    req.reviewed_by = request.user
    req.reviewed_at = timezone.now()
    req.rejection_reason = reason
    req.save()

    messages.success(request, f'Change request #{req.id} rejected.')
    return redirect('supply_chain:change_requests')


# Edit/Delete Views for other supply chain records
@supply_chain_required
def monthly_demand_edit(request, pk):
    demand = get_object_or_404(StockDemand.objects.filter(is_active=True), pk=pk)
    if request.method == 'POST':
        form = StockDemandForm(request.POST, instance=demand)
        if form.is_valid():
            if is_supply_chain_admin(request.user):
                form.save()
                ChangeRequest.objects.create(
                    model_name='StockDemand',
                    action='UPDATE',
                    target_id=demand.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='APPROVED',
                    reviewed_by=request.user,
                    reviewed_at=timezone.now(),
                )
                messages.success(request, 'Monthly demand entry updated.')
            else:
                ChangeRequest.objects.create(
                    model_name='StockDemand',
                    action='UPDATE',
                    target_id=demand.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='PENDING',
                )
                messages.success(request, 'Update request submitted for admin review.')
            return redirect('supply_chain:monthly_demand')
    else:
        form = StockDemandForm(instance=demand)
    return render(request, 'supply_chain/monthly_demand_edit.html', {'demand': demand, 'form': form})


@supply_chain_required
def monthly_demand_delete(request, pk):
    demand = get_object_or_404(StockDemand.objects.filter(is_active=True), pk=pk)
    if request.method == 'POST':
        if is_supply_chain_admin(request.user):
            demand.is_active = False
            demand.save(update_fields=['is_active'])
            ChangeRequest.objects.create(
                model_name='StockDemand',
                action='DELETE',
                target_id=demand.pk,
                proposed_data={},
                requested_by=request.user,
                status='APPROVED',
                reviewed_by=request.user,
                reviewed_at=timezone.now(),
            )
            messages.success(request, 'Monthly demand entry archived.')
        else:
            ChangeRequest.objects.create(
                model_name='StockDemand',
                action='DELETE',
                target_id=demand.pk,
                proposed_data={},
                requested_by=request.user,
                status='PENDING',
            )
            messages.success(request, 'Deletion request submitted for admin review.')
        return redirect('supply_chain:monthly_demand')
    return render(request, 'supply_chain/confirm_delete.html', {
        'object': demand,
        'title': 'Delete Monthly Demand',
        'back_url': reverse('supply_chain:monthly_demand')
    })


@supply_chain_required
def transaction_edit(request, pk):
    txn = get_object_or_404(StockTransaction.objects.filter(is_active=True), pk=pk)
    page_key_map = {'OPENING': 'opening', 'RECEIVING': 'receiving', 'ISSUANCE': 'issuance', 'ADJUSTMENT': 'adjustment'}
    page_key = page_key_map.get(txn.transaction_type, 'dashboard')

    if request.method == 'POST':
        form = StockTransactionForm(request.POST, instance=txn)
        if form.is_valid():
            if is_supply_chain_admin(request.user):
                form.save()
                ChangeRequest.objects.create(
                    model_name='StockTransaction',
                    action='UPDATE',
                    target_id=txn.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='APPROVED',
                    reviewed_by=request.user,
                    reviewed_at=timezone.now(),
                )
                messages.success(request, 'Transaction entry updated.')
            else:
                ChangeRequest.objects.create(
                    model_name='StockTransaction',
                    action='UPDATE',
                    target_id=txn.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='PENDING',
                )
                messages.success(request, 'Update request submitted for admin review.')
            return redirect('supply_chain:' + page_key)
    else:
        form = StockTransactionForm(instance=txn)
    return render(request, 'supply_chain/transaction_edit.html', {
        'transaction': txn,
        'form': form,
        'back_url': reverse('supply_chain:' + page_key)
    })


@supply_chain_required
def transaction_delete(request, pk):
    txn = get_object_or_404(StockTransaction.objects.filter(is_active=True), pk=pk)
    page_key_map = {'OPENING': 'opening', 'RECEIVING': 'receiving', 'ISSUANCE': 'issuance', 'ADJUSTMENT': 'adjustment'}
    page_key = page_key_map.get(txn.transaction_type, 'dashboard')

    if request.method == 'POST':
        if is_supply_chain_admin(request.user):
            txn.is_active = False
            txn.save(update_fields=['is_active'])
            ChangeRequest.objects.create(
                model_name='StockTransaction',
                action='DELETE',
                target_id=txn.pk,
                proposed_data={},
                requested_by=request.user,
                status='APPROVED',
                reviewed_by=request.user,
                reviewed_at=timezone.now(),
            )
            messages.success(request, 'Transaction archived.')
        else:
            ChangeRequest.objects.create(
                model_name='StockTransaction',
                action='DELETE',
                target_id=txn.pk,
                proposed_data={},
                requested_by=request.user,
                status='PENDING',
            )
            messages.success(request, 'Deletion request submitted for admin review.')
        return redirect('supply_chain:' + page_key)
    return render(request, 'supply_chain/confirm_delete.html', {
        'object': txn,
        'title': 'Delete Transaction',
        'back_url': reverse('supply_chain:' + page_key)
    })


@supply_chain_required
def physical_count_edit(request, pk):
    pc = get_object_or_404(PhysicalStockCount.objects.filter(is_active=True), pk=pk)
    if request.method == 'POST':
        form = PhysicalStockCountForm(request.POST, instance=pc)
        if form.is_valid():
            if is_supply_chain_admin(request.user):
                accuracy = compute_inventory_accuracy(form.cleaned_data['physical_sheet_qty'], pc.system_sheet_qty)
                instance = form.save(commit=False)
                instance.accuracy_percent = accuracy
                instance.save()
                ChangeRequest.objects.create(
                    model_name='PhysicalStockCount',
                    action='UPDATE',
                    target_id=pc.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='APPROVED',
                    reviewed_by=request.user,
                    reviewed_at=timezone.now(),
                )
                messages.success(request, 'Physical stock count updated.')
            else:
                ChangeRequest.objects.create(
                    model_name='PhysicalStockCount',
                    action='UPDATE',
                    target_id=pc.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='PENDING',
                )
                messages.success(request, 'Update request submitted for admin review.')
            return redirect('supply_chain:physical_counts')
    else:
        form = PhysicalStockCountForm(instance=pc)
    return render(request, 'supply_chain/physical_count_edit.html', {'pc': pc, 'form': form})


@supply_chain_required
def physical_count_delete(request, pk):
    pc = get_object_or_404(PhysicalStockCount.objects.filter(is_active=True), pk=pk)
    if request.method == 'POST':
        if is_supply_chain_admin(request.user):
            pc.is_active = False
            pc.save(update_fields=['is_active'])
            ChangeRequest.objects.create(
                model_name='PhysicalStockCount',
                action='DELETE',
                target_id=pc.pk,
                proposed_data={},
                requested_by=request.user,
                status='APPROVED',
                reviewed_by=request.user,
                reviewed_at=timezone.now(),
            )
            messages.success(request, 'Physical stock count entry archived.')
        else:
            ChangeRequest.objects.create(
                model_name='PhysicalStockCount',
                action='DELETE',
                target_id=pc.pk,
                proposed_data={},
                requested_by=request.user,
                status='PENDING',
            )
            messages.success(request, 'Deletion request submitted for admin review.')
        return redirect('supply_chain:physical_counts')
    return render(request, 'supply_chain/confirm_delete.html', {
        'object': pc,
        'title': 'Delete Physical Count',
        'back_url': reverse('supply_chain:physical_counts')
    })


@supply_chain_required
def item_delete(request, pk):
    item = get_object_or_404(RawMaterialSku.objects.filter(is_active=True), pk=pk)
    if request.method == 'POST':
        if is_supply_chain_admin(request.user):
            item.is_active = False
            item.save(update_fields=['is_active'])
            ChangeRequest.objects.create(
                model_name='RawMaterialSku',
                action='DELETE',
                target_id=item.pk,
                proposed_data={},
                requested_by=request.user,
                status='APPROVED',
                reviewed_by=request.user,
                reviewed_at=timezone.now(),
            )
            messages.success(request, f'Raw Material SKU {item.sku} archived.')
        else:
            ChangeRequest.objects.create(
                model_name='RawMaterialSku',
                action='DELETE',
                target_id=item.pk,
                proposed_data={},
                requested_by=request.user,
                status='PENDING',
            )
            messages.success(request, 'Deletion request submitted for admin review.')
        return redirect('supply_chain:items')
    return render(request, 'supply_chain/confirm_delete.html', {
        'object': item,
        'title': 'Delete Raw Material SKU',
        'back_url': reverse('supply_chain:items')
    })


@supply_chain_required
@require_POST
def item_reactivate(request, pk):
    """Restore a soft-deleted (archived) Raw Material SKU so it maps again."""
    if not is_supply_chain_admin(request.user):
        messages.error(request, 'Only administrators can reactivate SKUs.')
        return redirect(f"{reverse('supply_chain:items')}?show=archived")

    item = get_object_or_404(RawMaterialSku.objects.filter(is_active=False), pk=pk)
    item.is_active = True
    item.save(update_fields=['is_active'])
    ChangeRequest.objects.create(
        model_name='RawMaterialSku',
        action='UPDATE',
        target_id=item.pk,
        proposed_data={'is_active': True},
        requested_by=request.user,
        status='APPROVED',
        reviewed_by=request.user,
        reviewed_at=timezone.now(),
    )
    messages.success(request, f'Raw Material SKU {item.sku} reactivated.')
    return redirect('supply_chain:items')


@supply_chain_required
@require_POST
def bulk_delete(request):
    model_name = request.POST.get('model_name')
    redirect_url = request.POST.get('redirect_url') or 'supply_chain:dashboard'
    selected_ids = request.POST.getlist('selected_ids')

    ALLOWED_MODELS = {'RawMaterialSku', 'StockDemand', 'StockTransaction', 'PhysicalStockCount'}
    if model_name not in ALLOWED_MODELS:
        messages.error(request, "Invalid model specified for deletion.")
        return redirect(redirect_url)

    if not selected_ids:
        messages.warning(request, "No items were selected for deletion.")
        return redirect(redirect_url)

    from django.apps import apps
    model_class = apps.get_model('supply_chain', model_name)

    is_admin = is_supply_chain_admin(request.user)
    processed = 0

    for pk in selected_ids:
        try:
            instance = model_class.objects.get(pk=pk, is_active=True)
            if is_admin:
                instance.is_active = False
                instance.save(update_fields=['is_active'])
                ChangeRequest.objects.create(
                    model_name=model_name,
                    action='DELETE',
                    target_id=instance.pk,
                    proposed_data={},
                    requested_by=request.user,
                    reviewed_by=request.user,
                    reviewed_at=timezone.now(),
                    status='APPROVED',
                )
            else:
                ChangeRequest.objects.create(
                    model_name=model_name,
                    action='DELETE',
                    target_id=instance.pk,
                    proposed_data={},
                    requested_by=request.user,
                    status='PENDING',
                )
            processed += 1
        except model_class.DoesNotExist:
            continue

    if processed > 0:
        if is_admin:
            messages.success(request, f"Successfully archived {processed} record(s).")
        else:
            messages.success(request, f"Submitted deletion requests for {processed} record(s) for admin review.")
    else:
        messages.error(request, "No valid records were processed for deletion.")

    return redirect(redirect_url)


@supply_chain_required
def issuance_approve(request, pk):
    txn = get_object_or_404(StockTransaction.objects.filter(is_active=True), pk=pk, transaction_type='ISSUANCE', is_approved=False)
    txn.is_approved = True
    txn.save(update_fields=['is_approved'])
    messages.success(request, f"Approved issuance for {txn.raw_material_sku.sku if txn.raw_material_sku else 'N/A'}")
    return redirect('supply_chain:issuance')


@supply_chain_required
@require_POST
def issuance_bulk_approve(request):
    selected_ids = request.POST.getlist('selected_ids')
    if not selected_ids:
        messages.warning(request, "No entries selected.")
        return redirect('supply_chain:issuance')
    
    txns = StockTransaction.objects.filter(pk__in=selected_ids, transaction_type='ISSUANCE', is_approved=False, is_active=True)
    count = txns.count()
    txns.update(is_approved=True)
    messages.success(request, f"Successfully approved {count} issuance(s).")
    return redirect('supply_chain:issuance')


# ---------------------------------------------------------------------------
# Item Request module
# ---------------------------------------------------------------------------

def _item_request_role(user):
    if user.is_superuser:
        return 'admin'
    profile = getattr(user, 'profile', None)
    return (getattr(profile, 'role', '') or '').strip().lower()


def _is_manager(user):
    return _item_request_role(user) in ('admin', 'manager')


def _is_supply_chain(user):
    return _item_request_role(user) in ('admin', 'supply_chain')


def _stage_decisions(approvals):
    """Latest recorded decision for each review stage, for the review tabs/print."""
    decision_actions = ('APPROVE', 'REJECT', 'REVISE')
    ordered = list(approvals)
    return {
        stage: next(
            (a for a in reversed(ordered) if a.stage == stage and a.action in decision_actions),
            None,
        )
        for stage in ('MANAGER', 'SUPPLY_CHAIN')
    }


@item_request_access_required
def item_request_list(request):
    requests_qs = (
        ItemRequest.objects.filter(is_active=True)
        .select_related('request_type', 'department', 'machine', 'raised_by')
        .prefetch_related('approvals')
        .order_by('-created_at')
    )

    status = request.GET.get('status')
    if status:
        requests_qs = requests_qs.filter(status=status)
    req_type = request.GET.get('type')
    if req_type:
        requests_qs = requests_qs.filter(request_type_id=req_type)
    department = request.GET.get('department')
    if department:
        requests_qs = requests_qs.filter(department_id=department)
    if request.GET.get('mine') == '1':
        requests_qs = requests_qs.filter(raised_by=request.user)

    if request.GET.get('export') == 'xlsx':
        return export_item_requests(requests_qs, filename='item_requests.xlsx')

    can_manage_reviews = _is_manager(request.user)
    can_manage_sc = _is_supply_chain(request.user)

    rows = []
    for req in requests_qs:
        days_in_stage, sla_limit, breached = sla_status(req, req.approvals.all())
        rows.append({'request': req, 'sla_breached': breached, 'days_in_stage': days_in_stage, 'sla_limit': sla_limit})

    # Role-scoped queues, so each person lands on the slice of work that is
    # actually theirs instead of scanning one flat list for it.
    needs_action_rows = [
        r for r in rows if (
            (can_manage_reviews and r['request'].status == 'MGR_REVIEW')
            or (can_manage_sc and r['request'].status == 'SC_REVIEW')
            or (r['request'].raised_by_id == request.user.id and r['request'].status == 'NEEDS_REVISION')
        )
    ]
    mine_rows = [r for r in rows if r['request'].raised_by_id == request.user.id]
    in_progress_rows = [
        r for r in rows if r['request'].status in ('SUBMITTED', 'MGR_REVIEW', 'SC_REVIEW', 'NEEDS_REVISION')
    ]
    procurement_rows = [r for r in rows if r['request'].status in ('APPROVED', 'IN_PROCUREMENT', 'RECEIVED')]
    closed_rows = [r for r in rows if r['request'].status in ('CLOSED', 'REJECTED')]

    show_needs_action_tab = can_manage_reviews or can_manage_sc or bool(needs_action_rows)
    if show_needs_action_tab and needs_action_rows:
        active_tab = 'needs_action'
    elif not (can_manage_reviews or can_manage_sc):
        active_tab = 'mine'
    else:
        active_tab = 'all'

    return render(request, 'supply_chain/item_request/list.html', {
        'rows': rows,
        'needs_action_rows': needs_action_rows,
        'mine_rows': mine_rows,
        'in_progress_rows': in_progress_rows,
        'procurement_rows': procurement_rows,
        'closed_rows': closed_rows,
        'show_needs_action_tab': show_needs_action_tab,
        'active_tab': active_tab,
        'types': ItemRequestType.objects.filter(is_active=True),
        'departments': ItemRequestDepartment.objects.filter(is_active=True),
        'status_choices': ItemRequest.STATUS_CHOICES,
        'filters': request.GET,
        'is_superuser': request.user.is_superuser,
    })


@item_request_access_required
def item_request_create(request):
    duplicates = []
    if request.method == 'POST':
        form = ItemRequestForm(request.POST, request.FILES)
        if form.is_valid():
            department = form.cleaned_data['department']
            duplicates = ItemRequest.objects.filter(
                pk__in=ItemRequest.find_open_duplicates(form.cleaned_data['item_title'], department.id).values_list('pk', flat=True)
            ).select_related('request_type') if request.POST.get('confirm_duplicate') != '1' else []

            if duplicates:
                messages.warning(
                    request,
                    'A similar open item request already exists for this department. '
                    'Review it below, or tick "Submit anyway" to proceed.',
                )
            else:
                item_request = form.save(commit=False)
                item_request.raised_by = request.user
                item_request.save()
                submit_request(item_request, request.user)
                messages.success(request, f'Item request "{item_request.item_title}" submitted for manager review.')
                return redirect('supply_chain:item_request_detail', pk=item_request.pk)
    else:
        form = ItemRequestForm()

    return render(request, 'supply_chain/item_request/form.html', {
        'form': form,
        'is_create': True,
        'duplicates': duplicates,
    })


@item_request_access_required
def item_request_detail(request, pk):
    item_request = get_object_or_404(
        ItemRequest.objects.filter(is_active=True).select_related('request_type', 'department', 'machine', 'raised_by'),
        pk=pk,
    )
    approvals = item_request.approvals.select_related('actor').all()
    procurement = getattr(item_request, 'procurement', None)

    can_review_manager = item_request.status == 'MGR_REVIEW' and _is_manager(request.user)
    can_review_sc = item_request.status == 'SC_REVIEW' and _is_supply_chain(request.user)
    can_resubmit = item_request.status in ('NEEDS_REVISION',) and item_request.raised_by_id == request.user.id
    can_edit_procurement = item_request.status in ('IN_PROCUREMENT', 'RECEIVED', 'CLOSED') and _is_supply_chain(request.user)
    can_change_edit = item_request.status in ('APPROVED', 'IN_PROCUREMENT', 'RECEIVED', 'CLOSED')

    days_in_stage, sla_limit, breached = sla_status(item_request, approvals)

    # Latest recorded decision per review stage, for the in-built review tabs.
    decisions = _stage_decisions(approvals)
    manager_decision = decisions['MANAGER']
    sc_decision = decisions['SUPPLY_CHAIN']
    active_review_tab = 'sc' if (can_review_sc or item_request.status == 'SC_REVIEW') else 'manager'

    # The REVISE entry that put the request into NEEDS_REVISION, for the revision banner.
    revision_note = approvals.last() if item_request.status == 'NEEDS_REVISION' else None

    if can_resubmit:
        active_page_tab = 'details'
    elif can_review_manager or can_review_sc:
        active_page_tab = 'review'
    else:
        active_page_tab = 'details'

    return render(request, 'supply_chain/item_request/detail.html', {
        'item_request': item_request,
        'approvals': approvals,
        'procurement': procurement,
        'review_form': ItemRequestReviewForm(),
        'manager_decision': manager_decision,
        'sc_decision': sc_decision,
        'active_review_tab': active_review_tab,
        'active_page_tab': active_page_tab,
        'revision_note': revision_note,
        'can_review_manager': can_review_manager,
        'can_review_sc': can_review_sc,
        'can_resubmit': can_resubmit,
        'can_edit_procurement': can_edit_procurement,
        'can_change_edit': can_change_edit,
        'is_superuser': request.user.is_superuser,
        'sla_days_in_stage': days_in_stage,
        'sla_limit': sla_limit,
        'sla_breached': breached,
    })


@item_request_access_required
@require_POST
def item_request_review(request, pk):
    item_request = get_object_or_404(ItemRequest, pk=pk)

    if item_request.status == 'MGR_REVIEW' and _is_manager(request.user):
        stage = 'MANAGER'
    elif item_request.status == 'SC_REVIEW' and _is_supply_chain(request.user):
        stage = 'SUPPLY_CHAIN'
    else:
        raise PermissionDenied

    form = ItemRequestReviewForm(request.POST)
    if form.is_valid():
        review_request(
            item_request,
            request.user,
            stage=stage,
            action=form.cleaned_data['action'],
            comment=form.cleaned_data['comment'],
        )
        messages.success(request, f'Request {item_request.request_no or item_request.pk} updated.')
    else:
        messages.error(request, 'Please provide a comment for this decision.')

    return redirect('supply_chain:item_request_detail', pk=pk)


@item_request_access_required
def item_request_resubmit(request, pk):
    item_request = get_object_or_404(ItemRequest, pk=pk, raised_by=request.user)
    if item_request.status not in ('NEEDS_REVISION',):
        raise PermissionDenied

    if request.method == 'POST':
        form = ItemRequestForm(request.POST, request.FILES, instance=item_request)
        if form.is_valid():
            form.save()
            resubmit_request(item_request, request.user)
            messages.success(request, f'Request {item_request.request_no or item_request.pk} resubmitted.')
            return redirect('supply_chain:item_request_detail', pk=item_request.pk)
    else:
        form = ItemRequestForm(instance=item_request)

    return render(request, 'supply_chain/item_request/form.html', {
        'form': form,
        'is_create': False,
        'item_request': item_request,
    })


@item_request_access_required
def item_request_procurement(request, pk):
    item_request = get_object_or_404(ItemRequest, pk=pk)
    if not _is_supply_chain(request.user):
        raise PermissionDenied
    procurement = getattr(item_request, 'procurement', None)
    if procurement is None:
        raise PermissionDenied

    if request.method == 'POST':
        form = ItemProcurementTimelineForm(request.POST, instance=procurement)
        if form.is_valid():
            timeline = form.save(commit=False)
            timeline.updated_by = request.user
            timeline.save()
            if timeline.received_date and item_request.status == 'IN_PROCUREMENT':
                item_request.status = 'RECEIVED'
                item_request.save(update_fields=['status'])
            messages.success(request, 'Procurement timeline updated.')
            return redirect('supply_chain:item_request_detail', pk=pk)
    else:
        form = ItemProcurementTimelineForm(instance=procurement)

    return render(request, 'supply_chain/item_request/procurement_form.html', {
        'form': form,
        'item_request': item_request,
        'procurement': procurement,
        'quotes': procurement.quotes.all(),
        'quote_form': ItemRequestQuoteForm(),
    })


@item_request_access_required
@require_POST
def item_request_quote_add(request, pk):
    item_request = get_object_or_404(ItemRequest, pk=pk)
    if not _is_supply_chain(request.user):
        raise PermissionDenied
    procurement = getattr(item_request, 'procurement', None)
    if procurement is None:
        raise PermissionDenied

    form = ItemRequestQuoteForm(request.POST, request.FILES)
    if form.is_valid():
        quote = form.save(commit=False)
        quote.procurement = procurement
        quote.uploaded_by = request.user
        quote.save()
        messages.success(request, f'Quote from {quote.supplier} added.')
    else:
        messages.error(request, 'Please correct the errors in the quote form.')
    return redirect('supply_chain:item_request_procurement', pk=pk)


@item_request_access_required
@require_POST
def item_request_quote_delete(request, pk, quote_id):
    item_request = get_object_or_404(ItemRequest, pk=pk)
    if not _is_supply_chain(request.user):
        raise PermissionDenied
    quote = get_object_or_404(ItemRequestQuote, pk=quote_id, procurement__request=item_request)
    quote.delete()
    messages.success(request, 'Quote removed.')
    return redirect('supply_chain:item_request_procurement', pk=pk)


@item_request_access_required
@require_POST
def item_request_type_add(request):
    form = ItemRequestTypeQuickForm(request.POST)
    if form.is_valid():
        obj = form.save()
        return JsonResponse({'ok': True, 'id': obj.id, 'name': obj.name})
    return JsonResponse({'ok': False, 'errors': form.errors}, status=400)


@item_request_access_required
@require_POST
def item_request_department_add(request):
    form = DepartmentQuickForm(request.POST)
    if form.is_valid():
        obj = form.save()
        return JsonResponse({'ok': True, 'id': obj.id, 'name': obj.name})
    return JsonResponse({'ok': False, 'errors': form.errors}, status=400)


@item_request_access_required
def item_request_kpi_dashboard(request):
    requests_qs = ItemRequest.objects.filter(is_active=True).select_related('request_type', 'department')

    req_type = request.GET.get('type')
    if req_type:
        requests_qs = requests_qs.filter(request_type_id=req_type)
    department = request.GET.get('department')
    if department:
        requests_qs = requests_qs.filter(department_id=department)

    rows, summary = build_item_request_kpi_data(requests_qs)
    summary['sla_breached_count'] = sum(
        1 for r in requests_qs if r.is_open and sla_status(r)[2]
    )

    return render(request, 'supply_chain/item_request/kpi_dashboard.html', {
        'kpi_rows': rows,
        'kpi_summary': summary,
        'types': ItemRequestType.objects.filter(is_active=True),
        'departments': ItemRequestDepartment.objects.filter(is_active=True),
        'filters': request.GET,
    })


@item_request_access_required
def item_request_print(request, pk):
    item_request = get_object_or_404(
        ItemRequest.objects.select_related('request_type', 'department', 'machine', 'raised_by'),
        pk=pk,
    )
    decisions = _stage_decisions(item_request.approvals.select_related('actor').all())
    return render(request, 'supply_chain/item_request/print.html', {
        'item_request': item_request,
        'manager_decision': decisions['MANAGER'],
        'sc_decision': decisions['SUPPLY_CHAIN'],
        'procurement': getattr(item_request, 'procurement', None),
    })


@item_request_access_required
@require_POST
def item_request_delete(request, pk):
    if not request.user.is_superuser:
        raise PermissionDenied
    item_request = get_object_or_404(ItemRequest, pk=pk, is_active=True)
    _soft_delete_item_request(item_request, request.user, request.POST.get('reason', ''))
    messages.success(request, f'Item request {item_request.request_no or item_request.pk} deleted.')
    return redirect('supply_chain:item_request_list')


def _soft_delete_item_request(item_request, user, reason=''):
    item_request.is_active = False
    item_request.deleted_by = user
    item_request.deleted_at = timezone.now()
    item_request.save(update_fields=['is_active', 'deleted_by', 'deleted_at'])
    ItemRequestApproval.objects.create(
        request=item_request, actor=user, action='DELETE', stage='SUPPLY_CHAIN', comment=reason,
    )


@item_request_access_required
@require_POST
def item_request_bulk_delete(request):
    if not request.user.is_superuser:
        raise PermissionDenied
    selected_ids = request.POST.getlist('selected_ids')
    if not selected_ids:
        messages.warning(request, 'No item requests selected.')
        return redirect('supply_chain:item_request_list')

    reason = request.POST.get('reason', '')
    requests_qs = ItemRequest.objects.filter(pk__in=selected_ids, is_active=True)
    count = 0
    for item_request in requests_qs:
        _soft_delete_item_request(item_request, request.user, reason)
        count += 1
    messages.success(request, f'Deleted {count} item request(s).')
    return redirect('supply_chain:item_request_list')


@item_request_access_required
def item_request_change_edit(request, pk):
    """Post-approval edit of an item request, routed through change management for non-admins."""
    item_request = get_object_or_404(ItemRequest, pk=pk, is_active=True)
    if item_request.status not in ('APPROVED', 'IN_PROCUREMENT', 'RECEIVED', 'CLOSED'):
        raise PermissionDenied

    if request.method == 'POST':
        form = ItemRequestForm(request.POST, request.FILES, instance=item_request)
        if form.is_valid():
            if is_supply_chain_admin(request.user):
                form.save()
                ChangeRequest.objects.create(
                    model_name='ItemRequest',
                    action='UPDATE',
                    target_id=item_request.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                    status='APPROVED',
                    reviewed_by=request.user,
                    reviewed_at=timezone.now(),
                )
                ItemRequestApproval.objects.create(
                    request=item_request, actor=request.user, action='EDIT', stage='SUPPLY_CHAIN',
                )
                messages.success(request, f'Updated {item_request.request_no or item_request.pk}.')
            else:
                ChangeRequest.objects.create(
                    model_name='ItemRequest',
                    action='UPDATE',
                    target_id=item_request.pk,
                    proposed_data=serialize_cleaned_data(form.cleaned_data),
                    requested_by=request.user,
                )
                ItemRequestApproval.objects.create(
                    request=item_request, actor=request.user, action='EDIT_REQUESTED', stage='SUPPLY_CHAIN',
                )
                messages.info(request, 'Change submitted for admin review.')
            return redirect('supply_chain:item_request_detail', pk=pk)
    else:
        form = ItemRequestForm(instance=item_request)

    return render(request, 'supply_chain/item_request/form.html', {
        'form': form,
        'is_create': False,
        'is_change_edit': True,
        'item_request': item_request,
    })

