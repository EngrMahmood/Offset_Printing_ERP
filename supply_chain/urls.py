from django.urls import path

from . import views

app_name = 'supply_chain'

urlpatterns = [
    path('', views.dashboard, name='home'),
    path('dashboard/', views.dashboard, name='dashboard'),
    
    # Change Management Routes
    path('change-requests/', views.change_requests_list, name='change_requests'),
    path('change-requests/<int:pk>/', views.change_request_detail, name='change_request_detail'),
    path('change-requests/<int:pk>/approve/', views.change_request_approve, name='change_request_approve'),
    path('change-requests/<int:pk>/reject/', views.change_request_reject, name='change_request_reject'),
    path('bulk-delete/', views.bulk_delete, name='bulk_delete'),

    path('items/', views.item_list, name='items'),
    path('items/<int:pk>/edit/', views.item_edit, name='item_edit'),
    path('items/<int:pk>/delete/', views.item_delete, name='item_delete'),
    path('items/<int:pk>/reactivate/', views.item_reactivate, name='item_reactivate'),
    path('raw-material-skus/quick-add/', views.quick_add_raw_material_sku, name='quick_add_raw_material_sku'),
    
    path('monthly-demand/', views.monthly_demand, name='monthly_demand'),
    path('monthly-demand/<int:pk>/edit/', views.monthly_demand_edit, name='monthly_demand_edit'),
    path('monthly-demand/<int:pk>/delete/', views.monthly_demand_delete, name='monthly_demand_delete'),

    path('opening/', views.transaction_page, {'page_key': 'opening'}, name='opening'),
    path('receiving/', views.transaction_page, {'page_key': 'receiving'}, name='receiving'),
    path('issuance/', views.transaction_page, {'page_key': 'issuance'}, name='issuance'),
    path('issuance/<int:pk>/approve/', views.issuance_approve, name='issuance_approve'),
    path('issuance/bulk-approve/', views.issuance_bulk_approve, name='issuance_bulk_approve'),
    path('adjustment/', views.transaction_page, {'page_key': 'adjustment'}, name='adjustment'),
    path('transaction/<int:pk>/edit/', views.transaction_edit, name='transaction_edit'),
    path('transaction/<int:pk>/delete/', views.transaction_delete, name='transaction_delete'),

    path('reports/consumption/', views.consumption_reports, name='consumption_reports'),
    path('kpis/', views.kpi_dashboard, name='kpi_dashboard'),
    path('job-card-links/', views.jc_links, name='jc_links'),
    
    path('physical-counts/', views.physical_counts, name='physical_counts'),
    path('physical-counts/<int:pk>/edit/', views.physical_count_edit, name='physical_count_edit'),
    path('physical-counts/<int:pk>/delete/', views.physical_count_delete, name='physical_count_delete'),

    path('demand-gap/', views.demand_gap, name='demand_gap'),

    # Item Request Module
    path('item-requests/', views.item_request_list, name='item_request_list'),
    path('item-requests/new/', views.item_request_create, name='item_request_create'),
    path('item-requests/<int:pk>/', views.item_request_detail, name='item_request_detail'),
    path('item-requests/<int:pk>/review/', views.item_request_review, name='item_request_review'),
    path('item-requests/<int:pk>/resubmit/', views.item_request_resubmit, name='item_request_resubmit'),
    path('item-requests/<int:pk>/procurement/', views.item_request_procurement, name='item_request_procurement'),
    path('item-requests/type/add/', views.item_request_type_add, name='item_request_type_add'),
    path('item-requests/department/add/', views.item_request_department_add, name='item_request_department_add'),
    path('item-requests/kpis/', views.item_request_kpi_dashboard, name='item_request_kpi_dashboard'),
    path('item-requests/<int:pk>/print/', views.item_request_print, name='item_request_print'),
    path('item-requests/<int:pk>/delete/', views.item_request_delete, name='item_request_delete'),
    path('item-requests/bulk-delete/', views.item_request_bulk_delete, name='item_request_bulk_delete'),
    path('item-requests/<int:pk>/change-edit/', views.item_request_change_edit, name='item_request_change_edit'),
    path('item-requests/<int:pk>/quotes/add/', views.item_request_quote_add, name='item_request_quote_add'),
    path('item-requests/<int:pk>/quotes/<int:quote_id>/delete/', views.item_request_quote_delete, name='item_request_quote_delete'),
]
