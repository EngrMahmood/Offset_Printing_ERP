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
    path('raw-material-skus/quick-add/', views.quick_add_raw_material_sku, name='quick_add_raw_material_sku'),
    
    path('monthly-demand/', views.monthly_demand, name='monthly_demand'),
    path('monthly-demand/<int:pk>/edit/', views.monthly_demand_edit, name='monthly_demand_edit'),
    path('monthly-demand/<int:pk>/delete/', views.monthly_demand_delete, name='monthly_demand_delete'),

    path('opening/', views.transaction_page, {'page_key': 'opening'}, name='opening'),
    path('receiving/', views.transaction_page, {'page_key': 'receiving'}, name='receiving'),
    path('issuance/', views.transaction_page, {'page_key': 'issuance'}, name='issuance'),
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
]
