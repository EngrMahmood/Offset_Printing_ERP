from django.urls import path

from . import views

app_name = 'supply_chain'

urlpatterns = [
    path('', views.dashboard, name='home'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('items/', views.item_list, name='items'),
    path('items/<int:pk>/edit/', views.item_edit, name='item_edit'),
    path('monthly-demand/', views.monthly_demand, name='monthly_demand'),
    path('opening/', views.transaction_page, {'page_key': 'opening'}, name='opening'),
    path('receiving/', views.transaction_page, {'page_key': 'receiving'}, name='receiving'),
    path('issuance/', views.transaction_page, {'page_key': 'issuance'}, name='issuance'),
    path('adjustment/', views.transaction_page, {'page_key': 'adjustment'}, name='adjustment'),
    path('reports/consumption/', views.consumption_reports, name='consumption_reports'),
    path('kpis/', views.kpi_dashboard, name='kpi_dashboard'),
    path('job-card-links/', views.jc_links, name='jc_links'),
    path('physical-counts/', views.physical_counts, name='physical_counts'),
]
