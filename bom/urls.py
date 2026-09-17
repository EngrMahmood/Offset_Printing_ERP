from django.urls import path

from . import views

app_name = 'bom'

urlpatterns = [
    path('', views.dashboard, name='dashboard'),

    path('recipe-wizard/', views.recipe_wizard, name='recipe_wizard'),

    path('raw-items/', views.raw_item_list, name='raw_item_list'),
    path('raw-items/new/', views.raw_item_form, name='raw_item_create'),
    path('raw-items/<int:pk>/edit/', views.raw_item_form, name='raw_item_edit'),
    path('raw-items/import/', views.raw_item_import, name='raw_item_import'),
    path('raw-items/template/', views.raw_item_template_download, name='raw_item_template_download'),

    path('sku-bom-template/', views.sku_bom_template_download, name='sku_bom_template_download'),

    path('spec-attributes/', views.spec_attribute_list, name='spec_attribute_list'),
    path('spec-attributes/new/', views.spec_attribute_form, name='spec_attribute_create'),
    path('spec-attributes/<int:pk>/edit/', views.spec_attribute_form, name='spec_attribute_edit'),

    path('templates/', views.template_list, name='template_list'),
    path('templates/new/', views.template_form, name='template_create'),
    path('templates/<int:pk>/', views.template_detail, name='template_detail'),
    path('templates/<int:pk>/edit/', views.template_form, name='template_edit'),
    path('templates/<int:pk>/submit/', views.template_submit, name='template_submit'),
    path('templates/<int:pk>/approve/', views.template_approve, name='template_approve'),

    path('test-match/', views.test_match, name='test_match'),

    path('boms/', views.bom_list, name='bom_list'),
    path('boms/<int:pk>/', views.bom_detail, name='bom_detail'),
    path('boms/<int:pk>/regenerate/', views.bom_regenerate, name='bom_regenerate'),
    path('boms/<int:pk>/submit/', views.bom_submit, name='bom_submit'),
    path('boms/<int:pk>/review/', views.bom_review, name='bom_review'),
    path('boms/<int:pk>/approve/', views.bom_approve, name='bom_approve'),
    path('boms/<int:pk>/reject/', views.bom_reject, name='bom_reject'),

    path('requirements/', views.requirements_report, name='requirements_report'),
    path('requirements/raise-item-requests/', views.raise_item_requests, name='raise_item_requests'),
]
