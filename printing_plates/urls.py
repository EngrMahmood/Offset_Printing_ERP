from django.urls import path

from . import views

app_name = 'printing_plates'

urlpatterns = [
    path('', views.PlateDashboardView.as_view(), name='dashboard'),
    path('requests/', views.PlateRequestListView.as_view(), name='request_list'),
    path('queue/', views.PlateQueueView.as_view(), name='queue'),
    path('sent/', views.PlateSentListView.as_view(), name='sent_list'),
    path('received/', views.PlateReceivedListView.as_view(), name='received_list'),
    path('request/add/', views.PlateRequestCreateView.as_view(), name='request_add'),
    path('request/<int:pk>/', views.PlateRequestDetailView.as_view(), name='request_detail'),
    path('request/<int:pk>/action/', views.PlateRequestActionView.as_view(), name='request_action'),
    path('admin/cancel-stale/', views.BulkCancelStalePlateRequestsView.as_view(), name='bulk_cancel_stale'),
]
