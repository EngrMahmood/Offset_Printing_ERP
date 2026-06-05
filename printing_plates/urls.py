from django.urls import path

from . import views

app_name = 'printing_plates'

urlpatterns = [
    path('', views.PlateQueueView.as_view(), name='queue'),
    path('status-board/', views.PlateStatusBoardView.as_view(), name='status_board'),
    path('request/add/', views.PlateRequestCreateView.as_view(), name='request_add'),
    path('request/<int:pk>/', views.PlateRequestDetailView.as_view(), name='request_detail'),
]
