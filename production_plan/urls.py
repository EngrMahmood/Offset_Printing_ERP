from django.urls import path

from . import views

app_name = 'production_plan'

urlpatterns = [
    path('', views.board, name='board'),
    path('settings/', views.settings_page, name='settings'),
    path('jobs/search/', views.job_search, name='job_search'),
    path('<int:plan_id>/approve/', views.approve, name='approve'),
    path('<int:plan_id>/recalculate/', views.recalculate, name='recalculate'),
    path('<int:plan_id>/items/add/', views.item_add, name='item_add'),
    path('<int:plan_id>/items/<int:item_id>/remove/', views.item_remove, name='item_remove'),
    path('<int:plan_id>/items/<int:item_id>/move/', views.item_move, name='item_move'),
    path('<int:plan_id>/items/<int:item_id>/pin/', views.item_pin, name='item_pin'),
    path('<int:plan_id>/priority/', views.set_priority, name='set_priority'),
    path('<int:plan_id>/blocks/add/', views.block_add, name='block_add'),
    path('<int:plan_id>/blocks/<int:block_id>/remove/', views.block_remove, name='block_remove'),
    path('<int:plan_id>/export/<str:kind>/', views.export, name='export'),
    path('<int:plan_id>/share/text/', views.share_text, name='share_text'),
    path('<int:plan_id>/share/image/', views.share_image, name='share_image'),
    path('<int:plan_id>/email/', views.email_now, name='email_now'),
]
