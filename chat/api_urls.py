from django.urls import path

from . import api

urlpatterns = [
    path('rooms/', api.RoomViewSet.as_view({'get': 'list', 'post': 'create'}), name='room-list'),
    path('rooms/<int:pk>/', api.RoomViewSet.as_view({'get': 'retrieve'}), name='room-detail'),
    path('rooms/<int:pk>/participants/', api.RoomViewSet.as_view({'post': 'add_participant'}), name='room-add-participant'),
    path('rooms/<int:pk>/participants/<int:user_id>/', api.RoomViewSet.as_view({'delete': 'remove_participant'}), name='room-remove-participant'),
    path('rooms/<int:pk>/leave/', api.RoomViewSet.as_view({'post': 'leave'}), name='room-leave'),
    path('rooms/<int:pk>/read/', api.RoomViewSet.as_view({'post': 'mark_read'}), name='room-mark-read'),

    path('rooms/<int:room_id>/messages/', api.MessageListCreateView.as_view(), name='message-list-create'),
    path('rooms/<int:room_id>/messages/<int:message_id>/', api.MessageDetailView.as_view(), name='message-detail'),
    path('rooms/<int:room_id>/attachments/', api.AttachmentUploadView.as_view(), name='attachment-upload'),
    path('rooms/<int:room_id>/calls/', api.CallSessionListView.as_view(), name='call-list'),

    path('ice-config/', api.IceConfigView.as_view(), name='ice-config'),
    path('users/', api.ChattableUserListView.as_view(), name='chattable-users'),
]
