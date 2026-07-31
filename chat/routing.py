from django.urls import re_path

from . import consumers

websocket_urlpatterns = [
    re_path(r'^ws/chat/room/(?P<room_id>\d+)/$', consumers.ChatConsumer.as_asgi()),
    re_path(r'^ws/chat/presence/$', consumers.PresenceConsumer.as_asgi()),
    re_path(r'^ws/chat/call/(?P<room_id>\d+)/$', consumers.CallSignalingConsumer.as_asgi()),
]
