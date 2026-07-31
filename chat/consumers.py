import json

from channels.db import database_sync_to_async
from channels.generic.websocket import AsyncJsonWebsocketConsumer

from core.permissions import user_has_permission


class ChatAccessMixin:
    async def user_has_chat_access(self, user):
        if not user or not user.is_authenticated:
            return False
        return await database_sync_to_async(
            lambda: user.is_superuser or user_has_permission(user, 'nav.chat')
        )()


class ChatConsumer(ChatAccessMixin, AsyncJsonWebsocketConsumer):
    """One socket per open room. Handles ephemeral events (typing) and
    fans out `chat.event` messages sent by REST views via the channel layer.
    Message persistence itself happens over REST, not here."""

    async def connect(self):
        self.room_id = self.scope['url_route']['kwargs']['room_id']
        self.group_name = f'chat_room_{self.room_id}'
        user = self.scope.get('user')

        if not await self.user_has_chat_access(user):
            await self.close(code=4403)
            return

        is_participant = await database_sync_to_async(self._is_active_participant)(user)
        if not is_participant:
            await self.close(code=4403)
            return

        await self.channel_layer.group_add(self.group_name, self.channel_name)
        await self.accept()

    def _is_active_participant(self, user):
        from .models import ChatParticipant
        return ChatParticipant.objects.filter(room_id=self.room_id, user=user, left_at__isnull=True).exists()

    async def disconnect(self, close_code):
        if hasattr(self, 'group_name'):
            await self.channel_layer.group_discard(self.group_name, self.channel_name)

    async def receive_json(self, content, **kwargs):
        event = content.get('event')
        user = self.scope.get('user')
        if event == 'typing':
            await self.channel_layer.group_send(self.group_name, {
                'type': 'chat.event',
                'payload': {
                    'event': 'typing',
                    'user_id': user.id,
                    'username': user.username,
                    'is_typing': bool(content.get('is_typing')),
                },
            })

    async def chat_event(self, event):
        await self.send_json(event['payload'])


class PresenceConsumer(ChatAccessMixin, AsyncJsonWebsocketConsumer):
    """One long-lived socket per session: online/offline presence,
    cross-room unread-badge pushes, and incoming-call notifications."""

    async def connect(self):
        user = self.scope.get('user')
        if not await self.user_has_chat_access(user):
            await self.close(code=4403)
            return

        self.user_group = f'chat_user_{user.id}'
        await self.channel_layer.group_add(self.user_group, self.channel_name)
        await self.channel_layer.group_add('chat_presence', self.channel_name)
        await self.accept()
        await self.channel_layer.group_send('chat_presence', {
            'type': 'chat.event',
            'payload': {'event': 'presence_online', 'user_id': user.id},
        })

    async def disconnect(self, close_code):
        user = self.scope.get('user')
        if hasattr(self, 'user_group'):
            await self.channel_layer.group_discard(self.user_group, self.channel_name)
            await self.channel_layer.group_discard('chat_presence', self.channel_name)
        if user and getattr(user, 'is_authenticated', False):
            await self.channel_layer.group_send('chat_presence', {
                'type': 'chat.event',
                'payload': {'event': 'presence_offline', 'user_id': user.id},
            })

    async def chat_event(self, event):
        await self.send_json(event['payload'])


class CallSignalingConsumer(ChatAccessMixin, AsyncJsonWebsocketConsumer):
    """Pure relay for WebRTC offer/answer/ICE candidates between the
    participants of one room's call. Also records CallSession rows so
    call history stays populated via REST reads."""

    async def connect(self):
        self.room_id = self.scope['url_route']['kwargs']['room_id']
        self.group_name = f'call_room_{self.room_id}'
        user = self.scope.get('user')

        if not await self.user_has_chat_access(user):
            await self.close(code=4403)
            return

        is_participant = await database_sync_to_async(self._is_active_participant)(user)
        if not is_participant:
            await self.close(code=4403)
            return

        await self.channel_layer.group_add(self.group_name, self.channel_name)
        await self.accept()

    def _is_active_participant(self, user):
        from .models import ChatParticipant
        return ChatParticipant.objects.filter(room_id=self.room_id, user=user, left_at__isnull=True).exists()

    def _room_label(self, sender):
        from .models import ChatRoom
        room = ChatRoom.objects.get(pk=self.room_id)
        if room.room_type == 'group':
            return room.name or 'Group Chat'
        return sender.get_full_name() or sender.username

    async def disconnect(self, close_code):
        if hasattr(self, 'group_name'):
            await self.channel_layer.group_discard(self.group_name, self.channel_name)

    async def receive_json(self, content, **kwargs):
        event = content.get('event')
        user = self.scope.get('user')

        if event == 'call-invite':
            call_id = await database_sync_to_async(self._create_call_session)(user, content.get('call_type', 'audio'))
            content['call_id'] = call_id
            content['room_label'] = await database_sync_to_async(self._room_label)(user)
            content['sender_name'] = user.get_full_name() or user.username
            await self._notify_other_participants(user, 'incoming_call', content)
        elif event in {'call-decline', 'hangup'}:
            await database_sync_to_async(self._end_call_session)(content.get('call_id'), event)

        content.setdefault('from_user_id', user.id)
        await self.channel_layer.group_send(self.group_name, {'type': 'chat.event', 'payload': content})

    async def _notify_other_participants(self, sender, event_type, data):
        member_ids = await database_sync_to_async(self._other_participant_ids)(sender)
        for uid in member_ids:
            await self.channel_layer.group_send(f'chat_user_{uid}', {
                'type': 'chat.event',
                'payload': {'event': event_type, 'room_id': int(self.room_id), **data},
            })

    def _other_participant_ids(self, sender):
        from .models import ChatParticipant
        return list(
            ChatParticipant.objects.filter(room_id=self.room_id, left_at__isnull=True)
            .exclude(user=sender)
            .values_list('user_id', flat=True)
        )

    def _create_call_session(self, user, call_type):
        from .models import CallParticipant, CallSession
        call = CallSession.objects.create(room_id=self.room_id, initiated_by=user, call_type=call_type, status='ringing')
        CallParticipant.objects.create(call=call, user=user)
        return call.id

    def _end_call_session(self, call_id, event):
        from django.utils import timezone
        from .models import CallSession
        if not call_id:
            return
        status_map = {'hangup': 'ended', 'call-decline': 'declined'}
        CallSession.objects.filter(pk=call_id).update(
            status=status_map.get(event, 'ended'),
            ended_at=timezone.now(),
            end_reason='hangup' if event == 'hangup' else 'declined',
        )

    async def chat_event(self, event):
        await self.send_json(event['payload'])
