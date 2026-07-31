from django.conf import settings
from django.contrib.auth.models import User
from django.db import transaction
from django.shortcuts import get_object_or_404
from django.utils import timezone
from rest_framework import status, viewsets
from rest_framework.pagination import CursorPagination
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from rest_framework.views import APIView

from .media_utils import build_thumbnail, classify_file_type
from .models import Attachment, CallSession, ChatParticipant, ChatRoom, Message
from .permissions import (
    CanCreateGroup,
    ChatAccessPermission,
    can_delete_any_message,
    can_manage_group,
)
from .realtime import broadcast_room_event, notify_users

MESSAGE_EDIT_WINDOW_MINUTES = getattr(settings, 'CHAT_MESSAGE_EDIT_WINDOW_MINUTES', 15)
MAX_ATTACHMENT_MB = getattr(settings, 'CHAT_MAX_ATTACHMENT_MB', 25)


def _active_participant(room, user):
    return ChatParticipant.objects.filter(room=room, user=user, left_at__isnull=True).first()


def _require_participant(room, user):
    participant = _active_participant(room, user)
    if participant is None:
        return None
    return participant


def _mark_own_message_read(room, user, message_id):
    """A sender has implicitly read their own message — keep their unread
    count from counting messages they wrote themselves."""
    ChatParticipant.objects.filter(room=room, user=user, left_at__isnull=True).update(
        last_read_message_id=message_id,
    )


def _notify_new_message(room, message, sender, preview):
    """Push a toast-worthy event (site-wide, not just to open chat tabs) to
    every other active participant. Room label mirrors ChatRoom.display_name_for:
    the sender's name for a DM, the group name for a group."""
    if room.room_type == 'group':
        room_label = room.name or 'Group Chat'
    else:
        room_label = sender.get_full_name() or sender.username

    other_user_ids = list(
        room.participants.filter(left_at__isnull=True).exclude(user=sender).values_list('user_id', flat=True)
    )
    notify_users(other_user_ids, 'new_message', {
        'room_id': room.id,
        'room_label': room_label,
        'sender_name': sender.get_full_name() or sender.username,
        'message_id': message.id,
        'preview': preview[:140],
    })


class MessageCursorPagination(CursorPagination):
    page_size = 30
    ordering = '-id'


class RoomViewSet(viewsets.ViewSet):
    permission_classes = [IsAuthenticated, ChatAccessPermission]

    def get_room_or_404(self, request, pk, require_participant=True):
        room = get_object_or_404(ChatRoom, pk=pk)
        if require_participant and _active_participant(room, request.user) is None:
            return None
        return room

    def list(self, request):
        from .serializers import ChatRoomListSerializer

        rooms = (
            ChatRoom.objects.filter(participants__user=request.user, participants__left_at__isnull=True)
            .distinct()
            .order_by('-last_message_at', '-created_at')
        )
        serializer = ChatRoomListSerializer(rooms, many=True, context={'request': request})
        return Response(serializer.data)

    def create(self, request):
        from .serializers import ChatRoomDetailSerializer

        room_type = request.data.get('room_type', 'dm')
        if room_type == 'dm':
            other_user_id = request.data.get('user_id')
            if not other_user_id:
                return Response({'detail': 'user_id is required for a DM.'}, status=status.HTTP_400_BAD_REQUEST)
            other_user = get_object_or_404(User, pk=other_user_id)
            if other_user.id == request.user.id:
                return Response({'detail': 'Cannot start a DM with yourself.'}, status=status.HTTP_400_BAD_REQUEST)
            room, _created = ChatRoom.get_or_create_dm(request.user, other_user)
        elif room_type == 'group':
            if not (request.user.is_superuser or CanCreateGroup().has_permission(request, self)):
                return Response({'detail': 'You do not have permission to create group chats.'}, status=status.HTTP_403_FORBIDDEN)
            name = (request.data.get('name') or '').strip()
            if not name:
                return Response({'detail': 'Group name is required.'}, status=status.HTTP_400_BAD_REQUEST)
            member_ids = request.data.get('member_ids') or []
            with transaction.atomic():
                room = ChatRoom.objects.create(room_type='group', name=name, created_by=request.user)
                ChatParticipant.objects.create(room=room, user=request.user, role='admin')
                for uid in member_ids:
                    if str(uid) == str(request.user.id):
                        continue
                    member = User.objects.filter(pk=uid).first()
                    if member:
                        ChatParticipant.objects.get_or_create(room=room, user=member, defaults={'role': 'member'})
        else:
            return Response({'detail': 'room_type must be dm or group.'}, status=status.HTTP_400_BAD_REQUEST)

        serializer = ChatRoomDetailSerializer(room, context={'request': request})
        return Response(serializer.data, status=status.HTTP_201_CREATED)

    def retrieve(self, request, pk=None):
        from .serializers import ChatRoomDetailSerializer

        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        serializer = ChatRoomDetailSerializer(room, context={'request': request})
        return Response(serializer.data)

    def add_participant(self, request, pk=None):
        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        if room.room_type != 'group':
            return Response({'detail': 'Only group rooms support membership changes.'}, status=status.HTTP_400_BAD_REQUEST)
        if not can_manage_group(request.user):
            return Response({'detail': 'You do not have permission to manage group members.'}, status=status.HTTP_403_FORBIDDEN)

        user_id = request.data.get('user_id')
        member = get_object_or_404(User, pk=user_id)
        participant, created = ChatParticipant.objects.get_or_create(
            room=room, user=member, defaults={'role': 'member'},
        )
        if not created and participant.left_at is not None:
            participant.left_at = None
            participant.save(update_fields=['left_at'])

        broadcast_room_event(room.id, 'participant_added', {'user_id': member.id})
        return Response(status=status.HTTP_201_CREATED)

    def remove_participant(self, request, pk=None, user_id=None):
        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        if not can_manage_group(request.user):
            return Response({'detail': 'You do not have permission to manage group members.'}, status=status.HTTP_403_FORBIDDEN)

        participant = get_object_or_404(ChatParticipant, room=room, user_id=user_id, left_at__isnull=True)
        participant.left_at = timezone.now()
        participant.save(update_fields=['left_at'])
        broadcast_room_event(room.id, 'participant_removed', {'user_id': int(user_id)})
        return Response(status=status.HTTP_204_NO_CONTENT)

    def leave(self, request, pk=None):
        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        participant = _active_participant(room, request.user)
        participant.left_at = timezone.now()
        participant.save(update_fields=['left_at'])
        broadcast_room_event(room.id, 'participant_removed', {'user_id': request.user.id})
        return Response(status=status.HTTP_204_NO_CONTENT)

    def mark_read(self, request, pk=None):
        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        participant = _active_participant(room, request.user)
        last_message = room.messages.order_by('-id').first()
        if last_message:
            participant.last_read_message_id = last_message.id
            participant.save(update_fields=['last_read_message_id'])
        notify_users([request.user.id], 'unread_count_changed', {'room_id': room.id, 'unread_count': 0})
        return Response(status=status.HTTP_200_OK)


class MessageListCreateView(APIView):
    permission_classes = [IsAuthenticated, ChatAccessPermission]
    pagination_class = MessageCursorPagination

    def _get_room(self, request, room_id):
        room = get_object_or_404(ChatRoom, pk=room_id)
        if _active_participant(room, request.user) is None:
            return None
        return room

    def get(self, request, room_id):
        from .serializers import MessageSerializer

        room = self._get_room(request, room_id)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)

        queryset = room.messages.select_related('sender').prefetch_related('attachments').order_by('-id')
        paginator = self.pagination_class()
        page = paginator.paginate_queryset(queryset, request, view=self)
        serializer = MessageSerializer(page, many=True, context={'request': request})
        return paginator.get_paginated_response(serializer.data)

    def post(self, request, room_id):
        from .serializers import MessageSerializer

        room = self._get_room(request, room_id)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)

        body = (request.data.get('body') or '').strip()
        reply_to_id = request.data.get('reply_to')
        if not body:
            return Response({'detail': 'Message body cannot be empty.'}, status=status.HTTP_400_BAD_REQUEST)

        message = Message.objects.create(
            room=room, sender=request.user, body=body, message_type='text',
            reply_to_id=reply_to_id or None,
        )
        room.touch_last_message(message.created_at)
        _mark_own_message_read(room, request.user, message.id)

        serializer = MessageSerializer(message, context={'request': request})
        broadcast_room_event(room.id, 'message_created', {'message': serializer.data})
        _notify_new_message(room, message, request.user, preview=body)

        return Response(serializer.data, status=status.HTTP_201_CREATED)


class MessageDetailView(APIView):
    permission_classes = [IsAuthenticated, ChatAccessPermission]

    def _get_message(self, request, room_id, message_id):
        room = get_object_or_404(ChatRoom, pk=room_id)
        if _active_participant(room, request.user) is None:
            return None, None
        message = get_object_or_404(Message, pk=message_id, room=room)
        return room, message

    def patch(self, request, room_id, message_id):
        from .serializers import MessageSerializer

        room, message = self._get_message(request, room_id, message_id)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        if message.sender_id != request.user.id:
            return Response({'detail': 'You can only edit your own messages.'}, status=status.HTTP_403_FORBIDDEN)
        if (timezone.now() - message.created_at).total_seconds() > MESSAGE_EDIT_WINDOW_MINUTES * 60:
            return Response({'detail': 'Edit window has expired.'}, status=status.HTTP_400_BAD_REQUEST)

        body = (request.data.get('body') or '').strip()
        if not body:
            return Response({'detail': 'Message body cannot be empty.'}, status=status.HTTP_400_BAD_REQUEST)

        message.body = body
        message.edited_at = timezone.now()
        message.save(update_fields=['body', 'edited_at'])

        serializer = MessageSerializer(message, context={'request': request})
        broadcast_room_event(room.id, 'message_edited', {'message': serializer.data})
        return Response(serializer.data)

    def delete(self, request, room_id, message_id):
        room, message = self._get_message(request, room_id, message_id)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)

        is_own = message.sender_id == request.user.id
        within_window = (timezone.now() - message.created_at).total_seconds() <= MESSAGE_EDIT_WINDOW_MINUTES * 60
        if not ((is_own and within_window) or can_delete_any_message(request.user)):
            return Response({'detail': 'You cannot delete this message.'}, status=status.HTTP_403_FORBIDDEN)

        message.soft_delete()
        broadcast_room_event(room.id, 'message_deleted', {'message_id': message.id})
        return Response(status=status.HTTP_204_NO_CONTENT)


class AttachmentUploadView(APIView):
    permission_classes = [IsAuthenticated, ChatAccessPermission]

    def post(self, request, room_id):
        from .serializers import MessageSerializer

        room = get_object_or_404(ChatRoom, pk=room_id)
        if _active_participant(room, request.user) is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)

        uploaded_file = request.FILES.get('file')
        if uploaded_file is None:
            return Response({'detail': 'file is required.'}, status=status.HTTP_400_BAD_REQUEST)

        max_bytes = MAX_ATTACHMENT_MB * 1024 * 1024
        if uploaded_file.size > max_bytes:
            return Response(
                {'detail': f'File exceeds the {MAX_ATTACHMENT_MB}MB limit.'},
                status=status.HTTP_400_BAD_REQUEST,
            )

        content_type = uploaded_file.content_type or ''
        file_type = classify_file_type(content_type)
        caption = (request.data.get('body') or '').strip()

        with transaction.atomic():
            message = Message.objects.create(
                room=room, sender=request.user, body=caption, message_type='file',
            )
            attachment = Attachment.objects.create(
                message=message,
                file=uploaded_file,
                original_filename=uploaded_file.name,
                content_type=content_type,
                file_type=file_type,
                size_bytes=uploaded_file.size,
                uploaded_by=request.user,
            )
            thumb = build_thumbnail(uploaded_file, content_type)
            if thumb is not None:
                attachment.thumbnail.save(thumb.name, thumb, save=True)

        room.touch_last_message(message.created_at)
        _mark_own_message_read(room, request.user, message.id)

        serializer = MessageSerializer(message, context={'request': request})
        broadcast_room_event(room.id, 'message_created', {'message': serializer.data})
        _notify_new_message(room, message, request.user, preview=caption or ('📎 ' + attachment.original_filename))

        return Response(serializer.data, status=status.HTTP_201_CREATED)


class CallSessionListView(APIView):
    permission_classes = [IsAuthenticated, ChatAccessPermission]

    def get(self, request, room_id):
        from .serializers import CallSessionSerializer

        room = get_object_or_404(ChatRoom, pk=room_id)
        if _active_participant(room, request.user) is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)

        calls = room.calls.select_related('initiated_by').prefetch_related('call_participants__user')[:50]
        serializer = CallSessionSerializer(calls, many=True, context={'request': request})
        return Response(serializer.data)


class IceConfigView(APIView):
    permission_classes = [IsAuthenticated, ChatAccessPermission]

    def get(self, request):
        return Response({'ice_servers': getattr(settings, 'CHAT_ICE_SERVERS', [])})


class ChattableUserListView(APIView):
    """Users the current user can start a DM/group with — anyone else with chat access."""

    permission_classes = [IsAuthenticated, ChatAccessPermission]

    def get(self, request):
        from core.permissions import user_has_permission

        users = User.objects.filter(is_active=True).exclude(pk=request.user.id).order_by('username')
        data = [
            {'id': u.id, 'username': u.username, 'display_name': u.get_full_name() or u.username}
            for u in users
            if u.is_superuser or user_has_permission(u, 'nav.chat')
        ]
        return Response(data)
