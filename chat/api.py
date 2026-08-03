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
from .models import Attachment, CallSession, ChatParticipant, ChatRoom, Message, MessageReaction
from .permissions import (
    CanCreateGroup,
    ChatAccessPermission,
    can_delete_any_message,
    can_manage_group,
)
from .realtime import broadcast_room_event, notify_users

MESSAGE_EDIT_WINDOW_MINUTES = getattr(settings, 'CHAT_MESSAGE_EDIT_WINDOW_MINUTES', 15)
MAX_ATTACHMENT_MB = getattr(settings, 'CHAT_MAX_ATTACHMENT_MB', 25)
BUZZ_COOLDOWN_SECONDS = 5


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


def _parse_and_set_mentions(room, message):
    """Server-side @name parsing: for every '@' in the message body, check
    whether a participant's username or full name (case-insensitively)
    appears right after it, and set Message.mentions accordingly.

    Matching against the actual candidate list — rather than a generic
    '@(\\w[\\w .]{0,40})'-style regex — is required because full names can
    contain spaces with no delimiter marking where the name ends and the
    rest of the sentence begins (e.g. "@John Smith please check" vs.
    "@Bob check this"); a regex permissive enough to capture multi-word
    names ends up swallowing trailing words too, so it stops matching real
    names at all. Longest-candidate-first plus a word-boundary check after
    the match keeps "@Bob" from matching inside "@Bobby"."""
    import re

    body = message.body or ''
    at_positions = [m.start() for m in re.finditer('@', body)]
    if not at_positions:
        message.mentions.clear()
        return []

    participants = room.participants.filter(left_at__isnull=True).select_related('user')
    candidates = []
    for participant in participants:
        user = participant.user
        for name in {user.username, (user.get_full_name() or '').strip()}:
            if name:
                candidates.append((name.lower(), user.id))
    candidates.sort(key=lambda c: -len(c[0]))

    matched_ids = set()
    for pos in at_positions:
        remainder = body[pos + 1:].lower()
        for name_lower, user_id in candidates:
            if not remainder.startswith(name_lower):
                continue
            end = pos + 1 + len(name_lower)
            next_char = body[end:end + 1]
            if next_char.isalnum():
                continue
            matched_ids.add(user_id)
            break

    matched_ids = list(matched_ids)
    message.mentions.set(matched_ids)
    return matched_ids


def _notify_new_message(room, message, sender, preview, mentioned_user_ids=None):
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
        'mentioned_user_ids': mentioned_user_ids or [],
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
            ChatRoom.objects.filter(
                participants__user=request.user, participants__left_at__isnull=True, is_archived=False,
            )
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

    def update_settings(self, request, pk=None):
        from .serializers import ChatRoomDetailSerializer

        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        if room.room_type != 'group':
            return Response({'detail': 'Only group rooms have settings.'}, status=status.HTTP_400_BAD_REQUEST)
        if not can_manage_group(request.user):
            return Response({'detail': 'You do not have permission to manage this group.'}, status=status.HTTP_403_FORBIDDEN)

        update_fields = []
        if 'name' in request.data:
            name = (request.data.get('name') or '').strip()
            if not name:
                return Response({'detail': 'Group name cannot be empty.'}, status=status.HTTP_400_BAD_REQUEST)
            room.name = name
            update_fields.append('name')
        if 'description' in request.data:
            room.description = (request.data.get('description') or '').strip()
            update_fields.append('description')
        if 'avatar' in request.FILES:
            room.avatar = request.FILES['avatar']
            update_fields.append('avatar')
        if update_fields:
            room.save(update_fields=update_fields)

        serializer = ChatRoomDetailSerializer(room, context={'request': request})
        broadcast_room_event(room.id, 'group_updated', {'room': serializer.data})
        return Response(serializer.data)

    def destroy(self, request, pk=None):
        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        if room.room_type != 'group':
            return Response({'detail': 'Only group rooms can be deleted.'}, status=status.HTTP_400_BAD_REQUEST)
        # Deliberately a higher trust tier than can_manage_group — deleting a
        # group is materially more destructive than renaming/describing it.
        if not request.user.is_superuser:
            return Response({'detail': 'Only a superuser can delete a group.'}, status=status.HTTP_403_FORBIDDEN)

        room.is_archived = True
        room.save(update_fields=['is_archived'])
        member_ids = list(room.participants.filter(left_at__isnull=True).values_list('user_id', flat=True))
        broadcast_room_event(room.id, 'group_deleted', {'room_id': room.id})
        notify_users(member_ids, 'group_deleted', {'room_id': room.id})
        return Response(status=status.HTTP_204_NO_CONTENT)

    def pin_message(self, request, pk=None):
        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        if room.room_type == 'group' and not can_manage_group(request.user):
            return Response({'detail': 'You do not have permission to pin messages in this group.'}, status=status.HTTP_403_FORBIDDEN)

        message_id = request.data.get('message_id')
        message = get_object_or_404(Message, pk=message_id, room=room, is_deleted=False)
        room.pinned_message = message
        room.save(update_fields=['pinned_message'])

        from .serializers import MessageSerializer
        payload = MessageSerializer(message, context={'request': request}).data
        broadcast_room_event(room.id, 'pin_updated', {'pinned_message': payload})
        return Response(payload)

    def unpin_message(self, request, pk=None):
        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        if room.room_type == 'group' and not can_manage_group(request.user):
            return Response({'detail': 'You do not have permission to unpin messages in this group.'}, status=status.HTTP_403_FORBIDDEN)

        room.pinned_message = None
        room.save(update_fields=['pinned_message'])
        broadcast_room_event(room.id, 'pin_updated', {'pinned_message': None})
        return Response(status=status.HTTP_204_NO_CONTENT)

    def buzz(self, request, pk=None):
        from django.core.cache import cache

        room = self.get_room_or_404(request, pk)
        if room is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)

        cooldown_key = f'chat:buzz_cooldown:{room.id}:{request.user.id}'
        if cache.get(cooldown_key):
            return Response({'detail': 'Please wait before buzzing again.'}, status=status.HTTP_429_TOO_MANY_REQUESTS)
        cache.set(cooldown_key, True, timeout=BUZZ_COOLDOWN_SECONDS)

        sender_name = request.user.get_full_name() or request.user.username
        other_user_ids = list(
            room.participants.filter(left_at__isnull=True).exclude(user=request.user).values_list('user_id', flat=True)
        )
        payload = {'room_id': room.id, 'from_user_id': request.user.id, 'from_display_name': sender_name}
        broadcast_room_event(room.id, 'buzz', payload)
        notify_users(other_user_ids, 'buzz', payload)
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
        # Tell other participants their message has now been seen up to here,
        # so senders can show a live "Seen" indicator without a manual refresh.
        broadcast_room_event(room.id, 'read_receipt_updated', {
            'user_id': request.user.id,
            'last_read_message_id': last_message.id if last_message else None,
        })
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

        forwarded_from_id = request.data.get('forwarded_from_message_id')
        source_message = None
        if forwarded_from_id:
            source_message = Message.objects.filter(pk=forwarded_from_id, is_deleted=False).select_related('room').first()
            if source_message is None or _active_participant(source_message.room, request.user) is None:
                return Response({'detail': 'That message is not available to forward.'}, status=status.HTTP_403_FORBIDDEN)

        body = (request.data.get('body') or (source_message.body if source_message else '')).strip()
        reply_to_id = request.data.get('reply_to')
        if not body and not (source_message and source_message.attachments.exists()):
            return Response({'detail': 'Message body cannot be empty.'}, status=status.HTTP_400_BAD_REQUEST)

        message = Message.objects.create(
            room=room, sender=request.user, body=body,
            message_type=source_message.message_type if source_message else 'text',
            reply_to_id=reply_to_id or None,
            forwarded_from=source_message,
        )
        if source_message:
            for att in source_message.attachments.all():
                # Re-reference the same underlying file — no re-upload round trip.
                Attachment.objects.create(
                    message=message, file=att.file.name, original_filename=att.original_filename,
                    content_type=att.content_type, file_type=att.file_type,
                    thumbnail=att.thumbnail.name if att.thumbnail else None,
                    size_bytes=att.size_bytes, uploaded_by=request.user,
                )
        room.touch_last_message(message.created_at)
        _mark_own_message_read(room, request.user, message.id)
        mentioned_ids = _parse_and_set_mentions(room, message) if room.room_type == 'group' else []

        serializer = MessageSerializer(message, context={'request': request})
        broadcast_room_event(room.id, 'message_created', {'message': serializer.data})
        _notify_new_message(room, message, request.user, preview=body, mentioned_user_ids=mentioned_ids)

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
        if room.room_type == 'group':
            _parse_and_set_mentions(room, message)

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


class MessageReactionView(APIView):
    """Toggle the current user's reaction on a message: picking the same
    emoji again removes it, picking a different one replaces it — one
    reaction per user per message, matching WhatsApp."""

    permission_classes = [IsAuthenticated, ChatAccessPermission]

    def post(self, request, room_id, message_id):
        room = get_object_or_404(ChatRoom, pk=room_id)
        if _active_participant(room, request.user) is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        message = get_object_or_404(Message, pk=message_id, room=room, is_deleted=False)

        emoji = (request.data.get('emoji') or '').strip()
        if not emoji:
            return Response({'detail': 'emoji is required.'}, status=status.HTTP_400_BAD_REQUEST)

        existing = MessageReaction.objects.filter(message=message, user=request.user).first()
        if existing and existing.emoji == emoji:
            existing.delete()
        elif existing:
            existing.emoji = emoji
            existing.save(update_fields=['emoji'])
        else:
            MessageReaction.objects.create(message=message, user=request.user, emoji=emoji)

        payload = {'message_id': message.id, 'reactions': self._aggregate(message)}
        broadcast_room_event(room.id, 'reaction_updated', payload)
        return Response(payload)

    def _aggregate(self, message):
        grouped = {}
        for r in message.reactions.all():
            grouped.setdefault(r.emoji, []).append(r.user_id)
        return [{'emoji': emoji, 'user_ids': user_ids} for emoji, user_ids in grouped.items()]


class MessageReadByView(APIView):
    """Who has read a given message, derived from ChatParticipant's
    per-room watermark (no per-message receipt table — see plan addendum
    for the trade-off: shows 'seen up to here', not an exact per-message
    read timestamp)."""

    permission_classes = [IsAuthenticated, ChatAccessPermission]

    def get(self, request, room_id, message_id):
        room = get_object_or_404(ChatRoom, pk=room_id)
        if _active_participant(room, request.user) is None:
            return Response({'detail': 'Not a member of this room.'}, status=status.HTTP_403_FORBIDDEN)
        message = get_object_or_404(Message, pk=message_id, room=room)

        readers = ChatParticipant.objects.filter(
            room=room, left_at__isnull=True, last_read_message_id__gte=message.id,
        ).exclude(user_id=message.sender_id).select_related('user')
        return Response([
            {'id': p.user.id, 'display_name': p.user.get_full_name() or p.user.username}
            for p in readers
        ])


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


class OnlineUsersView(APIView):
    """Snapshot of every connected user's status ('online' or 'away'), for a
    client's initial state on page load — presence deltas afterward arrive
    via PresenceConsumer's presence_changed WebSocket broadcasts. Anyone
    absent from the map is offline."""

    permission_classes = [IsAuthenticated, ChatAccessPermission]

    def get(self, request):
        from . import presence
        return Response({'statuses': presence.get_statuses()})


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
