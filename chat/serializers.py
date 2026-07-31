from django.contrib.auth.models import User
from rest_framework import serializers

from .models import Attachment, CallParticipant, CallSession, ChatParticipant, ChatRoom, Message


class UserBriefSerializer(serializers.ModelSerializer):
    display_name = serializers.SerializerMethodField()

    class Meta:
        model = User
        fields = ('id', 'username', 'display_name')

    def get_display_name(self, obj):
        return obj.get_full_name() or obj.username


class ChatParticipantSerializer(serializers.ModelSerializer):
    user = UserBriefSerializer(read_only=True)

    class Meta:
        model = ChatParticipant
        fields = ('id', 'user', 'role', 'joined_at', 'left_at', 'is_muted', 'last_read_message_id')


class AttachmentSerializer(serializers.ModelSerializer):
    url = serializers.SerializerMethodField()
    thumbnail_url = serializers.SerializerMethodField()

    class Meta:
        model = Attachment
        fields = (
            'id', 'url', 'thumbnail_url', 'original_filename', 'content_type',
            'file_type', 'size_bytes', 'created_at',
        )

    def get_url(self, obj):
        request = self.context.get('request')
        return request.build_absolute_uri(obj.file.url) if request and obj.file else None

    def get_thumbnail_url(self, obj):
        request = self.context.get('request')
        if obj.thumbnail and request:
            return request.build_absolute_uri(obj.thumbnail.url)
        return None


class MessageSerializer(serializers.ModelSerializer):
    sender = UserBriefSerializer(read_only=True)
    attachments = AttachmentSerializer(many=True, read_only=True)
    reactions = serializers.SerializerMethodField()
    mentions = serializers.SerializerMethodField()
    forwarded_from = serializers.SerializerMethodField()

    class Meta:
        model = Message
        fields = (
            'id', 'room', 'sender', 'body', 'message_type', 'reply_to',
            'created_at', 'edited_at', 'is_deleted', 'attachments',
            'reactions', 'mentions', 'forwarded_from',
        )
        read_only_fields = ('room', 'sender', 'message_type', 'created_at', 'edited_at', 'is_deleted')

    def get_reactions(self, obj):
        grouped = {}
        for r in obj.reactions.all():
            grouped.setdefault(r.emoji, []).append(r.user_id)
        return [{'emoji': emoji, 'user_ids': user_ids} for emoji, user_ids in grouped.items()]

    def get_mentions(self, obj):
        return list(obj.mentions.values_list('id', flat=True))

    def get_forwarded_from(self, obj):
        return bool(obj.forwarded_from_id)

    def validate_body(self, value):
        if not value.strip() and not self.initial_data.get('_has_attachment'):
            raise serializers.ValidationError('Message body cannot be empty.')
        return value


class ChatRoomListSerializer(serializers.ModelSerializer):
    display_name = serializers.SerializerMethodField()
    unread_count = serializers.SerializerMethodField()
    last_message = serializers.SerializerMethodField()
    participant_count = serializers.SerializerMethodField()
    other_user_id = serializers.SerializerMethodField()
    avatar_url = serializers.SerializerMethodField()

    class Meta:
        model = ChatRoom
        fields = (
            'id', 'room_type', 'name', 'description', 'avatar_url', 'display_name', 'created_at', 'is_archived',
            'last_message_at', 'unread_count', 'last_message', 'participant_count',
            'other_user_id',
        )

    def get_avatar_url(self, obj):
        request = self.context.get('request')
        if obj.avatar and request:
            return request.build_absolute_uri(obj.avatar.url)
        return None

    def get_display_name(self, obj):
        user = self.context['request'].user
        return obj.display_name_for(user)

    def get_other_user_id(self, obj):
        # Used client-side for DM online/offline indicators — meaningless for groups.
        if obj.room_type != 'dm':
            return None
        user = self.context['request'].user
        other = obj.participants.filter(left_at__isnull=True).exclude(user=user).first()
        return other.user_id if other else None

    def get_participant_count(self, obj):
        return obj.participants.filter(left_at__isnull=True).count()

    def get_unread_count(self, obj):
        user = self.context['request'].user
        participant = obj.participants.filter(user=user).first()
        if not participant:
            return 0
        qs = obj.messages.filter(is_deleted=False)
        if participant.last_read_message_id:
            qs = qs.filter(id__gt=participant.last_read_message_id)
        return qs.count()

    def get_last_message(self, obj):
        last = obj.messages.filter(is_deleted=False).order_by('-id').first()
        if not last:
            return None
        return {
            'id': last.id,
            'body': last.body[:140],
            'sender': last.sender.username if last.sender else None,
            'message_type': last.message_type,
            # .isoformat(), not the raw datetime: this dict is built by hand
            # (not a nested Field), so it skips DRF's JSON-response datetime
            # handling — and when this serializer's .data is broadcast over
            # the channel layer (group_updated), msgpack can't serialize a
            # raw datetime.datetime at all, which crashed that broadcast.
            'created_at': last.created_at.isoformat() if last.created_at else None,
        }


class ChatRoomDetailSerializer(ChatRoomListSerializer):
    participants = ChatParticipantSerializer(many=True, read_only=True)
    pinned_message = serializers.SerializerMethodField()

    class Meta(ChatRoomListSerializer.Meta):
        fields = ChatRoomListSerializer.Meta.fields + ('participants', 'pinned_message')

    def get_pinned_message(self, obj):
        if not obj.pinned_message_id or obj.pinned_message.is_deleted:
            return None
        return MessageSerializer(obj.pinned_message, context=self.context).data


class CallParticipantSerializer(serializers.ModelSerializer):
    user = UserBriefSerializer(read_only=True)

    class Meta:
        model = CallParticipant
        fields = ('id', 'user', 'joined_at', 'left_at')


class CallSessionSerializer(serializers.ModelSerializer):
    initiated_by = UserBriefSerializer(read_only=True)
    call_participants = CallParticipantSerializer(many=True, read_only=True)

    class Meta:
        model = CallSession
        fields = (
            'id', 'room', 'call_type', 'status', 'initiated_by', 'started_at',
            'answered_at', 'ended_at', 'end_reason', 'call_participants',
        )
