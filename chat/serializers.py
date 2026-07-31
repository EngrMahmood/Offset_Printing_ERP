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

    class Meta:
        model = Message
        fields = (
            'id', 'room', 'sender', 'body', 'message_type', 'reply_to',
            'created_at', 'edited_at', 'is_deleted', 'attachments',
        )
        read_only_fields = ('room', 'sender', 'message_type', 'created_at', 'edited_at', 'is_deleted')

    def validate_body(self, value):
        if not value.strip() and not self.initial_data.get('_has_attachment'):
            raise serializers.ValidationError('Message body cannot be empty.')
        return value


class ChatRoomListSerializer(serializers.ModelSerializer):
    display_name = serializers.SerializerMethodField()
    unread_count = serializers.SerializerMethodField()
    last_message = serializers.SerializerMethodField()
    participant_count = serializers.SerializerMethodField()

    class Meta:
        model = ChatRoom
        fields = (
            'id', 'room_type', 'name', 'display_name', 'created_at', 'is_archived',
            'last_message_at', 'unread_count', 'last_message', 'participant_count',
        )

    def get_display_name(self, obj):
        user = self.context['request'].user
        return obj.display_name_for(user)

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
            'created_at': last.created_at,
        }


class ChatRoomDetailSerializer(ChatRoomListSerializer):
    participants = ChatParticipantSerializer(many=True, read_only=True)

    class Meta(ChatRoomListSerializer.Meta):
        fields = ChatRoomListSerializer.Meta.fields + ('participants',)


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
