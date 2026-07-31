from django.contrib import admin

from .models import Attachment, CallParticipant, CallSession, ChatParticipant, ChatRoom, Message


@admin.register(ChatRoom)
class ChatRoomAdmin(admin.ModelAdmin):
    list_display = ('id', 'room_type', 'name', 'created_by', 'created_at', 'is_archived', 'last_message_at')
    list_filter = ('room_type', 'is_archived')
    search_fields = ('name',)


@admin.register(ChatParticipant)
class ChatParticipantAdmin(admin.ModelAdmin):
    list_display = ('id', 'room', 'user', 'role', 'joined_at', 'left_at')
    list_filter = ('role',)


@admin.register(Message)
class MessageAdmin(admin.ModelAdmin):
    list_display = ('id', 'room', 'sender', 'message_type', 'created_at', 'is_deleted')
    list_filter = ('message_type', 'is_deleted')


@admin.register(Attachment)
class AttachmentAdmin(admin.ModelAdmin):
    list_display = ('id', 'message', 'original_filename', 'file_type', 'size_bytes', 'created_at')
    list_filter = ('file_type',)


@admin.register(CallSession)
class CallSessionAdmin(admin.ModelAdmin):
    list_display = ('id', 'room', 'call_type', 'status', 'started_at', 'ended_at')
    list_filter = ('call_type', 'status')


@admin.register(CallParticipant)
class CallParticipantAdmin(admin.ModelAdmin):
    list_display = ('id', 'call', 'user', 'joined_at', 'left_at')
