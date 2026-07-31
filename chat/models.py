import uuid

from django.conf import settings
from django.db import models
from django.utils import timezone

User = settings.AUTH_USER_MODEL


def chat_attachment_path(instance, filename):
    room_id = instance.message.room_id
    now = timezone.now()
    safe_name = filename.replace('/', '_').replace('\\', '_')
    return f'chat_media/{room_id}/{now:%Y}/{now:%m}/{uuid.uuid4().hex}_{safe_name}'


def chat_thumbnail_path(instance, filename):
    room_id = instance.message.room_id
    now = timezone.now()
    return f'chat_media/{room_id}/{now:%Y}/{now:%m}/thumbs/{uuid.uuid4().hex}_{filename}'


class ChatRoom(models.Model):
    ROOM_TYPE_CHOICES = [
        ('dm', 'Direct Message'),
        ('group', 'Group'),
    ]

    room_type = models.CharField(max_length=10, choices=ROOM_TYPE_CHOICES)
    name = models.CharField(max_length=150, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='chat_rooms_created')
    created_at = models.DateTimeField(auto_now_add=True)
    is_archived = models.BooleanField(default=False)
    last_message_at = models.DateTimeField(null=True, blank=True, db_index=True)

    class Meta:
        ordering = ['-last_message_at', '-created_at']

    def __str__(self):
        if self.room_type == 'group':
            return self.name or f'Group #{self.pk}'
        return f'DM #{self.pk}'

    def display_name_for(self, user):
        if self.room_type == 'group':
            return self.name or 'Group Chat'
        other = self.participants.exclude(user=user).select_related('user').first()
        return other.user.get_full_name() or other.user.username if other else 'Direct Message'

    @classmethod
    def get_or_create_dm(cls, user_a, user_b):
        existing = (
            cls.objects.filter(room_type='dm', participants__user=user_a)
            .filter(participants__user=user_b)
            .first()
        )
        if existing:
            return existing, False

        room = cls.objects.create(room_type='dm', created_by=user_a)
        ChatParticipant.objects.create(room=room, user=user_a, role='member')
        ChatParticipant.objects.create(room=room, user=user_b, role='member')
        return room, True

    def touch_last_message(self, when=None):
        self.last_message_at = when or timezone.now()
        self.save(update_fields=['last_message_at'])


class ChatParticipant(models.Model):
    ROLE_CHOICES = [
        ('member', 'Member'),
        ('admin', 'Admin'),
    ]

    room = models.ForeignKey(ChatRoom, on_delete=models.CASCADE, related_name='participants')
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='chat_participations')
    joined_at = models.DateTimeField(auto_now_add=True)
    left_at = models.DateTimeField(null=True, blank=True)
    role = models.CharField(max_length=10, choices=ROLE_CHOICES, default='member')
    last_read_message_id = models.PositiveIntegerField(null=True, blank=True)
    is_muted = models.BooleanField(default=False)

    class Meta:
        unique_together = ('room', 'user')
        indexes = [
            models.Index(fields=['user', 'room']),
            models.Index(fields=['room', 'left_at']),
        ]

    def __str__(self):
        return f'{self.user_id} in room {self.room_id}'

    @property
    def is_active(self):
        return self.left_at is None


class Message(models.Model):
    MESSAGE_TYPE_CHOICES = [
        ('text', 'Text'),
        ('file', 'File'),
        ('system', 'System'),
        ('call', 'Call'),
    ]

    room = models.ForeignKey(ChatRoom, on_delete=models.CASCADE, related_name='messages')
    sender = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='chat_messages')
    body = models.TextField(blank=True)
    message_type = models.CharField(max_length=10, choices=MESSAGE_TYPE_CHOICES, default='text')
    reply_to = models.ForeignKey('self', on_delete=models.SET_NULL, null=True, blank=True, related_name='replies')
    created_at = models.DateTimeField(auto_now_add=True, db_index=True)
    edited_at = models.DateTimeField(null=True, blank=True)
    is_deleted = models.BooleanField(default=False)
    deleted_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ['created_at']
        indexes = [
            models.Index(fields=['room', 'created_at']),
            models.Index(fields=['room', '-id']),
        ]

    def __str__(self):
        return f'Message #{self.pk} in room {self.room_id}'

    def soft_delete(self):
        self.is_deleted = True
        self.deleted_at = timezone.now()
        self.body = ''
        self.save(update_fields=['is_deleted', 'deleted_at', 'body'])


class Attachment(models.Model):
    FILE_TYPE_CHOICES = [
        ('image', 'Image'),
        ('doc', 'Document'),
        ('other', 'Other'),
    ]

    message = models.ForeignKey(Message, on_delete=models.CASCADE, related_name='attachments')
    file = models.FileField(upload_to=chat_attachment_path)
    original_filename = models.CharField(max_length=255)
    content_type = models.CharField(max_length=120, blank=True)
    file_type = models.CharField(max_length=10, choices=FILE_TYPE_CHOICES, default='other')
    thumbnail = models.ImageField(upload_to=chat_thumbnail_path, null=True, blank=True)
    size_bytes = models.PositiveIntegerField(default=0)
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='chat_attachments')
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.original_filename


class CallSession(models.Model):
    CALL_TYPE_CHOICES = [
        ('audio', 'Audio'),
        ('video', 'Video'),
    ]
    STATUS_CHOICES = [
        ('ringing', 'Ringing'),
        ('active', 'Active'),
        ('ended', 'Ended'),
        ('missed', 'Missed'),
        ('declined', 'Declined'),
    ]
    END_REASON_CHOICES = [
        ('hangup', 'Hangup'),
        ('timeout', 'Timeout'),
        ('declined', 'Declined'),
        ('error', 'Error'),
    ]

    room = models.ForeignKey(ChatRoom, on_delete=models.CASCADE, related_name='calls')
    initiated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='chat_calls_initiated')
    call_type = models.CharField(max_length=10, choices=CALL_TYPE_CHOICES, default='audio')
    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default='ringing')
    started_at = models.DateTimeField(auto_now_add=True)
    answered_at = models.DateTimeField(null=True, blank=True)
    ended_at = models.DateTimeField(null=True, blank=True)
    end_reason = models.CharField(max_length=10, choices=END_REASON_CHOICES, blank=True)

    class Meta:
        ordering = ['-started_at']

    def __str__(self):
        return f'{self.call_type} call in room {self.room_id} ({self.status})'


class CallParticipant(models.Model):
    call = models.ForeignKey(CallSession, on_delete=models.CASCADE, related_name='call_participants')
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='chat_call_participations')
    joined_at = models.DateTimeField(null=True, blank=True)
    left_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        unique_together = ('call', 'user')

    def __str__(self):
        return f'{self.user_id} in call {self.call_id}'
