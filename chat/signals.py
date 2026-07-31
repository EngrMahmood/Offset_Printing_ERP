from django.db.models.signals import post_delete
from django.dispatch import receiver

from .models import Attachment


@receiver(post_delete, sender=Attachment)
def delete_attachment_files(sender, instance, **kwargs):
    """Remove the file/thumbnail from disk when an Attachment row is deleted,
    so cleanup_chat_media (and any manual deletes) don't leave orphans."""
    for field in (instance.file, instance.thumbnail):
        if field:
            field.storage.delete(field.name)
