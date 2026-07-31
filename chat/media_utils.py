import io

from django.core.files.base import ContentFile

IMAGE_CONTENT_TYPES = {'image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/bmp'}
THUMBNAIL_MAX_SIZE = (320, 320)


def classify_file_type(content_type):
    if content_type in IMAGE_CONTENT_TYPES:
        return 'image'
    if content_type.startswith('audio/'):
        return 'audio'
    if content_type in {
        'application/pdf', 'application/msword', 'text/plain',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
    }:
        return 'doc'
    return 'other'


def build_thumbnail(uploaded_file, content_type):
    """Generate a small JPEG thumbnail for an image upload. Returns a Django
    ContentFile, or None if the file isn't an image or Pillow can't read it."""
    if content_type not in IMAGE_CONTENT_TYPES:
        return None

    from PIL import Image

    try:
        uploaded_file.seek(0)
        image = Image.open(uploaded_file)
        image = image.convert('RGB')
        image.thumbnail(THUMBNAIL_MAX_SIZE)
        buffer = io.BytesIO()
        image.save(buffer, format='JPEG', quality=80)
        buffer.seek(0)
        return ContentFile(buffer.read(), name='thumb.jpg')
    except Exception:
        return None
    finally:
        uploaded_file.seek(0)
