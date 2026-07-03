"""In-app notification helpers (navbar bell + side toasts)."""

from __future__ import annotations

from django.contrib.auth import get_user_model
from django.urls import NoReverseMatch, reverse

from core.models import Notification

User = get_user_model()


def users_with_roles(*roles):
    role_set = {str(role or '').strip().lower() for role in roles if role}
    if not role_set:
        return User.objects.none()
    return User.objects.filter(
        is_active=True,
        profile__role__in=role_set,
    ).distinct()


def notify_users(
    users,
    *,
    event_type,
    title,
    message='',
    link='',
    entity_type='',
    entity_id=None,
    actor=None,
    exclude_actor=True,
):
    """Create one Notification row per recipient. Returns created count."""
    event_type = (event_type or '').strip()[:80]
    title = (title or '').strip()[:200]
    if not event_type or not title:
        return 0

    recipient_ids = set()
    for user in users:
        if user is None:
            continue
        user_id = getattr(user, 'pk', None) or user
        if not user_id:
            continue
        if exclude_actor and actor is not None and user_id == getattr(actor, 'pk', actor):
            continue
        recipient_ids.add(int(user_id))

    if not recipient_ids:
        return 0

    rows = [
        Notification(
            user_id=user_id,
            event_type=event_type,
            title=title,
            message=(message or '').strip(),
            link=(link or '').strip()[:500],
            entity_type=(entity_type or '').strip()[:60],
            entity_id=entity_id,
            created_by=actor if getattr(actor, 'pk', None) else None,
        )
        for user_id in recipient_ids
    ]
    Notification.objects.bulk_create(rows)
    return len(rows)


def notify_roles(
    roles,
    *,
    event_type,
    title,
    message='',
    link='',
    entity_type='',
    entity_id=None,
    actor=None,
    exclude_actor=True,
):
    return notify_users(
        users_with_roles(*roles),
        event_type=event_type,
        title=title,
        message=message,
        link=link,
        entity_type=entity_type,
        entity_id=entity_id,
        actor=actor,
        exclude_actor=exclude_actor,
    )


def _safe_reverse(name, *args, **kwargs):
    try:
        return reverse(name, args=args, kwargs=kwargs)
    except NoReverseMatch:
        return ''


def notify_plate_replacement(plate_request, actor=None):
    jc = plate_request.jc_number or 'job'
    link = _safe_reverse('printing_plates:request_detail', plate_request.pk)
    return notify_roles(
        ('graphics_designer', 'admin', 'manager'),
        event_type='plate.replacement_requested',
        title=f'Plate replacement: {jc}',
        message=(plate_request.damaged_colors or plate_request.remarks or 'Replacement plates requested.')[:300],
        link=link,
        entity_type='plate_request',
        entity_id=plate_request.pk,
        actor=actor,
    )


def notify_plate_cancelled(plate_request, actor=None):
    jc = plate_request.jc_number or 'job'
    link = _safe_reverse('printing_plates:request_detail', plate_request.pk)
    return notify_roles(
        ('planner', 'admin', 'manager'),
        event_type='plate.cancelled',
        title=f'Plate request cancelled: {jc}',
        message=(plate_request.progress or plate_request.remarks or 'Plates marked not required.')[:300],
        link=link,
        entity_type='plate_request',
        entity_id=plate_request.pk,
        actor=actor,
    )


def notify_sku_pending_review(recipe, actor=None):
    link = _safe_reverse('planning:sku_recipe_edit', recipe.pk)
    return notify_roles(
        ('qc', 'admin', 'manager'),
        event_type='sku.pending_review',
        title=f'SKU ready for review: {recipe.sku}',
        message=(recipe.job_name or 'Submitted for QC review.')[:300],
        link=link,
        entity_type='sku_recipe',
        entity_id=recipe.pk,
        actor=actor,
    )


def notify_sku_sent_back(recipe, actor=None):
    link = _safe_reverse('planning:sku_recipe_edit', recipe.pk)
    recipients = list(users_with_roles('planner', 'graphics_designer', 'admin', 'manager'))
    if recipe.created_by_id:
        recipients.append(recipe.created_by)
    return notify_users(
        recipients,
        event_type='sku.sent_back',
        title=f'SKU sent back: {recipe.sku}',
        message=(recipe.rejection_comment or 'Returned to draft for correction.')[:300],
        link=link,
        entity_type='sku_recipe',
        entity_id=recipe.pk,
        actor=actor,
    )


def serialize_notification(item):
    return {
        'id': item.pk,
        'event_type': item.event_type,
        'title': item.title,
        'message': item.message,
        'link': item.link,
        'is_read': item.is_read,
        'created_at': item.created_at.isoformat() if item.created_at else '',
        'created_at_display': item.created_at.strftime('%Y-%m-%d %H:%M') if item.created_at else '',
    }
