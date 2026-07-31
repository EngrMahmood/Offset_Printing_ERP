"""Thin wrapper around the Channels layer so REST views can fan out events
without importing channels directly everywhere. No-ops safely if the channel
layer isn't configured (e.g. Channels not installed yet in Phase 1)."""

import logging

logger = logging.getLogger(__name__)


def _get_layer():
    try:
        from channels.layers import get_channel_layer
    except ImportError:
        return None
    return get_channel_layer()


def _send(group, payload):
    layer = _get_layer()
    if layer is None:
        return
    from asgiref.sync import async_to_sync
    try:
        async_to_sync(layer.group_send)(group, {'type': 'chat.event', 'payload': payload})
    except Exception:
        # The DB write behind this event has already committed by the time
        # views call broadcast_room_event/notify_user(s) — a flaky Redis
        # connection should never turn into a 500 on message send/edit/
        # delete. Clients resync via reconnect + room reload/loadRoomList().
        logger.warning('chat realtime broadcast to %s failed', group, exc_info=True)


def broadcast_room_event(room_id, event_type, data):
    _send(f'chat_room_{room_id}', {'event': event_type, **data})


def notify_user(user_id, event_type, data):
    _send(f'chat_user_{user_id}', {'event': event_type, **data})


def notify_users(user_ids, event_type, data):
    for uid in user_ids:
        notify_user(uid, event_type, data)
