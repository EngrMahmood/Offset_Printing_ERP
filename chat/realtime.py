"""Thin wrapper around the Channels layer so REST views can fan out events
without importing channels directly everywhere. No-ops safely if the channel
layer isn't configured (e.g. Channels not installed yet in Phase 1)."""


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
    async_to_sync(layer.group_send)(group, {'type': 'chat.event', 'payload': payload})


def broadcast_room_event(room_id, event_type, data):
    _send(f'chat_room_{room_id}', {'event': event_type, **data})


def notify_user(user_id, event_type, data):
    _send(f'chat_user_{user_id}', {'event': event_type, **data})


def notify_users(user_ids, event_type, data):
    for uid in user_ids:
        notify_user(uid, event_type, data)
