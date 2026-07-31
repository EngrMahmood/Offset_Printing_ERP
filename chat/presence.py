"""Redis-backed online-user tracking, separate from the Channels layer.

Presence needs a queryable set ("who is online right now"), which the
Channels group mechanism doesn't provide (groups are write-only fan-out).
Uses a plain Redis set + a per-user connection refcount (so a user with two
browser tabs open isn't marked offline when only one tab closes).
"""
import logging

from django.conf import settings

logger = logging.getLogger(__name__)

ONLINE_SET_KEY = 'chat:online_users'
REFCOUNT_KEY_FMT = 'chat:online_refcount:{}'

_client = None


def _get_client():
    global _client
    if _client is None:
        import redis
        # protocol=2 (RESP2): the Redis instance this project targets is
        # v5.0.14, which predates the RESP3 HELLO handshake redis-py
        # negotiates by default — same reasoning as CHANNEL_LAYERS/CACHES
        # in settings.py.
        _client = redis.from_url(settings.REDIS_URL, decode_responses=True, protocol=2)
    return _client


def mark_online(user_id):
    """Increments this user's connection refcount. Returns True only on the
    0 -> 1 transition (i.e. they were actually offline before this)."""
    try:
        client = _get_client()
        count = client.incr(REFCOUNT_KEY_FMT.format(user_id))
        if count == 1:
            client.sadd(ONLINE_SET_KEY, user_id)
            return True
        return False
    except Exception:
        logger.warning('chat presence mark_online failed for user %s', user_id, exc_info=True)
        return False


def mark_offline(user_id):
    """Decrements this user's connection refcount. Returns True only on the
    1 -> 0 transition (their last open tab/connection just closed)."""
    try:
        client = _get_client()
        key = REFCOUNT_KEY_FMT.format(user_id)
        count = client.decr(key)
        if count <= 0:
            client.delete(key)
            client.srem(ONLINE_SET_KEY, user_id)
            return True
        return False
    except Exception:
        logger.warning('chat presence mark_offline failed for user %s', user_id, exc_info=True)
        return False


def get_online_user_ids():
    try:
        client = _get_client()
        return [int(uid) for uid in client.smembers(ONLINE_SET_KEY)]
    except Exception:
        logger.warning('chat presence get_online_user_ids failed', exc_info=True)
        return []
