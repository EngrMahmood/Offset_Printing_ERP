"""Redis-backed presence tracking, separate from the Channels layer.

Presence needs a queryable set ("who is online right now"), which the
Channels group mechanism doesn't provide (groups are write-only fan-out).

Previously this used a plain Redis set + per-user connection refcount. That
scheme had a real bug: a refcount is only ever decremented by a clean
`disconnect()` call, so any ungraceful disconnect — a crashed browser tab, a
killed server process, a LAN PC losing power or network — left the user
stuck "online" forever with no way to self-correct. (Confirmed in the wild:
the online set accumulated entries for user ids that no longer even existed
in the database.)

This version uses two Redis sorted sets scored by timestamp instead, so
staleness self-heals on every read via a TTL cutoff — no refcounting, no
manual cleanup ever required:

- `chat:presence:conn`   — refreshed by a periodic server-side heartbeat for
  every open PresenceConsumer connection. A user has *some* connection open
  (status is "online" or "away") iff their score is within CONN_TTL seconds.
- `chat:presence:active` — refreshed whenever the client reports real
  activity (mouse/keyboard/visibility). A connected user is "online" if
  their score is within ACTIVE_TTL seconds, otherwise "away".

A user absent from `chat:presence:conn` (score expired) is simply "offline".
"""
import logging
import time

from django.conf import settings

logger = logging.getLogger(__name__)

CONN_KEY = 'chat:presence:conn'
ACTIVE_KEY = 'chat:presence:active'
REFCOUNT_KEY_FMT = 'chat:presence:refcount:{}'

# Heartbeats are sent every HEARTBEAT_INTERVAL seconds (see consumers.py);
# CONN_TTL gives a couple of missed beats' worth of grace before treating a
# connection as dead, so a slow tick or brief hiccup doesn't flap someone
# offline.
HEARTBEAT_INTERVAL = 20
CONN_TTL = 60
# "Away" after this long with no reported real activity, while still connected.
ACTIVE_TTL = 120

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


def heartbeat(user_id):
    """Refresh this user's connection liveness. Call once on connect and
    then every HEARTBEAT_INTERVAL seconds for as long as the socket is open."""
    try:
        _get_client().zadd(CONN_KEY, {str(user_id): time.time()})
    except Exception:
        logger.warning('chat presence heartbeat failed for user %s', user_id, exc_info=True)


def connect(user_id):
    """Call once per new PresenceConsumer connection (each browser tab).
    Bumps a connection refcount alongside the heartbeat, so a clean
    disconnect() of *one* tab doesn't wrongly clear presence while another
    tab from the same user is still open."""
    try:
        client = _get_client()
        client.incr(REFCOUNT_KEY_FMT.format(user_id))
        client.zadd(CONN_KEY, {str(user_id): time.time()})
    except Exception:
        logger.warning('chat presence connect failed for user %s', user_id, exc_info=True)


def disconnect(user_id):
    """Call once per PresenceConsumer disconnect. Only clears presence
    immediately if this was the user's last open tab — otherwise the
    remaining tab's own heartbeat loop keeps them correctly online. Even if
    this is never called (crash, killed process), the heartbeat TTL in
    get_statuses() expires the entry on its own within CONN_TTL seconds."""
    try:
        client = _get_client()
        key = REFCOUNT_KEY_FMT.format(user_id)
        count = client.decr(key)
        if count <= 0:
            client.delete(key)
            clear(user_id)
    except Exception:
        logger.warning('chat presence disconnect failed for user %s', user_id, exc_info=True)


def touch_activity(user_id):
    """Record real user activity (mouse/keyboard/visible-tab), promoting
    the user from 'away' back to 'online'. Implies a live connection too,
    so it also refreshes the connection heartbeat."""
    try:
        now = time.time()
        client = _get_client()
        client.zadd(CONN_KEY, {str(user_id): now})
        client.zadd(ACTIVE_KEY, {str(user_id): now})
    except Exception:
        logger.warning('chat presence touch_activity failed for user %s', user_id, exc_info=True)


def clear(user_id):
    """Best-effort immediate offline on a clean disconnect. Not required for
    correctness (TTL expiry is the real backstop) but makes the common case
    (closing a tab normally) reflect instantly instead of waiting for
    CONN_TTL to elapse."""
    try:
        client = _get_client()
        client.zrem(CONN_KEY, str(user_id))
        client.zrem(ACTIVE_KEY, str(user_id))
    except Exception:
        logger.warning('chat presence clear failed for user %s', user_id, exc_info=True)


def get_status(user_id):
    """Returns 'online', 'away', or 'offline' for a single user."""
    return get_statuses([user_id]).get(int(user_id), 'offline')


def get_statuses(user_ids=None):
    """Returns {user_id: 'online'|'away'} for currently-connected users.
    Users absent from the result are offline. Pass user_ids to limit/order
    the lookup (still O(1) per id via ZSCORE); omit it to return everyone
    currently connected."""
    try:
        client = _get_client()
        now = time.time()
        conn_cutoff = now - CONN_TTL
        active_cutoff = now - ACTIVE_TTL

        # Lazily evict stale entries so the sets don't grow unbounded —
        # cheap, and piggybacks on whatever request happens to ask.
        client.zremrangebyscore(CONN_KEY, '-inf', conn_cutoff)
        client.zremrangebyscore(ACTIVE_KEY, '-inf', active_cutoff)

        if user_ids is not None:
            connected_ids = set()
            for uid in user_ids:
                score = client.zscore(CONN_KEY, str(uid))
                if score is not None:
                    connected_ids.add(int(uid))
        else:
            connected_ids = {int(uid) for uid in client.zrangebyscore(CONN_KEY, conn_cutoff, '+inf')}

        if not connected_ids:
            return {}

        active_ids = {int(uid) for uid in client.zrangebyscore(ACTIVE_KEY, active_cutoff, '+inf')}
        return {uid: ('online' if uid in active_ids else 'away') for uid in connected_ids}
    except Exception:
        logger.warning('chat presence get_statuses failed', exc_info=True)
        return {}


def get_online_user_ids():
    """Back-compat helper: ids of anyone connected at all (online or away)."""
    return list(get_statuses().keys())
