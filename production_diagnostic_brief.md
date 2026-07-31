# Production diagnostic — chat module

## Context

This is the Django ERP at `Offset_Printing_ERP` (chat app: real-time messaging via Django Channels + Redis + DRF). A large batch of chat features (edit/delete, reactions, mentions, pins, forwarding, voice messages, notification sounds, a "buzz" nudge button, group settings/delete) was just merged to `main` and pulled to this production machine. Two problems appeared after the update:

1. **A group chat that was visible before is no longer showing** in the room list.
2. **An error occurs when starting a new chat** (clicking "New Chat" to start a DM).

There's also a standing question about whether the production Redis version is compatible — this codebase already works around an old Redis (v5.0.14 on the original dev machine) by forcing RESP2 protocol negotiation, since that version predates the RESP3 `HELLO` handshake `redis-py` tries by default.

Do NOT make destructive changes (no hard deletes, no `git reset --hard`, no dropping tables) without reporting findings first and getting confirmation. This is a diagnostic pass.

## Step 1 — Confirm the deploy actually landed

```bash
cd <production repo path>
git log --oneline -3
git status
python manage.py showmigrations chat
```
Expect the latest commit to be `01bac3c` ("Add WhatsApp-standard chat features: ..."), and `chat` migrations `0001_initial`, `0002_chatroom_avatar_chatroom_description_and_more`, `0003_alter_attachment_file_type` all marked `[X]`. If `0002`/`0003` are NOT applied, run:
```bash
python manage.py migrate
```
then restart the server process and re-test both issues before going further — an unapplied migration would fully explain both symptoms (missing `description`/`avatar`/`pinned_message` columns on `ChatRoom`, missing `'audio'` attachment type).

## Step 2 — Check Redis version and protocol config

```bash
redis-cli INFO server | grep redis_version
```
(On Windows, if `redis-cli` isn't on PATH, check wherever the Redis/Memurai service was installed, e.g. `"C:\Program Files\Memurai\memurai-cli.exe" INFO server`.)

This codebase's `Offset_ERP/settings.py` already sets `protocol: 2` (RESP2) on both `CHANNEL_LAYERS` and `CACHES` specifically to support Redis < 6. Grep it to confirm:
```bash
grep -n "protocol" Offset_ERP/settings.py
```
Should show `'protocol': 2,` in two places (`CHANNEL_LAYERS['default']['CONFIG']['hosts'][0]` and `CACHES['default']['OPTIONS']`). If the Redis version is 6+, RESP2 still works fine (backward compatible) — no action needed either way, just report the version.

If Redis errors show up in server logs (`redis.exceptions.ResponseError`, `TimeoutError`, `ConnectionError`), paste them.

## Step 3 — Find out why the group disappeared

The new "Delete Group" feature (superuser-only) does a **soft-archive**, not a hard delete — sets `ChatRoom.is_archived = True`, which the room list now excludes. Check whether that's what happened:

```bash
python manage.py shell -c "
from chat.models import ChatRoom
for r in ChatRoom.objects.filter(room_type='group').order_by('-id'):
    print(r.id, repr(r.name), 'archived=', r.is_archived, 'created_by=', r.created_by_id, 'last_msg=', r.last_message_at)
"
```

- If the missing group shows `archived= True`: that's the cause, and it's recoverable (no data lost — messages/attachments are untouched). To restore it:
  ```bash
  python manage.py shell -c "
  from chat.models import ChatRoom
  r = ChatRoom.objects.get(pk=<ROOM_ID>)
  r.is_archived = False
  r.save(update_fields=['is_archived'])
  print('restored', r.id)
  "
  ```
  Report back the room id/name before running the restore, in case it was intentionally deleted.
- If the group doesn't appear in this query at all (not just archived, truly gone), or `archived= False` but still missing from the UI: report that — it points to a different bug (e.g. the requesting user's `ChatParticipant` row got a `left_at` timestamp, or a permission/serializer error), not the archive feature. In that case also check:
  ```bash
  python manage.py shell -c "
  from chat.models import ChatParticipant
  for p in ChatParticipant.objects.filter(room_id=<ROOM_ID>):
      print(p.user.username, 'left_at=', p.left_at, 'role=', p.role)
  "
  ```

## Step 4 — Reproduce the "New Chat" error

1. Open the site in a browser, open DevTools (F12) → Console tab, and Network tab.
2. Click "New Chat", pick a user, click "Start Chat".
3. Capture and report:
   - Any red error in the Console (full text + stack trace).
   - The Network tab entry for the request this triggers (should be `POST /api/chat/rooms/`) — status code and response body.
   - The corresponding server-side log line/traceback at that timestamp (wherever Daphne's stdout/stderr or the log file is captured — check `logs/` in the repo root, or the console window running the server).

Likely culprits to check once you have the traceback:
- `chat/api.py`'s `RoomViewSet.create()` (handles `room_type: 'dm'`) — look for anything referencing a field that might not exist yet if Step 1's migration check failed.
- `chat/serializers.py`'s `ChatRoomDetailSerializer`/`ChatRoomListSerializer` — the response after creating a room serializes it immediately; if `description`/`avatar_url`/`pinned_message` fields were added to the serializer but migration 0002 wasn't applied, this is where it'd blow up (`OperationalError: no such column`).
- `chat/static/chat/js/chat_socket.js` around line 874-905 (`dmModal`/`dmStartBtn` handlers) if the error is purely client-side JS (e.g. a stale cached `chat_socket.js` from before the version bump — check the `<script src="...chat_socket.js?v=8">` tag actually loaded `v=8` in the Network tab, not a cached `v=7`; hard-refresh with Ctrl+Shift+R if unsure).

## Report back

Please report: the migration status, Redis version, whether the group was archived (and its room id/name if so — don't restore without confirming it wasn't intentional), and the exact error text + traceback from the New Chat flow. Don't restart/reset anything destructive beyond what's listed above without checking in first.
