(function () {
    const root = document.getElementById('chat-app-root');
    if (!root) return;

    function getCookie(name) {
        const match = document.cookie.match(new RegExp('(^| )' + name + '=([^;]+)'));
        return match ? match[2] : '';
    }

    const csrfToken = (document.querySelector('[name=csrfmiddlewaretoken]') || {}).value || getCookie('csrftoken');
    const currentUserId = parseInt(root.dataset.currentUserId, 10);
    const urls = window.CHAT_URLS || {};

    const state = {
        rooms: [],
        currentRoomId: null,
        roomSocket: null,
        presenceSocket: null,
        typingTimeout: null,
        oldestLoadedMessageId: null,
    };

    function escapeHtml(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    function api(url, options) {
        options = options || {};
        options.headers = Object.assign({
            'X-CSRFToken': csrfToken,
            'X-Requested-With': 'XMLHttpRequest',
        }, options.headers || {});
        if (options.body && !(options.body instanceof FormData) && typeof options.body !== 'string') {
            options.headers['Content-Type'] = 'application/json';
            options.body = JSON.stringify(options.body);
        }
        return fetch(url, options).then(function (response) {
            if (response.status === 204) return null;
            return response.json().then(function (data) {
                if (!response.ok) throw { status: response.status, data: data };
                return data;
            });
        });
    }

    function wsUrl(path) {
        const scheme = window.location.protocol === 'https:' ? 'wss' : 'ws';
        return scheme + '://' + window.location.host + path;
    }

    // ---- Room list -------------------------------------------------------

    function loadRoomList() {
        api(urls.roomList).then(function (rooms) {
            state.rooms = rooms || [];
            renderRoomList();
        }).catch(function () {
            document.getElementById('chat-room-items').innerHTML = '<div class="chat-empty-note">Could not load conversations.</div>';
        });
    }

    function initials(name) {
        return (name || '?').trim().split(/\s+/).slice(0, 2).map(function (p) { return p[0]; }).join('').toUpperCase();
    }

    function renderRoomList() {
        const container = document.getElementById('chat-room-items');
        if (!state.rooms.length) {
            container.innerHTML = '<div class="chat-empty-note">No conversations yet. Start a new chat.</div>';
            return;
        }
        container.innerHTML = state.rooms.map(function (room) {
            const active = room.id === state.currentRoomId ? ' is-active' : '';
            const preview = room.last_message ? escapeHtml(room.last_message.body || '[Attachment]') : 'No messages yet';
            const badge = room.unread_count > 0
                ? '<span class="chat-room-item__badge">' + (room.unread_count > 99 ? '99+' : room.unread_count) + '</span>' : '';
            return (
                '<div class="chat-room-item' + active + '" data-room-id="' + room.id + '">' +
                '<div class="chat-room-item__avatar">' + escapeHtml(initials(room.display_name)) + '</div>' +
                '<div class="chat-room-item__body">' +
                '<div class="chat-room-item__name">' + escapeHtml(room.display_name) + '</div>' +
                '<div class="chat-room-item__preview">' + preview + '</div>' +
                '</div>' + badge +
                '</div>'
            );
        }).join('');
    }

    function upsertRoomInList(roomId, patch) {
        const room = state.rooms.find(function (r) { return r.id === roomId; });
        if (room) Object.assign(room, patch);
    }

    document.getElementById('chat-room-items').addEventListener('click', function (event) {
        const item = event.target.closest('.chat-room-item');
        if (!item) return;
        openRoom(parseInt(item.dataset.roomId, 10));
    });

    // ---- Active thread -----------------------------------------------------

    function openRoom(roomId) {
        if (state.roomSocket) {
            state.roomSocket.close();
            state.roomSocket = null;
        }
        state.currentRoomId = roomId;
        state.oldestLoadedMessageId = null;
        renderRoomList();

        document.getElementById('chat-thread-empty').hidden = true;
        document.getElementById('chat-thread-active').hidden = false;
        document.getElementById('chat-thread-messages').innerHTML = '<div class="chat-empty-note">Loading…</div>';

        api(urls.roomDetail.replace('/0/', '/' + roomId + '/')).then(function (room) {
            document.getElementById('chat-thread-name').textContent = room.display_name;
        });

        loadMessages(roomId);
        markRead(roomId);
        connectRoomSocket(roomId);
        window.ChatApp.currentRoomId = roomId;
    }

    function loadMessages(roomId) {
        api((urls.messages || '').replace('/0/', '/' + roomId + '/')).then(function (page) {
            const messages = (page.results || []).slice().reverse();
            const container = document.getElementById('chat-thread-messages');
            container.innerHTML = '';
            messages.forEach(function (msg) { appendMessage(msg, false); });
            if (messages.length) state.oldestLoadedMessageId = messages[0].id;
            scrollToBottom();
        });
    }

    function scrollToBottom() {
        const el = document.getElementById('chat-thread-messages');
        el.scrollTop = el.scrollHeight;
    }

    function renderAttachment(att) {
        if (!att) return '';
        if (att.file_type === 'image') {
            return '<div class="chat-msg__attachment"><a href="' + escapeHtml(att.url) + '" target="_blank" rel="noopener">' +
                '<img src="' + escapeHtml(att.thumbnail_url || att.url) + '" alt="' + escapeHtml(att.original_filename) + '"></a></div>';
        }
        return '<div class="chat-msg__attachment"><a href="' + escapeHtml(att.url) + '" target="_blank" rel="noopener">' +
            '<i class="fas fa-file"></i> ' + escapeHtml(att.original_filename) + '</a></div>';
    }

    function appendMessage(msg, prepend) {
        const container = document.getElementById('chat-thread-messages');
        // The sender's own REST response and the room-group WebSocket broadcast
        // both deliver the same message (the socket fan-out includes the
        // sender's own connection) — skip if it's already rendered.
        if (container.querySelector('.chat-msg[data-message-id="' + msg.id + '"]')) return;

        const isOwn = msg.sender && msg.sender.id === currentUserId;
        const wrap = document.createElement('div');
        wrap.className = 'chat-msg ' + (isOwn ? 'is-own' : 'is-other') + (msg.is_deleted ? ' is-deleted' : '');
        wrap.dataset.messageId = msg.id;

        const senderLine = !isOwn ? '<div class="chat-msg__sender">' + escapeHtml(msg.sender ? msg.sender.display_name : 'Unknown') + '</div>' : '';
        const bodyText = msg.is_deleted ? 'Message deleted' : escapeHtml(msg.body);
        const attachments = msg.is_deleted ? '' : (msg.attachments || []).map(renderAttachment).join('');
        const time = new Date(msg.created_at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const editedTag = msg.edited_at ? ' (edited)' : '';

        wrap.innerHTML = senderLine +
            '<div class="chat-msg__bubble">' + bodyText + attachments + '</div>' +
            '<div class="chat-msg__meta">' + time + editedTag + '</div>';

        if (prepend) {
            container.insertBefore(wrap, container.firstChild);
        } else {
            container.appendChild(wrap);
        }
    }

    function markRead(roomId) {
        api((urls.roomRead || '').replace('/0/', '/' + roomId + '/'), { method: 'POST' }).then(function () {
            upsertRoomInList(roomId, { unread_count: 0 });
            renderRoomList();
        }).catch(function () { /* ignore */ });
    }

    // ---- Sending messages ---------------------------------------------------

    document.getElementById('chat-composer').addEventListener('submit', function (event) {
        event.preventDefault();
        if (!state.currentRoomId) return;
        const input = document.getElementById('chat-message-input');
        const body = input.value.trim();
        if (!body) return;

        api((urls.messages || '').replace('/0/', '/' + state.currentRoomId + '/'), {
            method: 'POST',
            body: { body: body },
        }).then(function (msg) {
            input.value = '';
            appendMessage(msg, false);
            scrollToBottom();
            loadRoomList();
        }).catch(function (err) {
            alert((err.data && err.data.detail) || 'Could not send message.');
        });
    });

    document.getElementById('chat-message-input').addEventListener('input', function () {
        sendTyping(true);
        clearTimeout(state.typingTimeout);
        state.typingTimeout = setTimeout(function () { sendTyping(false); }, 2000);
    });

    function sendTyping(isTyping) {
        if (state.roomSocket && state.roomSocket.readyState === WebSocket.OPEN) {
            state.roomSocket.send(JSON.stringify({ event: 'typing', is_typing: isTyping }));
        }
    }

    // ---- WebSocket wiring -----------------------------------------------------

    function connectRoomSocket(roomId) {
        const socket = new WebSocket(wsUrl('/ws/chat/room/' + roomId + '/'));
        socket.onmessage = function (event) {
            const payload = JSON.parse(event.data);
            handleRoomEvent(roomId, payload);
        };
        socket.onclose = function () {
            // Only reconnect if the user is still viewing this room (openRoom()
            // closes the old socket before switching, so a stale onclose here
            // would otherwise race a newer connection for a different room).
            if (state.currentRoomId === roomId) {
                setTimeout(function () {
                    if (state.currentRoomId === roomId) connectRoomSocket(roomId);
                }, 3000);
            }
        };
        state.roomSocket = socket;
    }

    function handleRoomEvent(roomId, payload) {
        if (roomId !== state.currentRoomId) return;
        if (payload.event === 'message_created') {
            appendMessage(payload.message, false);
            scrollToBottom();
            if (payload.message.sender && payload.message.sender.id !== currentUserId) {
                markRead(roomId);
            }
            loadRoomList();
        } else if (payload.event === 'message_edited') {
            const el = document.querySelector('.chat-msg[data-message-id="' + payload.message.id + '"]');
            if (el) {
                const bubble = el.querySelector('.chat-msg__bubble');
                bubble.innerHTML = escapeHtml(payload.message.body) + (payload.message.attachments || []).map(renderAttachment).join('');
                const meta = el.querySelector('.chat-msg__meta');
                meta.textContent = meta.textContent.replace(/\s*\(edited\)$/, '') + ' (edited)';
            }
        } else if (payload.event === 'message_deleted') {
            const el = document.querySelector('.chat-msg[data-message-id="' + payload.message_id + '"]');
            if (el) {
                el.classList.add('is-deleted');
                el.querySelector('.chat-msg__bubble').textContent = 'Message deleted';
            }
        } else if (payload.event === 'typing') {
            const label = document.getElementById('chat-thread-typing');
            if (payload.user_id !== currentUserId) {
                label.textContent = payload.is_typing ? payload.username + ' is typing…' : '';
            }
        } else if (payload.event === 'participant_added' || payload.event === 'participant_removed') {
            api(urls.roomDetail.replace('/0/', '/' + roomId + '/')).then(function (room) {
                document.getElementById('chat-thread-name').textContent = room.display_name;
            });
        }
    }

    function connectPresenceSocket() {
        const socket = new WebSocket(wsUrl('/ws/chat/presence/'));
        socket.onmessage = function (event) {
            const payload = JSON.parse(event.data);
            if (payload.event === 'unread_count_changed' || payload.event === 'new_message') {
                loadRoomList();
            } else if (payload.event === 'incoming_call') {
                window.ChatApp.onIncomingCall(payload);
            }
        };
        socket.onclose = function () {
            setTimeout(connectPresenceSocket, 4000);
        };
        state.presenceSocket = socket;
    }

    // ---- New DM / group modals -------------------------------------------

    let chattableUsers = null;

    function loadChattableUsers() {
        if (chattableUsers) return Promise.resolve(chattableUsers);
        return api(urls.users).then(function (users) {
            chattableUsers = users;
            return users;
        });
    }

    function closeModal(modalEl) {
        modalEl.classList.remove('open');
    }

    const dmModal = document.getElementById('chat-new-dm-modal');
    document.addEventListener('click', function (event) {
        if (event.target.closest('[data-modal-open="chat-new-dm-modal"]')) {
            const select = document.getElementById('chat-dm-user-select');
            select.innerHTML = '<option>Loading…</option>';
            loadChattableUsers().then(function (users) {
                select.innerHTML = users.map(function (u) {
                    return '<option value="' + u.id + '">' + escapeHtml(u.display_name) + '</option>';
                }).join('') || '<option value="">No other chat users found</option>';
            });
        }
        if (event.target.closest('[data-modal-open="chat-new-group-modal"]')) {
            const select = document.getElementById('chat-group-members-select');
            select.innerHTML = '<option>Loading…</option>';
            loadChattableUsers().then(function (users) {
                select.innerHTML = users.map(function (u) {
                    return '<option value="' + u.id + '">' + escapeHtml(u.display_name) + '</option>';
                }).join('') || '<option value="">No other chat users found</option>';
            });
        }
    });

    const dmStartBtn = document.getElementById('chat-dm-start-btn');
    if (dmStartBtn) {
        dmStartBtn.addEventListener('click', function () {
            const select = document.getElementById('chat-dm-user-select');
            const userId = select.value;
            if (!userId) return;
            api(urls.roomList, { method: 'POST', body: { room_type: 'dm', user_id: userId } })
                .then(function (room) {
                    closeModal(dmModal);
                    loadRoomList();
                    openRoom(room.id);
                }).catch(function (err) { alert((err.data && err.data.detail) || 'Could not start chat.'); });
        });
    }

    const groupModal = document.getElementById('chat-new-group-modal');
    const groupCreateBtn = document.getElementById('chat-group-create-btn');
    if (groupCreateBtn) {
        groupCreateBtn.addEventListener('click', function () {
            const name = document.getElementById('chat-group-name-input').value.trim();
            if (!name) { alert('Group name is required.'); return; }
            const select = document.getElementById('chat-group-members-select');
            const memberIds = Array.from(select.selectedOptions).map(function (o) { return o.value; }).filter(Boolean);
            api(urls.roomList, { method: 'POST', body: { room_type: 'group', name: name, member_ids: memberIds } })
                .then(function (room) {
                    closeModal(groupModal);
                    document.getElementById('chat-group-name-input').value = '';
                    loadRoomList();
                    openRoom(room.id);
                }).catch(function (err) { alert((err.data && err.data.detail) || 'Could not create group.'); });
        });
    }

    // Exposed for chat_upload.js / webrtc_call.js
    window.ChatApp = {
        api: api,
        urls: urls,
        csrfToken: csrfToken,
        currentUserId: currentUserId,
        currentRoomId: null,
        wsUrl: wsUrl,
        appendMessage: appendMessage,
        scrollToBottom: scrollToBottom,
        loadRoomList: loadRoomList,
        onIncomingCall: function () { /* overridden by webrtc_call.js */ },
    };

    loadRoomList();
    connectPresenceSocket();

    // Deep link from a toast notification, e.g. /chat/?room=12
    const deepLinkRoomId = parseInt(new URLSearchParams(window.location.search).get('room'), 10);
    if (deepLinkRoomId) {
        api(urls.roomDetail.replace('/0/', '/' + deepLinkRoomId + '/'))
            .then(function () { openRoom(deepLinkRoomId); })
            .catch(function () { /* room no longer accessible — ignore */ });
    }
})();
