(function () {
    // Docked windows never render on the full /chat/ page itself — that page
    // already gives the full experience (chat_socket.js), a mini popup on top
    // would just be a redundant second UI for the same room.
    if (document.getElementById('chat-app-root')) return;

    const navLink = document.getElementById('chat-nav-link');
    if (!navLink) return; // user has no chat access

    const badge = document.getElementById('chat-nav-badge');
    const roomListUrl = navLink.dataset.roomListUrl;
    const chatUrl = navLink.dataset.chatUrl;
    const currentUserId = navLink.dataset.currentUserId;
    const toastStack = document.getElementById('erp-toast-stack');

    const STORAGE_KEY = 'chat_dock_open_rooms';
    const MAX_EXPANDED = 3;
    const dockWindows = new Map(); // roomId -> { roomId, roomLabel, minimized, socket, lastActivity, el }

    function getCookie(name) {
        const match = document.cookie.match(new RegExp('(^| )' + name + '=([^;]+)'));
        return match ? match[2] : '';
    }
    const csrfToken = getCookie('csrftoken');

    function escapeHtml(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    function wsUrl(path) {
        const scheme = window.location.protocol === 'https:' ? 'wss' : 'ws';
        return scheme + '://' + window.location.host + path;
    }

    function apiFetch(url, options) {
        options = options || {};
        options.headers = Object.assign({
            'X-CSRFToken': csrfToken,
            'X-Requested-With': 'XMLHttpRequest',
        }, options.headers || {});
        if (options.body && typeof options.body !== 'string') {
            options.headers['Content-Type'] = 'application/json';
            options.body = JSON.stringify(options.body);
        }
        return fetch(url, options).then(function (r) {
            if (r.status === 204) return null;
            return r.json().then(function (data) {
                if (!r.ok) throw { status: r.status, data: data };
                return data;
            });
        });
    }

    // ---- Nav badge (unchanged from the old chat_notify_badge.js) -------------

    function setBadge(count) {
        if (!badge) return;
        if (count > 0) {
            badge.textContent = count > 99 ? '99+' : String(count);
            badge.style.display = 'inline-block';
        } else {
            badge.textContent = '';
            badge.style.display = 'none';
        }
    }

    function refreshBadge() {
        apiFetch(roomListUrl).then(function (rooms) {
            const total = (rooms || []).reduce(function (sum, r) { return sum + (r.unread_count || 0); }, 0);
            setBadge(total);
        }).catch(function () { /* ignore */ });
    }

    function showToast(title, message, link) {
        if (!toastStack) return;
        const toast = document.createElement('div');
        toast.className = 'erp-toast';
        toast.innerHTML =
            '<p class="erp-toast__title">' + escapeHtml(title) + '</p>' +
            (message ? '<p class="erp-toast__message">' + escapeHtml(message) + '</p>' : '') +
            '<div class="erp-toast__actions">' +
            '<a class="erp-toast__link" href="' + escapeHtml(link) + '">Open</a>' +
            '<button type="button" class="erp-toast__close">Dismiss</button>' +
            '</div>';
        toast.querySelector('.erp-toast__close').addEventListener('click', function () { toast.remove(); });
        toastStack.appendChild(toast);
        window.setTimeout(function () { toast.remove(); }, 8000);
    }

    // ---- Docked window persistence -------------------------------------------

    function persistState() {
        const list = Array.from(dockWindows.values()).map(function (w) {
            return { room_id: w.roomId, minimized: w.minimized };
        });
        try { localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); } catch (e) { /* ignore */ }
    }

    function loadPersistedRoomIds() {
        try {
            return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
        } catch (e) {
            return [];
        }
    }

    function getContainer() {
        let container = document.getElementById('chat-dock-container');
        if (!container) {
            container = document.createElement('div');
            container.id = 'chat-dock-container';
            document.body.appendChild(container);
        }
        return container;
    }

    // ---- Window lifecycle -------------------------------------------------

    function enforceExpandedCap(excludeRoomId) {
        const expanded = Array.from(dockWindows.values())
            .filter(function (w) { return !w.minimized && w.roomId !== excludeRoomId; })
            .sort(function (a, b) { return a.lastActivity - b.lastActivity; });
        while (expanded.length >= MAX_EXPANDED) {
            const oldest = expanded.shift();
            setMinimized(oldest, true);
        }
    }

    function openDockWindow(roomId, roomLabelHint) {
        const existing = dockWindows.get(roomId);
        if (existing) {
            bringToFront(existing);
            if (existing.minimized) setMinimized(existing, false);
            return existing;
        }

        enforceExpandedCap(roomId);

        const win = {
            roomId: roomId,
            roomLabel: roomLabelHint || 'Chat',
            minimized: false,
            socket: null,
            lastActivity: Date.now(),
            el: buildWindowEl(roomId, roomLabelHint || 'Chat'),
        };
        dockWindows.set(roomId, win);
        getContainer().appendChild(win.el);

        // Always fetch the room detail (for an accurate label) + recent history.
        fetchRoomDetail(win);
        fetchRecentMessages(win);
        connectWindowSocket(win);
        markRead(roomId);
        persistState();
        return win;
    }

    function bringToFront(win) {
        const container = getContainer();
        container.insertBefore(win.el, container.firstChild);
        win.lastActivity = Date.now();
    }

    function setMinimized(win, minimized) {
        win.minimized = minimized;
        win.el.classList.toggle('is-minimized', minimized);
        if (!minimized) {
            win.lastActivity = Date.now();
            markRead(win.roomId);
            scrollToBottom(win);
        }
        persistState();
    }

    function closeDockWindow(roomId) {
        const win = dockWindows.get(roomId);
        if (!win) return;
        if (win.socket) win.socket.close();
        win.el.remove();
        dockWindows.delete(roomId);
        persistState();
    }

    function markRead(roomId) {
        apiFetch('/api/chat/rooms/' + roomId + '/read/', { method: 'POST' })
            .then(refreshBadge)
            .catch(function () { /* ignore */ });
    }

    // ---- DOM building -------------------------------------------------------

    function buildWindowEl(roomId, roomLabel) {
        const el = document.createElement('div');
        el.className = 'chat-dock-win';
        el.dataset.roomId = roomId;
        el.innerHTML =
            '<div class="chat-dock-win__header">' +
            '<div class="chat-dock-win__title">' + escapeHtml(roomLabel) + '<div class="chat-dock-win__typing"></div></div>' +
            '<div class="chat-dock-win__controls">' +
            '<button type="button" class="chat-dock-win__minimize" title="Minimize">&#8211;</button>' +
            '<button type="button" class="chat-dock-win__close" title="Close">&times;</button>' +
            '</div></div>' +
            '<div class="chat-dock-win__body">' +
            '<div class="chat-dock-win__messages"></div>' +
            '<a class="chat-dock-win__open-full" href="' + escapeHtml(chatUrl) + '?room=' + roomId + '">Open full chat &rarr;</a>' +
            '<form class="chat-dock-win__footer">' +
            '<input type="text" class="chat-dock-win__input" placeholder="Type a message…" autocomplete="off">' +
            '<button type="submit" class="chat-dock-win__send"><i class="fas fa-paper-plane"></i></button>' +
            '</form></div>';

        el.querySelector('.chat-dock-win__header').addEventListener('click', function (event) {
            if (event.target.closest('.chat-dock-win__controls')) return;
            const win = dockWindows.get(roomId);
            if (win) setMinimized(win, !win.minimized);
        });
        el.querySelector('.chat-dock-win__minimize').addEventListener('click', function (event) {
            event.stopPropagation();
            const win = dockWindows.get(roomId);
            if (win) setMinimized(win, !win.minimized);
        });
        el.querySelector('.chat-dock-win__close').addEventListener('click', function (event) {
            event.stopPropagation();
            closeDockWindow(roomId);
        });
        el.querySelector('.chat-dock-win__footer').addEventListener('submit', function (event) {
            event.preventDefault();
            const input = el.querySelector('.chat-dock-win__input');
            const body = input.value.trim();
            if (!body) return;
            input.value = '';
            apiFetch('/api/chat/rooms/' + roomId + '/messages/', { method: 'POST', body: { body: body } })
                .then(function (msg) { appendDockMessage(dockWindows.get(roomId), msg); })
                .catch(function () { input.value = body; });
        });

        return el;
    }

    function appendDockMessage(win, msg) {
        if (!win) return;
        const container = win.el.querySelector('.chat-dock-win__messages');
        if (container.querySelector('.chat-dock-msg[data-message-id="' + msg.id + '"]')) return;

        const isOwn = msg.sender && String(msg.sender.id) === String(currentUserId);
        const wrap = document.createElement('div');
        wrap.className = 'chat-dock-msg ' + (isOwn ? 'is-own' : 'is-other');
        wrap.dataset.messageId = msg.id;
        const bodyText = msg.is_deleted ? 'Message deleted' : escapeHtml(msg.body || (msg.attachments && msg.attachments.length ? '📎 Attachment' : ''));
        wrap.innerHTML = '<div class="chat-dock-msg__bubble">' + bodyText + '</div>';
        container.appendChild(wrap);
        scrollToBottom(win);
        win.lastActivity = Date.now();

        if (!win.minimized) markRead(win.roomId);
    }

    function scrollToBottom(win) {
        const container = win.el.querySelector('.chat-dock-win__messages');
        container.scrollTop = container.scrollHeight;
    }

    function fetchRoomDetail(win) {
        apiFetch('/api/chat/rooms/' + win.roomId + '/').then(function (room) {
            win.roomLabel = room.display_name;
            win.el.querySelector('.chat-dock-win__title').firstChild.textContent = room.display_name;
        }).catch(function () { /* ignore */ });
    }

    function fetchRecentMessages(win) {
        apiFetch('/api/chat/rooms/' + win.roomId + '/messages/').then(function (page) {
            (page.results || []).slice().reverse().forEach(function (msg) { appendDockMessage(win, msg); });
        }).catch(function () { /* ignore */ });
    }

    function connectWindowSocket(win) {
        const socket = new WebSocket(wsUrl('/ws/chat/room/' + win.roomId + '/'));
        socket.onmessage = function (event) {
            const payload = JSON.parse(event.data);
            if (payload.event === 'message_created') {
                appendDockMessage(dockWindows.get(win.roomId), payload.message);
            } else if (payload.event === 'typing') {
                const label = win.el.querySelector('.chat-dock-win__typing');
                if (String(payload.user_id) !== String(currentUserId)) {
                    label.textContent = payload.is_typing ? 'typing…' : '';
                }
            }
        };
        socket.onclose = function () {
            if (dockWindows.has(win.roomId)) {
                setTimeout(function () {
                    const current = dockWindows.get(win.roomId);
                    if (current) connectWindowSocket(current);
                }, 3000);
            }
        };
        win.socket = socket;
    }

    // ---- Presence socket (badge + toast + auto-popup trigger) ----------------

    function isViewingRoomElsewhere(roomId) {
        // The full /chat/ page bails out of this whole script at the top, so
        // this only ever runs on non-chat pages — kept as a guard in case
        // that changes later.
        return window.ChatApp && window.ChatApp.currentRoomId === roomId;
    }

    function connectPresence() {
        const socket = new WebSocket(wsUrl('/ws/chat/presence/'));
        socket.onmessage = function (event) {
            const payload = JSON.parse(event.data);
            if (payload.event === 'new_message') {
                refreshBadge();
                if (isViewingRoomElsewhere(payload.room_id)) return;
                openDockWindow(payload.room_id, payload.room_label);
            } else if (payload.event === 'unread_count_changed') {
                refreshBadge();
            } else if (payload.event === 'incoming_call' && !isViewingRoomElsewhere(payload.room_id)) {
                showToast('Incoming call', 'Call in ' + (payload.room_label || 'a chat'), chatUrl + '?room=' + payload.room_id);
            }
        };
        socket.onclose = function () {
            setTimeout(connectPresence, 4000);
        };
    }

    // ---- Init -----------------------------------------------------------------

    loadPersistedRoomIds().forEach(function (entry) {
        const win = openDockWindow(entry.room_id, null);
        if (entry.minimized) setMinimized(win, true);
    });

    refreshBadge();
    connectPresence();
})();
