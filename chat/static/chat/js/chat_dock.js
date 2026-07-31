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

    function showToast(title, message, link, onClose) {
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
        function close() {
            toast.remove();
            if (onClose) onClose();
        }
        toast.querySelector('.erp-toast__close').addEventListener('click', close);
        toast.querySelector('.erp-toast__link').addEventListener('click', close);
        toastStack.appendChild(toast);
        window.setTimeout(close, 8000);
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

    function renderDockReactions(reactions) {
        // Read-only in dock popups — reacting requires the full chat page,
        // consistent with the "glance and reply" scope of docked windows.
        if (!reactions || !reactions.length) return '';
        return '<div class="chat-dock-msg__reactions">' + reactions.map(function (r) {
            return '<span class="chat-dock-msg__reaction-pill">' + r.emoji + ' ' + r.user_ids.length + '</span>';
        }).join('') + '</div>';
    }

    function appendDockMessage(win, msg) {
        if (!win) return;
        const container = win.el.querySelector('.chat-dock-win__messages');
        if (container.querySelector('.chat-dock-msg[data-message-id="' + msg.id + '"]')) return;

        const isOwn = msg.sender && String(msg.sender.id) === String(currentUserId);
        const wrap = document.createElement('div');
        wrap.className = 'chat-dock-msg ' + (isOwn ? 'is-own' : 'is-other');
        wrap.dataset.messageId = msg.id;
        const audioAttachment = !msg.is_deleted && (msg.attachments || []).find(function (a) { return a.file_type === 'audio'; });
        const bodyText = msg.is_deleted ? 'Message deleted' : escapeHtml(msg.body || (msg.attachments && msg.attachments.length && !audioAttachment ? '📎 Attachment' : ''));
        const audioHtml = audioAttachment ? '<audio controls preload="none" src="' + escapeHtml(audioAttachment.url) + '"></audio>' : '';
        wrap.innerHTML = '<div class="chat-dock-msg__bubble">' + bodyText + audioHtml + '</div>' + renderDockReactions(msg.reactions);
        container.appendChild(wrap);
        scrollToBottom(win);
        win.lastActivity = Date.now();

        if (!win.minimized) markRead(win.roomId);
    }

    function scrollToBottom(win) {
        const container = win.el.querySelector('.chat-dock-win__messages');
        container.scrollTop = container.scrollHeight;
    }

    function shakeDockWindow(win) {
        if (!win || !win.el) return;
        win.el.classList.remove('chat-buzz-shake');
        void win.el.offsetWidth;
        win.el.classList.add('chat-buzz-shake');
        setTimeout(function () { win.el.classList.remove('chat-buzz-shake'); }, 600);
        if (win.minimized) setMinimized(win, false);
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
            } else if (payload.event === 'message_edited') {
                const el = win.el.querySelector('.chat-dock-msg[data-message-id="' + payload.message.id + '"]');
                if (el) {
                    const bubble = el.querySelector('.chat-dock-msg__bubble');
                    bubble.textContent = payload.message.body || (payload.message.attachments && payload.message.attachments.length ? '📎 Attachment' : '');
                }
            } else if (payload.event === 'message_deleted') {
                const el = win.el.querySelector('.chat-dock-msg[data-message-id="' + payload.message_id + '"]');
                if (el) {
                    el.classList.add('is-deleted');
                    el.querySelector('.chat-dock-msg__bubble').textContent = 'Message deleted';
                }
            } else if (payload.event === 'reaction_updated') {
                const el = win.el.querySelector('.chat-dock-msg[data-message-id="' + payload.message_id + '"]');
                if (el) {
                    const existing = el.querySelector('.chat-dock-msg__reactions');
                    const html = renderDockReactions(payload.reactions);
                    if (existing) existing.outerHTML = html;
                    else if (html) el.querySelector('.chat-dock-msg__bubble').insertAdjacentHTML('afterend', html);
                }
            } else if (payload.event === 'group_updated') {
                const current = dockWindows.get(win.roomId);
                if (current) {
                    current.roomLabel = payload.room.display_name;
                    current.el.querySelector('.chat-dock-win__title').firstChild.textContent = payload.room.display_name;
                }
            } else if (payload.event === 'typing') {
                const label = win.el.querySelector('.chat-dock-win__typing');
                if (String(payload.user_id) !== String(currentUserId)) {
                    label.textContent = payload.is_typing ? 'typing…' : '';
                }
            } else if (payload.event === 'buzz') {
                if (String(payload.from_user_id) !== String(currentUserId)) {
                    shakeDockWindow(dockWindows.get(win.roomId));
                    if (window.ChatSound) window.ChatSound.playBuzz();
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
                const mentioned = (payload.mentioned_user_ids || []).map(String).indexOf(String(currentUserId)) !== -1;
                if (isViewingRoomElsewhere(payload.room_id)) return;
                if (window.ChatSound) window.ChatSound.playMessageDing();
                const win = openDockWindow(payload.room_id, payload.room_label);
                if (mentioned && win) win.el.classList.add('is-mentioned-flash');
                if (mentioned) {
                    showToast('You were mentioned', (payload.sender_name || 'Someone') + ' mentioned you in ' + (payload.room_label || 'a chat'), chatUrl + '?room=' + payload.room_id);
                }
            } else if (payload.event === 'unread_count_changed') {
                refreshBadge();
            } else if (payload.event === 'group_deleted') {
                if (dockWindows.has(payload.room_id)) closeDockWindow(payload.room_id);
                refreshBadge();
            } else if (payload.event === 'incoming_call' && !isViewingRoomElsewhere(payload.room_id)) {
                if (window.ChatSound) window.ChatSound.playRingtone();
                showToast('Incoming call', 'Call in ' + (payload.room_label || 'a chat'), chatUrl + '?room=' + payload.room_id, function () {
                    if (window.ChatSound) window.ChatSound.stopRingtone();
                });
            } else if (payload.event === 'buzz' && !isViewingRoomElsewhere(payload.room_id)) {
                // The room's own socket (if a dock window is already open) handles
                // the shake itself — this branch only covers the "not open yet"
                // case, same split as new_message's dock-open-vs-badge handling.
                if (!dockWindows.has(payload.room_id)) {
                    if (window.ChatSound) window.ChatSound.playBuzz();
                    const win = openDockWindow(payload.room_id, payload.from_display_name);
                    setTimeout(function () { shakeDockWindow(win); }, 50);
                }
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
