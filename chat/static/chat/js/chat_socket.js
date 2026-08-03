(function () {
    const root = document.getElementById('chat-app-root');
    if (!root) return;

    function getCookie(name) {
        const match = document.cookie.match(new RegExp('(^| )' + name + '=([^;]+)'));
        return match ? match[2] : '';
    }

    const csrfToken = (document.querySelector('[name=csrfmiddlewaretoken]') || {}).value || getCookie('csrftoken');
    const currentUserId = parseInt(root.dataset.currentUserId, 10);
    const canManageGroup = root.dataset.canManageGroup === '1';
    const isSuperuser = root.dataset.isSuperuser === '1';
    const urls = window.CHAT_URLS || {};

    const state = {
        rooms: [],
        currentRoomId: null,
        roomSocket: null,
        presenceSocket: null,
        typingTimeout: null,
        oldestLoadedMessageId: null,
        userStatuses: {}, // userId -> 'online' | 'away'; absent = offline
        currentRoomDetail: null,
    };

    function getUserStatus(userId) {
        return state.userStatuses[userId] || 'offline';
    }

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

    // urls.messageDetail is templated from {% url 'message-detail' 0 0 %}, i.e.
    // ".../rooms/0/messages/0/" — replace each 0 positionally (room id first).
    function messageDetailUrl(roomId, messageId) {
        return (urls.messageDetail || '').replace('rooms/0/', 'rooms/' + roomId + '/').replace('messages/0/', 'messages/' + messageId + '/');
    }

    function messageReadByUrl(roomId, messageId) {
        return (urls.messageReadBy || '').replace('rooms/0/', 'rooms/' + roomId + '/').replace('messages/0/', 'messages/' + messageId + '/');
    }

    function messageReactionsUrl(roomId, messageId) {
        return (urls.messageReactions || '').replace('rooms/0/', 'rooms/' + roomId + '/').replace('messages/0/', 'messages/' + messageId + '/');
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
            const otherStatus = room.other_user_id ? getUserStatus(room.other_user_id) : 'offline';
            const avatarClass = 'chat-room-item__avatar' + (otherStatus !== 'offline' ? ' is-' + otherStatus : '');
            return (
                '<div class="chat-room-item' + active + '" data-room-id="' + room.id + '">' +
                '<div class="' + avatarClass + '">' + escapeHtml(initials(room.display_name)) + '</div>' +
                '<div class="chat-room-item__body">' +
                '<div class="chat-room-item__name">' + escapeHtml(room.display_name) + '</div>' +
                '<div class="chat-room-item__preview">' + preview + '</div>' +
                '</div>' + badge +
                '</div>'
            );
        }).join('');
    }

    function updatePresenceLabel() {
        const room = state.rooms.find(function (r) { return r.id === state.currentRoomId; });
        const label = document.getElementById('chat-thread-presence');
        if (!label) return;
        const status = room && room.other_user_id ? getUserStatus(room.other_user_id) : 'offline';
        label.classList.remove('is-online', 'is-away');
        if (status === 'online') {
            label.textContent = 'Online';
            label.classList.add('is-online');
        } else if (status === 'away') {
            label.textContent = 'Away';
            label.classList.add('is-away');
        } else {
            label.textContent = '';
        }
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

        const roomDetailPromise = api(urls.roomDetail.replace('/0/', '/' + roomId + '/')).then(function (room) {
            document.getElementById('chat-thread-name').textContent = room.display_name;
            upsertRoomInList(roomId, { other_user_id: room.other_user_id });
            updatePresenceLabel();
            state.currentRoomDetail = room;
            renderPinnedBanner(room.pinned_message);
        });

        // Mention highlighting needs state.currentRoomDetail.participants (for
        // display-name lookup), so wait for room detail before rendering
        // messages rather than racing the two requests.
        roomDetailPromise.then(function () {
            loadMessages(roomId).then(function () { seedSeenIndicator(state.currentRoomDetail); });
        });
        markRead(roomId);
        connectRoomSocket(roomId);
        window.ChatApp.currentRoomId = roomId;
        updatePresenceLabel();
    }

    function loadMessages(roomId) {
        return api((urls.messages || '').replace('/0/', '/' + roomId + '/')).then(function (page) {
            const messages = (page.results || []).slice().reverse();
            const container = document.getElementById('chat-thread-messages');
            container.innerHTML = '';
            messages.forEach(function (msg) { appendMessage(msg, false); });
            if (messages.length) state.oldestLoadedMessageId = messages[0].id;
            scrollToBottom();
        });
    }

    // Read receipts only broadcast live (read_receipt_updated) — a sender who
    // opens the room later, after the reader already marked it read while
    // disconnected, would otherwise never see "Seen". Seed it from the room's
    // participant watermarks (already in ChatRoomDetailSerializer) instead.
    function seedSeenIndicator(room) {
        if (!room || !room.participants) return;
        let bestReaderId = null;
        let bestWatermark = 0;
        room.participants.forEach(function (p) {
            if (p.user.id === currentUserId || !p.last_read_message_id) return;
            if (p.last_read_message_id > bestWatermark) {
                bestWatermark = p.last_read_message_id;
                bestReaderId = p.user.id;
            }
        });
        if (bestReaderId) updateSeenIndicator(bestReaderId, bestWatermark);
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
        if (att.file_type === 'audio') {
            return '<div class="chat-msg__attachment"><audio controls preload="none" src="' + escapeHtml(att.url) + '"></audio></div>';
        }
        return '<div class="chat-msg__attachment"><a href="' + escapeHtml(att.url) + '" target="_blank" rel="noopener">' +
            '<i class="fas fa-file"></i> ' + escapeHtml(att.original_filename) + '</a></div>';
    }

    // ---- @mentions ---------------------------------------------------------

    function mentionDisplayNames(mentionUserIds) {
        if (!mentionUserIds || !mentionUserIds.length || !state.currentRoomDetail || !state.currentRoomDetail.participants) return [];
        const idSet = new Set(mentionUserIds);
        return state.currentRoomDetail.participants
            .filter(function (p) { return idSet.has(p.user.id); })
            .map(function (p) { return p.user.display_name; });
    }

    function highlightMentions(escapedBody, mentionUserIds) {
        const names = mentionDisplayNames(mentionUserIds);
        if (!names.length) return escapedBody;
        let result = escapedBody;
        names.forEach(function (name) {
            const escapedName = escapeHtml(name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            result = result.replace(new RegExp('@' + escapedName, 'g'), '<span class="chat-msg__mention">@' + escapeHtml(name) + '</span>');
        });
        return result;
    }

    function renderMentionSuggestions(query) {
        const box = document.getElementById('chat-mention-suggestions');
        if (!state.currentRoomDetail || state.currentRoomDetail.room_type !== 'group' || !state.currentRoomDetail.participants) {
            box.hidden = true;
            return;
        }
        const q = query.toLowerCase();
        const matches = state.currentRoomDetail.participants
            .filter(function (p) { return p.user.id !== currentUserId && p.user.display_name.toLowerCase().indexOf(q) !== -1; })
            .slice(0, 6);
        if (!matches.length) { box.hidden = true; return; }
        box.hidden = false;
        box.innerHTML = matches.map(function (p) {
            return '<div class="chat-mention-suggestions__item" data-name="' + escapeHtml(p.user.display_name) + '">' + escapeHtml(p.user.display_name) + '</div>';
        }).join('');
    }

    function renderReactionPills(reactions) {
        if (!reactions || !reactions.length) return '<div class="chat-msg__reactions"></div>';
        return '<div class="chat-msg__reactions">' + reactions.map(function (r) {
            const mine = r.user_ids.indexOf(currentUserId) !== -1;
            return '<button type="button" class="chat-msg__reaction-pill' + (mine ? ' is-mine' : '') + '" data-emoji="' + escapeHtml(r.emoji) + '">' +
                r.emoji + ' <span>' + r.user_ids.length + '</span></button>';
        }).join('') + '</div>';
    }

    function appendMessage(msg, prepend) {
        const container = document.getElementById('chat-thread-messages');
        // The sender's own REST response and the room-group WebSocket broadcast
        // both deliver the same message (the socket fan-out includes the
        // sender's own connection) — skip if it's already rendered.
        if (container.querySelector('.chat-msg[data-message-id="' + msg.id + '"]')) return;

        const isOwn = msg.sender && msg.sender.id === currentUserId;
        const wrap = document.createElement('div');
        const isMentioned = (msg.mentions || []).indexOf(currentUserId) !== -1;
        wrap.className = 'chat-msg ' + (isOwn ? 'is-own' : 'is-other') + (msg.is_deleted ? ' is-deleted' : '') + (isMentioned ? ' is-mentioning-me' : '');
        wrap.dataset.messageId = msg.id;

        const senderLine = !isOwn ? '<div class="chat-msg__sender">' + escapeHtml(msg.sender ? msg.sender.display_name : 'Unknown') + '</div>' : '';
        const bodyText = msg.is_deleted ? 'Message deleted' : highlightMentions(escapeHtml(msg.body), msg.mentions);
        const attachments = msg.is_deleted ? '' : (msg.attachments || []).map(renderAttachment).join('');
        const time = new Date(msg.created_at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const editedTag = msg.edited_at ? ' (edited)' : '';
        const menuBtn = !msg.is_deleted
            ? '<button type="button" class="chat-msg__menu-btn" title="Message options"><i class="fas fa-ellipsis-v"></i></button>' : '';
        const reactBtn = !msg.is_deleted
            ? '<button type="button" class="chat-msg__react-btn" title="React"><i class="far fa-smile"></i></button>' : '';
        const forwardedTag = msg.forwarded_from ? '<div class="chat-msg__forwarded"><i class="fas fa-share"></i> Forwarded</div>' : '';

        wrap.innerHTML = senderLine + forwardedTag +
            '<div class="chat-msg__bubble-row">' +
            '<div class="chat-msg__bubble">' + bodyText + attachments + '</div>' + reactBtn + menuBtn +
            '</div>' +
            renderReactionPills(msg.reactions) +
            '<div class="chat-msg__meta"><span class="chat-msg__meta-text">' + time + editedTag + '</span></div>';

        if (prepend) {
            container.insertBefore(wrap, container.firstChild);
        } else {
            container.appendChild(wrap);
        }
    }

    // ---- Edit / delete own messages ------------------------------------------

    function applyDeletedState(el) {
        el.classList.add('is-deleted');
        const bubble = el.querySelector('.chat-msg__bubble');
        if (bubble) bubble.textContent = 'Message deleted';
        const menuBtn = el.querySelector('.chat-msg__menu-btn');
        if (menuBtn) menuBtn.remove();
    }

    function closeOpenMenu() {
        const open = document.querySelector('.chat-msg__menu.open');
        if (open) open.remove();
    }

    function enterEditMode(el, msgId) {
        closeOpenMenu();
        const bubble = el.querySelector('.chat-msg__bubble');
        if (!bubble || el.classList.contains('is-editing')) return;
        el.classList.add('is-editing');
        const originalHtml = bubble.innerHTML;
        const originalText = bubble.textContent;

        bubble.innerHTML =
            '<textarea class="chat-msg__edit-input erp-input">' + escapeHtml(originalText) + '</textarea>' +
            '<div class="chat-msg__edit-actions">' +
            '<button type="button" class="erp-btn erp-btn-secondary chat-msg__edit-cancel">Cancel</button>' +
            '<button type="button" class="erp-btn erp-btn-primary chat-msg__edit-save">Save</button>' +
            '</div>';

        const textarea = bubble.querySelector('.chat-msg__edit-input');
        textarea.focus();
        textarea.setSelectionRange(textarea.value.length, textarea.value.length);

        function cancel() {
            el.classList.remove('is-editing');
            bubble.innerHTML = originalHtml;
        }

        bubble.querySelector('.chat-msg__edit-cancel').addEventListener('click', cancel);
        bubble.querySelector('.chat-msg__edit-save').addEventListener('click', function () {
            const newBody = textarea.value.trim();
            if (!newBody) { alert('Message cannot be empty.'); return; }
            api(messageDetailUrl(state.currentRoomId, msgId), {
                method: 'PATCH',
                body: { body: newBody },
            }).then(function (updated) {
                el.classList.remove('is-editing');
                bubble.innerHTML = highlightMentions(escapeHtml(updated.body), updated.mentions) + (updated.attachments || []).map(renderAttachment).join('');
                const meta = el.querySelector('.chat-msg__meta-text');
                if (meta && !/\(edited\)$/.test(meta.textContent)) meta.textContent += ' (edited)';
            }).catch(function (err) {
                alert((err.data && err.data.detail) || 'Could not edit message.');
                cancel();
            });
        });

        textarea.addEventListener('keydown', function (event) {
            if (event.key === 'Escape') cancel();
        });
    }

    function deleteMessage(el, msgId) {
        closeOpenMenu();
        if (!confirm('Delete this message?')) return;
        api(messageDetailUrl(state.currentRoomId, msgId), {
            method: 'DELETE',
        }).then(function () {
            applyDeletedState(el);
        }).catch(function (err) {
            alert((err.data && err.data.detail) || 'Could not delete message.');
        });
    }

    function sendReaction(msgId, emoji) {
        if (!state.currentRoomId) return;
        api(messageReactionsUrl(state.currentRoomId, msgId), {
            method: 'POST',
            body: { emoji: emoji },
        }).catch(function () { /* ignore */ });
    }

    function applyReactionUpdate(messageId, reactions) {
        const el = document.querySelector('.chat-msg[data-message-id="' + messageId + '"]');
        if (!el) return;
        const container = el.querySelector('.chat-msg__reactions');
        if (container) container.outerHTML = renderReactionPills(reactions);
    }

    document.getElementById('chat-thread-messages').addEventListener('click', function (event) {
        const reactBtn = event.target.closest('.chat-msg__react-btn');
        if (reactBtn && window.ChatEmojiPicker) {
            event.stopPropagation();
            const el = reactBtn.closest('.chat-msg');
            window.ChatEmojiPicker.open(reactBtn, function (ch) {
                sendReaction(el.dataset.messageId, ch);
            });
            return;
        }
        const pill = event.target.closest('.chat-msg__reaction-pill');
        if (pill) {
            const el = pill.closest('.chat-msg');
            sendReaction(el.dataset.messageId, pill.dataset.emoji);
            return;
        }
        const menuBtn = event.target.closest('.chat-msg__menu-btn');
        if (menuBtn) {
            event.stopPropagation();
            const existing = menuBtn.parentElement.querySelector('.chat-msg__menu');
            if (existing) { existing.remove(); return; }
            closeOpenMenu();
            const el = menuBtn.closest('.chat-msg');
            const msgId = el.dataset.messageId;
            const isOwnMsg = el.classList.contains('is-own');
            const isPinned = state.currentRoomDetail && state.currentRoomDetail.pinned_message
                && String(state.currentRoomDetail.pinned_message.id) === String(msgId);
            const canPin = state.currentRoomDetail && state.currentRoomDetail.room_type === 'group' ? canManageGroup : true;
            const menu = document.createElement('div');
            menu.className = 'chat-msg__menu open';
            menu.innerHTML =
                (isOwnMsg ? '<button type="button" class="chat-msg__menu-item" data-action="edit">Edit</button>' : '') +
                (isOwnMsg ? '<button type="button" class="chat-msg__menu-item chat-msg__menu-item--danger" data-action="delete">Delete</button>' : '') +
                '<button type="button" class="chat-msg__menu-item" data-action="forward">Forward</button>' +
                (canPin ? '<button type="button" class="chat-msg__menu-item" data-action="' + (isPinned ? 'unpin' : 'pin') + '">' + (isPinned ? 'Unpin' : 'Pin') + '</button>' : '');
            menuBtn.parentElement.appendChild(menu);
            return;
        }
        const actionBtn = event.target.closest('.chat-msg__menu-item');
        if (actionBtn) {
            const el = actionBtn.closest('.chat-msg');
            const msgId = el.dataset.messageId;
            const action = actionBtn.dataset.action;
            if (action === 'edit') enterEditMode(el, msgId);
            else if (action === 'delete') deleteMessage(el, msgId);
            else if (action === 'pin') pinMessage(msgId);
            else if (action === 'unpin') unpinMessage();
            else if (action === 'forward') openForwardModal(msgId);
            closeOpenMenu();
            return;
        }
        closeOpenMenu();
    });

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
            document.getElementById('chat-mention-suggestions').hidden = true;
            appendMessage(msg, false);
            scrollToBottom();
            loadRoomList();
        }).catch(function (err) {
            alert((err.data && err.data.detail) || 'Could not send message.');
        });
    });

    const emojiBtn = document.getElementById('chat-emoji-btn');
    if (emojiBtn && window.ChatEmojiPicker) {
        emojiBtn.addEventListener('click', function () {
            const input = document.getElementById('chat-message-input');
            window.ChatEmojiPicker.open(emojiBtn, function (ch) {
                const start = input.selectionStart || input.value.length;
                const end = input.selectionEnd || input.value.length;
                input.value = input.value.slice(0, start) + ch + input.value.slice(end);
                input.focus();
                input.setSelectionRange(start + ch.length, start + ch.length);
            });
        });
    }

    document.getElementById('chat-message-input').addEventListener('input', function () {
        sendTyping(true);
        clearTimeout(state.typingTimeout);
        state.typingTimeout = setTimeout(function () { sendTyping(false); }, 2000);

        const input = this;
        const cursor = input.selectionStart;
        const textBeforeCursor = input.value.slice(0, cursor);
        const match = textBeforeCursor.match(/@([\w .]{0,30})$/);
        if (match) {
            renderMentionSuggestions(match[1]);
        } else {
            document.getElementById('chat-mention-suggestions').hidden = true;
        }
    });

    document.getElementById('chat-mention-suggestions').addEventListener('click', function (event) {
        const item = event.target.closest('.chat-mention-suggestions__item');
        if (!item) return;
        const input = document.getElementById('chat-message-input');
        const cursor = input.selectionStart;
        const textBeforeCursor = input.value.slice(0, cursor);
        const match = textBeforeCursor.match(/@([\w .]{0,30})$/);
        if (!match) return;
        const before = input.value.slice(0, match.index);
        const after = input.value.slice(cursor);
        const insertion = '@' + item.dataset.name + ' ';
        input.value = before + insertion + after;
        input.focus();
        const newPos = (before + insertion).length;
        input.setSelectionRange(newPos, newPos);
        this.hidden = true;
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
                // Only counts as "seen" if the tab is actually visible and
                // focused right now — otherwise this room being merely the
                // "current" one in JS state (e.g. a background/minimized tab)
                // would mark messages read that the user never actually saw.
                // If it's not, markReadIfVisible() below catches up once the
                // tab is actually looked at again.
                if (document.visibilityState === 'visible' && document.hasFocus()) {
                    markRead(roomId);
                }
            }
            loadRoomList();
        } else if (payload.event === 'message_edited') {
            const el = document.querySelector('.chat-msg[data-message-id="' + payload.message.id + '"]');
            if (el && !el.classList.contains('is-editing')) {
                const bubble = el.querySelector('.chat-msg__bubble');
                bubble.innerHTML = highlightMentions(escapeHtml(payload.message.body), payload.message.mentions) + (payload.message.attachments || []).map(renderAttachment).join('');
                const meta = el.querySelector('.chat-msg__meta-text');
                if (meta && !/\(edited\)$/.test(meta.textContent)) meta.textContent += ' (edited)';
            }
        } else if (payload.event === 'message_deleted') {
            const el = document.querySelector('.chat-msg[data-message-id="' + payload.message_id + '"]');
            if (el) applyDeletedState(el);
        } else if (payload.event === 'reaction_updated') {
            applyReactionUpdate(payload.message_id, payload.reactions);
        } else if (payload.event === 'read_receipt_updated') {
            updateSeenIndicator(payload.user_id, payload.last_read_message_id);
        } else if (payload.event === 'typing') {
            const label = document.getElementById('chat-thread-typing');
            if (payload.user_id !== currentUserId) {
                label.textContent = payload.is_typing ? payload.username + ' is typing…' : '';
            }
        } else if (payload.event === 'participant_added' || payload.event === 'participant_removed') {
            api(urls.roomDetail.replace('/0/', '/' + roomId + '/')).then(function (room) {
                state.currentRoomDetail = room;
                document.getElementById('chat-thread-name').textContent = room.display_name;
                if (groupSettingsModal.classList.contains('open')) renderGroupMembers(room);
                loadRoomList();
            });
        } else if (payload.event === 'group_updated') {
            state.currentRoomDetail = payload.room;
            document.getElementById('chat-thread-name').textContent = payload.room.display_name;
            loadRoomList();
        } else if (payload.event === 'group_deleted') {
            closeModal(groupSettingsModal);
            document.getElementById('chat-thread-active').hidden = true;
            document.getElementById('chat-thread-empty').hidden = false;
            state.currentRoomId = null;
            window.ChatApp.currentRoomId = null;
            alert('This group was deleted.');
            loadRoomList();
        } else if (payload.event === 'pin_updated') {
            renderPinnedBanner(payload.pinned_message);
            if (state.currentRoomDetail) state.currentRoomDetail.pinned_message = payload.pinned_message;
        } else if (payload.event === 'buzz') {
            triggerBuzzShake();
            if (payload.from_user_id !== currentUserId && window.ChatSound) window.ChatSound.playBuzz();
        }
    }

    // ---- Buzz ---------------------------------------------------------------

    let buzzShakeTimeout = null;
    function triggerBuzzShake() {
        const target = document.getElementById('chat-thread-active');
        if (!target) return;
        target.classList.remove('chat-buzz-shake');
        // Force reflow so re-adding the class restarts the animation on repeat buzzes.
        void target.offsetWidth;
        target.classList.add('chat-buzz-shake');
        clearTimeout(buzzShakeTimeout);
        buzzShakeTimeout = setTimeout(function () { target.classList.remove('chat-buzz-shake'); }, 600);
    }

    // ---- Read receipts ---------------------------------------------------

    function appendSeenLabel(target) {
        const meta = target.querySelector('.chat-msg__meta');
        const seen = document.createElement('span');
        seen.className = 'chat-msg__seen';
        seen.textContent = ' · Seen';
        seen.title = 'Click to see who';
        seen.addEventListener('click', function (event) {
            event.stopPropagation();
            showReadBy(target, target.dataset.messageId);
        });
        meta.appendChild(seen);
    }

    function updateSeenIndicator(readerUserId, lastReadMessageId) {
        if (readerUserId === currentUserId) return; // ignore our own mark-read
        const container = document.getElementById('chat-thread-messages');
        container.querySelectorAll('.chat-msg__seen').forEach(function (el) { el.remove(); });
        if (!lastReadMessageId) return;

        const ownMsgs = Array.from(container.querySelectorAll('.chat-msg.is-own'));
        for (let i = ownMsgs.length - 1; i >= 0; i--) {
            if (parseInt(ownMsgs[i].dataset.messageId, 10) <= lastReadMessageId) {
                appendSeenLabel(ownMsgs[i]);
            }
        }
    }

    function showReadBy(el, msgId) {
        const existing = el.querySelector('.chat-msg__seen-popover');
        if (existing) { existing.remove(); return; }
        const popover = document.createElement('div');
        popover.className = 'chat-msg__seen-popover';
        popover.textContent = 'Loading…';
        el.querySelector('.chat-msg__meta').appendChild(popover);
        api(messageReadByUrl(state.currentRoomId, msgId)).then(function (readers) {
            popover.textContent = readers.length
                ? 'Seen by ' + readers.map(function (r) { return r.display_name; }).join(', ')
                : 'Not seen yet';
        }).catch(function () {
            popover.textContent = 'Could not load read status.';
        });
    }

    // ---- Pinned message ---------------------------------------------------

    function renderPinnedBanner(pinnedMessage) {
        const banner = document.getElementById('chat-thread-pinned');
        if (!pinnedMessage) {
            banner.hidden = true;
            banner.innerHTML = '';
            return;
        }
        const canUnpin = state.currentRoomDetail && state.currentRoomDetail.room_type === 'group' ? canManageGroup : true;
        banner.hidden = false;
        banner.innerHTML =
            '<i class="fas fa-thumbtack"></i>' +
            '<span class="chat-thread__pinned-text">' + escapeHtml((pinnedMessage.body || '[Attachment]').slice(0, 120)) + '</span>' +
            (canUnpin ? '<button type="button" class="chat-thread__pinned-unpin" title="Unpin">&times;</button>' : '');
        banner.dataset.messageId = pinnedMessage.id;
    }

    // ---- Forward message ---------------------------------------------------

    const forwardModal = document.getElementById('chat-forward-modal');
    let forwardMessageId = null;

    function openForwardModal(msgId) {
        forwardMessageId = msgId;
        const select = document.getElementById('chat-forward-room-select');
        select.innerHTML = state.rooms.map(function (r) {
            return '<option value="' + r.id + '">' + escapeHtml(r.display_name) + '</option>';
        }).join('') || '<option value="">No conversations available</option>';
        forwardModal.classList.add('open');
    }

    document.getElementById('chat-forward-send-btn').addEventListener('click', function () {
        const targetRoomId = document.getElementById('chat-forward-room-select').value;
        if (!targetRoomId || !forwardMessageId) return;
        api((urls.messages || '').replace('/0/', '/' + targetRoomId + '/'), {
            method: 'POST',
            body: { forwarded_from_message_id: forwardMessageId },
        }).then(function () {
            closeModal(forwardModal);
            if (String(targetRoomId) === String(state.currentRoomId)) loadMessages(state.currentRoomId);
            loadRoomList();
        }).catch(function (err) {
            alert((err.data && err.data.detail) || 'Could not forward message.');
        });
    });

    function pinMessage(msgId) {
        if (!state.currentRoomId) return;
        api((urls.roomPin || '').replace('/0/', '/' + state.currentRoomId + '/'), {
            method: 'POST',
            body: { message_id: msgId },
        }).catch(function (err) { alert((err.data && err.data.detail) || 'Could not pin message.'); });
    }

    function unpinMessage() {
        if (!state.currentRoomId) return;
        api((urls.roomPin || '').replace('/0/', '/' + state.currentRoomId + '/'), { method: 'DELETE' })
            .catch(function () { /* ignore */ });
    }

    document.getElementById('chat-thread-pinned').addEventListener('click', function (event) {
        if (event.target.closest('.chat-thread__pinned-unpin')) {
            event.stopPropagation();
            if (!state.currentRoomId) return;
            api((urls.roomPin || '').replace('/0/', '/' + state.currentRoomId + '/'), { method: 'DELETE' }).catch(function () { /* ignore */ });
            return;
        }
        const msgId = this.dataset.messageId;
        const el = document.querySelector('.chat-msg[data-message-id="' + msgId + '"]');
        if (el) {
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
            el.classList.add('is-highlighted');
            setTimeout(function () { el.classList.remove('is-highlighted'); }, 1500);
        }
    });

    // ---- Buzz button --------------------------------------------------------

    const buzzBtn = document.getElementById('chat-buzz-btn');
    if (buzzBtn) {
        buzzBtn.addEventListener('click', function () {
            if (!state.currentRoomId || buzzBtn.disabled) return;
            buzzBtn.disabled = true;
            setTimeout(function () { buzzBtn.disabled = false; }, 5000);
            triggerBuzzShake();
            if (window.ChatSound) window.ChatSound.playBuzz(); // immediate feedback for the sender too
            api((urls.roomBuzz || '').replace('/0/', '/' + state.currentRoomId + '/'), { method: 'POST' })
                .catch(function () { /* cooldown or transient failure — button re-enables on its own timer */ });
        });
    }

    // ---- Group settings panel ---------------------------------------------

    const groupSettingsModal = document.getElementById('chat-group-settings-modal');
    let pendingAvatarFile = null;

    function renderGroupMembers(room) {
        const container = document.getElementById('chat-group-settings-members');
        const countEl = document.getElementById('chat-group-settings-member-count');
        const active = (room.participants || []).filter(function (p) { return !p.left_at; });
        countEl.textContent = active.length;
        container.innerHTML = active.map(function (p) {
            const status = getUserStatus(p.user.id);
            const canRemove = canManageGroup && p.user.id !== currentUserId;
            return '<div class="chat-group-settings__member-row" data-user-id="' + p.user.id + '">' +
                '<span class="chat-group-settings__member-dot chat-group-settings__member-dot--' + status + '"></span>' +
                '<span class="chat-group-settings__member-name">' + escapeHtml(p.user.display_name) + (p.user.id === currentUserId ? ' (you)' : '') + '</span>' +
                (p.role === 'admin' ? '<span class="chat-group-settings__member-role">Admin</span>' : '') +
                (canRemove ? '<button type="button" class="chat-group-settings__member-remove" data-user-id="' + p.user.id + '">Remove</button>' : '') +
                '</div>';
        }).join('') || '<div class="chat-empty-note">No members.</div>';

        const addRow = document.getElementById('chat-group-settings-add-row');
        addRow.hidden = !canManageGroup;
        if (canManageGroup) {
            const activeIds = new Set(active.map(function (p) { return p.user.id; }));
            loadChattableUsers().then(function (users) {
                const select = document.getElementById('chat-group-settings-add-select');
                const available = users.filter(function (u) { return !activeIds.has(u.id); });
                select.innerHTML = '<option value="">Add a member…</option>' + available.map(function (u) {
                    return '<option value="' + u.id + '">' + escapeHtml(u.display_name) + '</option>';
                }).join('');
            });
        }
    }

    document.getElementById('chat-group-settings-members').addEventListener('click', function (event) {
        const btn = event.target.closest('.chat-group-settings__member-remove');
        if (!btn || !state.currentRoomId) return;
        const userId = btn.dataset.userId;
        if (!confirm('Remove this member from the group?')) return;
        api((urls.roomRemoveParticipant || '').replace('/0/0/', '/' + state.currentRoomId + '/' + userId + '/'), { method: 'DELETE' })
            .catch(function (err) { alert((err.data && err.data.detail) || 'Could not remove member.'); });
    });

    document.getElementById('chat-group-settings-add-btn').addEventListener('click', function () {
        const select = document.getElementById('chat-group-settings-add-select');
        const userId = select.value;
        if (!userId || !state.currentRoomId) return;
        api((urls.roomAddParticipant || '').replace('/0/', '/' + state.currentRoomId + '/'), {
            method: 'POST',
            body: { user_id: userId },
        }).catch(function (err) { alert((err.data && err.data.detail) || 'Could not add member.'); });
    });

    document.getElementById('chat-room-info-btn').addEventListener('click', function () {
        const room = state.currentRoomDetail;
        if (!room || room.room_type !== 'group') return;
        pendingAvatarFile = null;
        document.getElementById('chat-group-settings-name').value = room.name || '';
        document.getElementById('chat-group-settings-description').value = room.description || '';
        document.getElementById('chat-group-settings-name').disabled = !canManageGroup;
        document.getElementById('chat-group-settings-description').disabled = !canManageGroup;
        document.getElementById('chat-group-settings-avatar-btn').hidden = !canManageGroup;
        document.getElementById('chat-group-settings-save-btn').hidden = !canManageGroup;
        const preview = document.getElementById('chat-group-settings-avatar-preview');
        preview.style.backgroundImage = room.avatar_url ? 'url(' + room.avatar_url + ')' : '';
        preview.textContent = room.avatar_url ? '' : initials(room.name);
        document.getElementById('chat-group-settings-danger').hidden = !isSuperuser;
        renderGroupMembers(room);
        groupSettingsModal.classList.add('open');
    });

    document.getElementById('chat-group-settings-avatar-btn').addEventListener('click', function () {
        document.getElementById('chat-group-settings-avatar-input').click();
    });
    document.getElementById('chat-group-settings-avatar-input').addEventListener('change', function () {
        const file = this.files[0];
        if (!file) return;
        pendingAvatarFile = file;
        const preview = document.getElementById('chat-group-settings-avatar-preview');
        preview.style.backgroundImage = 'url(' + URL.createObjectURL(file) + ')';
        preview.textContent = '';
    });

    document.getElementById('chat-group-settings-save-btn').addEventListener('click', function () {
        if (!state.currentRoomId) return;
        const formData = new FormData();
        formData.append('name', document.getElementById('chat-group-settings-name').value.trim());
        formData.append('description', document.getElementById('chat-group-settings-description').value.trim());
        if (pendingAvatarFile) formData.append('avatar', pendingAvatarFile);

        api((urls.roomUpdateSettings || '').replace('/0/', '/' + state.currentRoomId + '/'), {
            method: 'PATCH',
            body: formData,
        }).then(function (room) {
            state.currentRoomDetail = room;
            document.getElementById('chat-thread-name').textContent = room.display_name;
            closeModal(groupSettingsModal);
            loadRoomList();
        }).catch(function (err) {
            alert((err.data && err.data.detail) || 'Could not update group.');
        });
    });

    document.getElementById('chat-group-settings-delete-btn').addEventListener('click', function () {
        if (!state.currentRoomId) return;
        if (!confirm('Delete this group? This hides it for everyone; message history is preserved.')) return;
        api(urls.roomDetail.replace('/0/', '/' + state.currentRoomId + '/'), { method: 'DELETE' })
            .then(function () {
                closeModal(groupSettingsModal);
            }).catch(function (err) {
                alert((err.data && err.data.detail) || 'Could not delete group.');
            });
    });

    function connectPresenceSocket() {
        const socket = new WebSocket(wsUrl('/ws/chat/presence/'));
        socket.onmessage = function (event) {
            const payload = JSON.parse(event.data);
            if (payload.event === 'unread_count_changed' || payload.event === 'new_message') {
                if (payload.event === 'new_message' && payload.room_id !== state.currentRoomId && window.ChatSound) {
                    window.ChatSound.playMessageDing();
                }
                loadRoomList();
            } else if (payload.event === 'incoming_call') {
                window.ChatApp.onIncomingCall(payload);
            } else if (payload.event === 'presence_changed') {
                if (payload.status === 'offline') {
                    delete state.userStatuses[payload.user_id];
                } else {
                    state.userStatuses[payload.user_id] = payload.status;
                }
                renderRoomList();
                updatePresenceLabel();
                renderOnlineUsersPanel();
            } else if (payload.event === 'group_deleted') {
                if (payload.room_id === state.currentRoomId) {
                    handleRoomEvent(payload.room_id, payload);
                } else {
                    loadRoomList();
                }
            }
        };
        socket.onclose = function () {
            setTimeout(connectPresenceSocket, 4000);
        };
        state.presenceSocket = socket;
    }

    function loadOnlineSnapshot() {
        if (!urls.presenceOnline) return;
        api(urls.presenceOnline).then(function (data) {
            state.userStatuses = {};
            const statuses = data.statuses || {};
            Object.keys(statuses).forEach(function (uid) { state.userStatuses[uid] = statuses[uid]; });
            renderRoomList();
            updatePresenceLabel();
            renderOnlineUsersPanel();
        }).catch(function () { /* ignore */ });
    }

    // ---- Activity ping (drives online vs away) ------------------------------
    // A connection alone only proves "not offline" — real activity (mouse,
    // keyboard, or the tab becoming visible again) is what promotes a user
    // from away back to online. Throttled so it doesn't spam the socket.

    let lastActivityPingAt = 0;
    function pingActivity() {
        const now = Date.now();
        if (now - lastActivityPingAt < 15000) return;
        lastActivityPingAt = now;
        if (state.presenceSocket && state.presenceSocket.readyState === WebSocket.OPEN) {
            state.presenceSocket.send(JSON.stringify({ event: 'activity' }));
        }
    }
    ['mousemove', 'keydown', 'click', 'touchstart'].forEach(function (evt) {
        document.addEventListener(evt, pingActivity, { passive: true });
    });
    function markReadIfVisible() {
        if (document.visibilityState === 'visible' && document.hasFocus() && state.currentRoomId) {
            markRead(state.currentRoomId);
        }
    }
    document.addEventListener('visibilitychange', function () {
        if (document.visibilityState === 'visible') { pingActivity(); markReadIfVisible(); }
    });
    window.addEventListener('focus', markReadIfVisible);

    // ---- Online users panel ------------------------------------------------

    const onlineToggleBtn = document.getElementById('chat-online-toggle-btn');
    const onlinePanel = document.getElementById('chat-online-panel');
    const onlinePanelItems = document.getElementById('chat-online-panel-items');
    const STATUS_LABELS = { online: 'Online', away: 'Away', offline: 'Offline' };
    const STATUS_ORDER = { online: 0, away: 1, offline: 2 };

    function renderOnlineUsersPanel() {
        if (!onlinePanelItems) return;
        if (!chattableUsers) {
            onlinePanelItems.innerHTML = '<div class="chat-empty-note">Loading…</div>';
            return;
        }
        const sorted = chattableUsers.slice().sort(function (a, b) {
            return STATUS_ORDER[getUserStatus(a.id)] - STATUS_ORDER[getUserStatus(b.id)];
        });
        onlinePanelItems.innerHTML = sorted.length
            ? sorted.map(function (u) {
                const status = getUserStatus(u.id);
                return '<button type="button" class="chat-online-panel__row" data-user-id="' + u.id + '">' +
                    '<span class="chat-online-panel__dot chat-online-panel__dot--' + status + '"></span>' +
                    '<span class="chat-online-panel__name">' + escapeHtml(u.display_name) + '</span>' +
                    '<span class="chat-online-panel__status">' + STATUS_LABELS[status] + '</span>' +
                    '</button>';
            }).join('')
            : '<div class="chat-empty-note">No other chat users found.</div>';
    }

    if (onlineToggleBtn && onlinePanel) {
        onlineToggleBtn.addEventListener('click', function () {
            const isHidden = onlinePanel.hidden;
            onlinePanel.hidden = !isHidden;
            if (isHidden) {
                loadChattableUsers().then(renderOnlineUsersPanel);
            }
        });
    }

    if (onlinePanelItems) {
        onlinePanelItems.addEventListener('click', function (event) {
            const row = event.target.closest('.chat-online-panel__row');
            if (!row) return;
            const userId = row.dataset.userId;
            api(urls.roomList, { method: 'POST', body: { room_type: 'dm', user_id: userId } })
                .then(function (room) {
                    onlinePanel.hidden = true;
                    loadRoomList();
                    openRoom(room.id);
                }).catch(function (err) { alert((err.data && err.data.detail) || 'Could not start chat.'); });
        });
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
        isUserOnline: function (userId) { return getUserStatus(parseInt(userId, 10)) !== 'offline'; },
        getUserStatus: function (userId) { return getUserStatus(parseInt(userId, 10)); },
        getCurrentRoomDetail: function () { return state.currentRoomDetail; },
    };

    loadRoomList();
    loadOnlineSnapshot();
    loadChattableUsers().then(renderOnlineUsersPanel);
    connectPresenceSocket();

    // Deep link from a toast notification, e.g. /chat/?room=12
    const deepLinkRoomId = parseInt(new URLSearchParams(window.location.search).get('room'), 10);
    if (deepLinkRoomId) {
        api(urls.roomDetail.replace('/0/', '/' + deepLinkRoomId + '/'))
            .then(function () { openRoom(deepLinkRoomId); })
            .catch(function () { /* room no longer accessible — ignore */ });
    }
})();
