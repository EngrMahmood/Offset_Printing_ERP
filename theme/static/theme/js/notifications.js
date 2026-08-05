(function () {
    const root = document.getElementById('erp-notify-root');
    if (!root) return;

    const listUrl = root.dataset.listUrl;
    const markAllUrl = root.dataset.markAllUrl;
    const markReadUrlTemplate = root.dataset.markReadUrlTemplate;
    const csrfToken = (
        document.querySelector('[name=csrfmiddlewaretoken]') || {}
    ).value || getCookie('csrftoken');

    const btn = document.getElementById('erp-notify-btn');
    const badge = document.getElementById('erp-notify-badge');
    const panel = document.getElementById('erp-notify-panel');
    const listEl = document.getElementById('erp-notify-list');
    const markAllBtn = document.getElementById('erp-notify-mark-all');
    const toastStack = document.getElementById('erp-toast-stack');

    let latestId = 0;
    let open = false;
    const toastedIds = loadToastedIds();

    function getCookie(name) {
        const match = document.cookie.match(new RegExp('(^| )' + name + '=([^;]+)'));
        return match ? match[2] : '';
    }

    function loadToastedIds() {
        try {
            return new Set(JSON.parse(sessionStorage.getItem('erp_toasted_ids') || '[]'));
        } catch (e) {
            return new Set();
        }
    }

    function saveToastedIds() {
        try {
            sessionStorage.setItem('erp_toasted_ids', JSON.stringify(Array.from(toastedIds).slice(-100)));
        } catch (e) { /* ignore */ }
    }

    function setBadge(count) {
        const value = Number(count || 0);
        if (!badge) return;
        if (value > 0) {
            badge.textContent = value > 99 ? '99+' : String(value);
            badge.classList.add('is-visible');
        } else {
            badge.textContent = '';
            badge.classList.remove('is-visible');
        }
    }

    function renderList(items) {
        if (!listEl) return;
        if (!items.length) {
            listEl.innerHTML = '<div class="erp-notify-empty">No notifications yet.</div>';
            return;
        }
        listEl.innerHTML = items.map(function (item) {
            const unreadClass = item.is_read ? '' : ' is-unread';
            const message = item.message
                ? '<p class="erp-notify-item__message">' + escapeHtml(item.message) + '</p>'
                : '';
            return (
                '<a class="erp-notify-item' + unreadClass + '" href="' + escapeAttr(item.link || '#') + '" data-id="' + item.id + '">' +
                '<p class="erp-notify-item__title">' + escapeHtml(item.title) + '</p>' +
                message +
                '<div class="erp-notify-item__time">' + escapeHtml(item.created_at_display || '') + '</div>' +
                '</a>'
            );
        }).join('');
    }

    function escapeHtml(value) {
        return String(value || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function escapeAttr(value) {
        return escapeHtml(value).replace(/'/g, '&#39;');
    }

    function showToast(item) {
        if (!toastStack || toastedIds.has(item.id)) return;
        toastedIds.add(item.id);
        saveToastedIds();

        const toast = document.createElement('div');
        toast.className = 'erp-toast';
        toast.innerHTML =
            '<p class="erp-toast__title">' + escapeHtml(item.title) + '</p>' +
            (item.message ? '<p class="erp-toast__message">' + escapeHtml(item.message) + '</p>' : '') +
            '<div class="erp-toast__actions">' +
            (item.link ? '<a class="erp-toast__link" href="' + escapeAttr(item.link) + '">Open</a>' : '') +
            '<button type="button" class="erp-toast__close">Dismiss</button>' +
            '</div>';

        const close = function () {
            toast.remove();
        };
        toast.querySelector('.erp-toast__close').addEventListener('click', close);
        const link = toast.querySelector('.erp-toast__link');
        if (link) {
            link.addEventListener('click', function () {
                markRead(item.id);
            });
        }
        toastStack.appendChild(toast);
        window.setTimeout(close, 8000);
    }

    function markRead(id) {
        if (!id || !markReadUrlTemplate) return Promise.resolve();
        const url = markReadUrlTemplate.replace('/0/', '/' + id + '/');
        return fetch(url, {
            method: 'POST',
            headers: {
                'X-CSRFToken': csrfToken,
                'X-Requested-With': 'XMLHttpRequest',
            },
        }).then(function (response) { return response.json(); })
            .then(function (data) {
                if (data && data.ok) setBadge(data.unread_count);
            }).catch(function () { /* ignore */ });
    }

    function markAllRead() {
        return fetch(markAllUrl, {
            method: 'POST',
            headers: {
                'X-CSRFToken': csrfToken,
                'X-Requested-With': 'XMLHttpRequest',
            },
        }).then(function (response) { return response.json(); })
            .then(function (data) {
                if (data && data.ok) {
                    setBadge(0);
                    refreshList();
                }
            }).catch(function () { /* ignore */ });
    }

    function refreshList() {
        return fetch(listUrl + '?limit=20', {
            headers: { 'X-Requested-With': 'XMLHttpRequest' },
        }).then(function (response) { return response.json(); })
            .then(function (data) {
                if (!data || !data.ok) return;
                setBadge(data.unread_count);
                renderList(data.items || []);
                (data.items || []).forEach(function (item) {
                    if (item.id > latestId) latestId = item.id;
                });
            }).catch(function () { /* ignore */ });
    }

    function pollNew() {
        const url = listUrl + '?limit=10&unread=1' + (latestId ? ('&since_id=' + latestId) : '');
        return fetch(url, {
            headers: { 'X-Requested-With': 'XMLHttpRequest' },
        }).then(function (response) { return response.json(); })
            .then(function (data) {
                if (!data || !data.ok) return;
                setBadge(data.unread_count);
                const items = data.items || [];
                items.slice().reverse().forEach(function (item) {
                    if (item.id > latestId) {
                        latestId = item.id;
                        showToast(item);
                    }
                });
                if (open) refreshList();
            }).catch(function () { /* ignore */ });
    }

    function positionPanel() {
        // Panel is position:fixed; anchor it under the bell button but clamp
        // so it can never run off either edge of the viewport, regardless of
        // where the button sits in the header (it's rarely at the true
        // right edge — username/logout sit past it on narrow screens).
        const rect = btn.getBoundingClientRect();
        const width = Math.min(360, window.innerWidth * 0.92);
        let right = window.innerWidth - rect.right;
        const left = window.innerWidth - right - width;
        if (left < 8) right = window.innerWidth - width - 8;
        if (right < 8) right = 8;
        panel.style.right = right + 'px';
        panel.style.top = (rect.bottom + 8) + 'px';
    }

    if (btn && panel) {
        btn.addEventListener('click', function (event) {
            event.stopPropagation();
            open = !open;
            if (open) positionPanel();
            panel.classList.toggle('is-open', open);
            if (open) refreshList();
        });
        document.addEventListener('click', function (event) {
            if (!root.contains(event.target)) {
                open = false;
                panel.classList.remove('is-open');
            }
        });
        window.addEventListener('resize', function () {
            if (open) positionPanel();
        });
    }

    if (listEl) {
        listEl.addEventListener('click', function (event) {
            const item = event.target.closest('.erp-notify-item');
            if (!item) return;
            const id = item.getAttribute('data-id');
            markRead(id);
        });
    }

    if (markAllBtn) {
        markAllBtn.addEventListener('click', function (event) {
            event.preventDefault();
            markAllRead();
        });
    }

    refreshList().then(function () {
        // Toast only unread items not yet shown this session.
        fetch(listUrl + '?limit=5&unread=1', {
            headers: { 'X-Requested-With': 'XMLHttpRequest' },
        }).then(function (response) { return response.json(); })
            .then(function (data) {
                if (!data || !data.ok) return;
                (data.items || []).slice().reverse().forEach(function (item) {
                    if (item.id > latestId) latestId = item.id;
                    showToast(item);
                });
            }).catch(function () { /* ignore */ });
    });

    window.setInterval(pollNew, 45000);
})();
