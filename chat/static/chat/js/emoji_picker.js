// Shared, dependency-free emoji picker popover used by the compose box and
// message reactions. Static embedded emoji list — no CDN/external fetch,
// since this is an offline LAN deployment. Renders via the OS/browser's own
// emoji font, no image assets needed.
(function () {
    const CATEGORIES = [
        {
            name: 'Smileys',
            emoji: ['😀', '😃', '😄', '😁', '😆', '😅', '🤣', '😂', '🙂', '🙃', '😉', '😊', '😇', '🥰', '😍',
                '🤩', '😘', '😋', '😛', '😜', '🤪', '🤑', '🤗', '🤭', '🤫', '🤔', '🤐', '😐', '😑', '😶',
                '😏', '😒', '🙄', '😬', '😌', '😴', '🤒', '🤕', '🤢', '🥵', '🥶', '😵', '🤯', '🥳', '😎',
                '🤓', '🧐', '😢', '😭', '😡', '😱', '😨'],
        },
        {
            name: 'Gestures',
            emoji: ['👍', '👎', '👌', '✌️', '🤞', '🤟', '🤘', '🤙', '👈', '👉', '👆', '👇', '☝️', '👋', '🤚',
                '✋', '👏', '🙌', '👐', '🙏', '✊', '👊', '💪', '🖕'],
        },
        {
            name: 'Hearts',
            emoji: ['❤️', '🧡', '💛', '💚', '💙', '💜', '🖤', '🤍', '🤎', '💔', '❣️', '💕', '💞', '💓', '💗', '💖'],
        },
        {
            name: 'Objects',
            emoji: ['🔥', '🎉', '🎊', '✨', '⭐', '💯', '✅', '❌', '⚠️', '📎', '📌', '📅', '⏰', '📷', '🎥',
                '📞', '💬', '✉️', '📦', '🔔', '💡', '🔧', '🖨️'],
        },
    ];

    let openPopover = null;
    let outsideClickHandler = null;

    function close() {
        if (openPopover) {
            openPopover.remove();
            openPopover = null;
        }
        if (outsideClickHandler) {
            document.removeEventListener('mousedown', outsideClickHandler);
            outsideClickHandler = null;
        }
    }

    function open(anchorEl, onSelect) {
        if (openPopover) { close(); return; }

        const popover = document.createElement('div');
        popover.className = 'chat-emoji-picker';

        const tabs = document.createElement('div');
        tabs.className = 'chat-emoji-picker__tabs';
        const grid = document.createElement('div');
        grid.className = 'chat-emoji-picker__grid';

        function renderCategory(index) {
            grid.innerHTML = '';
            CATEGORIES[index].emoji.forEach(function (ch) {
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'chat-emoji-picker__item';
                btn.textContent = ch;
                btn.addEventListener('click', function () {
                    onSelect(ch);
                    close();
                });
                grid.appendChild(btn);
            });
        }

        CATEGORIES.forEach(function (cat, index) {
            const tab = document.createElement('button');
            tab.type = 'button';
            tab.className = 'chat-emoji-picker__tab' + (index === 0 ? ' is-active' : '');
            tab.textContent = cat.name;
            tab.addEventListener('click', function () {
                tabs.querySelectorAll('.chat-emoji-picker__tab').forEach(function (t) { t.classList.remove('is-active'); });
                tab.classList.add('is-active');
                renderCategory(index);
            });
            tabs.appendChild(tab);
        });

        popover.appendChild(tabs);
        popover.appendChild(grid);
        document.body.appendChild(popover);
        renderCategory(0);

        const rect = anchorEl.getBoundingClientRect();
        const popoverRect = popover.getBoundingClientRect();
        let top = rect.top - popoverRect.height - 6;
        if (top < 8) top = rect.bottom + 6;
        let left = rect.left;
        if (left + popoverRect.width > window.innerWidth - 8) left = window.innerWidth - popoverRect.width - 8;
        popover.style.top = Math.max(8, top) + 'px';
        popover.style.left = Math.max(8, left) + 'px';

        openPopover = popover;
        outsideClickHandler = function (event) {
            if (!popover.contains(event.target) && event.target !== anchorEl) close();
        };
        setTimeout(function () { document.addEventListener('mousedown', outsideClickHandler); }, 0);
    }

    window.ChatEmojiPicker = { open: open, close: close };
})();
