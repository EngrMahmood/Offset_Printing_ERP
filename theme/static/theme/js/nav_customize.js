// Lets each user choose which top-nav modules stay inline vs. get tucked
// into a "More" dropdown, and reorder both — no code change needed. Layout
// is stored per-browser in localStorage (not synced across devices/browsers
// yet, but nothing here would block adding server-side sync later).
(function () {
    var STORAGE_KEY = 'erp_nav_layout_v1';

    var modulesEl = document.getElementById('erp-topnav-modules');
    var moreWrap = document.getElementById('erp-topnav-more');
    var moreBtn = document.getElementById('erp-topnav-more-btn');
    var moreMenu = document.getElementById('erp-topnav-more-menu');
    var customizeBtn = document.getElementById('erp-nav-customize-btn');
    var modal = document.getElementById('erp-nav-customize-modal');
    var list = document.getElementById('erp-nav-customize-list');
    var saveBtn = document.getElementById('erp-nav-customize-save');
    var resetBtn = document.getElementById('erp-nav-customize-reset');

    if (!modulesEl) return;

    // Catalog: every nav link that existed in the DOM on page load, keyed by
    // its stable data-nav-key, in the server-rendered (permission-filtered,
    // default) order. This is the single source of truth for "what exists
    // right now" — customization only ever reorders/relocates these nodes,
    // never invents or drops one.
    var catalog = Array.prototype.slice.call(
        modulesEl.querySelectorAll('a.erp-topnav-module[data-nav-key]')
    );
    var nodesByKey = {};
    var defaultOrder = [];
    catalog.forEach(function (node) {
        var key = node.getAttribute('data-nav-key');
        nodesByKey[key] = node;
        defaultOrder.push(key);
    });

    function loadLayout() {
        var pinned = [];
        var overflow = [];
        try {
            var raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
            if (raw && Array.isArray(raw.pinned) && Array.isArray(raw.overflow)) {
                pinned = raw.pinned.filter(function (k) { return nodesByKey[k]; });
                overflow = raw.overflow.filter(function (k) { return nodesByKey[k]; });
            }
        } catch (e) { /* fall through to default */ }

        // New modules (permission granted since last save, or first ever
        // load) default to pinned so they're never silently unreachable.
        var known = {};
        pinned.concat(overflow).forEach(function (k) { known[k] = true; });
        defaultOrder.forEach(function (k) {
            if (!known[k]) pinned.push(k);
        });

        if (pinned.length === 0 && overflow.length === 0) {
            pinned = defaultOrder.slice();
        }
        return { pinned: pinned, overflow: overflow };
    }

    function saveLayout(layout) {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(layout));
        } catch (e) { /* ignore (private browsing / quota) */ }
    }

    function applyLayout(layout) {
        layout.pinned.forEach(function (key) {
            modulesEl.appendChild(nodesByKey[key]);
        });
        layout.overflow.forEach(function (key) {
            moreMenu.appendChild(nodesByKey[key]);
        });
        moreBtn.style.display = layout.overflow.length ? '' : 'none';
    }

    var currentLayout = loadLayout();
    applyLayout(currentLayout);

    // ── "More" dropdown open/close ──────────────────────────────────────
    if (moreBtn) {
        moreBtn.addEventListener('click', function (event) {
            event.stopPropagation();
            moreMenu.classList.toggle('is-open');
        });
        document.addEventListener('click', function (event) {
            if (!moreWrap.contains(event.target)) {
                moreMenu.classList.remove('is-open');
            }
        });
    }

    // ── Customize modal ──────────────────────────────────────────────────
    function labelFor(key) {
        var node = nodesByKey[key];
        return node ? (node.getAttribute('data-nav-label') || node.textContent.trim()) : key;
    }

    function buildListRow(key) {
        var li = document.createElement('li');
        li.className = 'erp-nav-customize-item';
        li.draggable = true;
        li.setAttribute('data-key', key);
        li.innerHTML = '<i class="fas fa-grip-vertical erp-nav-customize-handle"></i><span>' + labelFor(key) + '</span>';
        return li;
    }

    function buildDivider() {
        var li = document.createElement('li');
        li.className = 'erp-nav-customize-divider';
        li.textContent = 'In "More" menu below';
        return li;
    }

    function openCustomizeModal() {
        if (!modal || !list) return;
        list.innerHTML = '';
        currentLayout.pinned.forEach(function (key) {
            list.appendChild(buildListRow(key));
        });
        list.appendChild(buildDivider());
        currentLayout.overflow.forEach(function (key) {
            list.appendChild(buildListRow(key));
        });
        modal.classList.add('open');
    }

    if (customizeBtn) {
        customizeBtn.addEventListener('click', openCustomizeModal);
    }

    // ── Drag-and-drop reordering within the single list (divider included
    // as a drop-target boundary between the two sections) ───────────────
    var dragEl = null;

    if (list) {
        list.addEventListener('dragstart', function (event) {
            var item = event.target.closest('.erp-nav-customize-item');
            if (!item) return;
            dragEl = item;
            item.classList.add('is-dragging');
            event.dataTransfer.effectAllowed = 'move';
            event.dataTransfer.setData('text/plain', item.getAttribute('data-key'));
        });

        list.addEventListener('dragend', function () {
            if (dragEl) dragEl.classList.remove('is-dragging');
            dragEl = null;
        });

        list.addEventListener('dragover', function (event) {
            if (!dragEl) return;
            event.preventDefault();
            var target = event.target.closest('.erp-nav-customize-item, .erp-nav-customize-divider');
            if (!target || target === dragEl) return;
            var rect = target.getBoundingClientRect();
            var before = (event.clientY - rect.top) < rect.height / 2;
            list.insertBefore(dragEl, before ? target : target.nextSibling);
        });
    }

    if (saveBtn) {
        saveBtn.addEventListener('click', function () {
            var pinned = [];
            var overflow = [];
            var seenDivider = false;
            Array.prototype.forEach.call(list.children, function (child) {
                if (child.classList.contains('erp-nav-customize-divider')) {
                    seenDivider = true;
                    return;
                }
                var key = child.getAttribute('data-key');
                (seenDivider ? overflow : pinned).push(key);
            });
            currentLayout = { pinned: pinned, overflow: overflow };
            saveLayout(currentLayout);
            applyLayout(currentLayout);
            modal.classList.remove('open');
        });
    }

    if (resetBtn) {
        resetBtn.addEventListener('click', function () {
            currentLayout = { pinned: defaultOrder.slice(), overflow: [] };
            saveLayout(currentLayout);
            applyLayout(currentLayout);
            modal.classList.remove('open');
        });
    }
})();
