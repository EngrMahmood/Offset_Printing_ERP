// Lets each user choose which top-nav modules go in Row 1, Row 2, or a
// "More" dropdown, and reorder within each — no code change needed. Layout
// is saved server-side (UserProfile.nav_layout) so it's the same on every
// device/browser the user logs into, not just the one it was set on.
(function () {
    var LOCAL_FALLBACK_KEY = 'erp_nav_layout_v2'; // bumped: v1 was the old 2-bucket {pinned,overflow} shape

    var modulesEl = document.getElementById('erp-topnav-modules');
    var row2Wrap = document.getElementById('erp-topnav-row2');
    var row2El = document.getElementById('erp-topnav-modules-row2');
    var moreBtn = document.getElementById('erp-topnav-more-btn');
    var moreMenu = document.getElementById('erp-topnav-more-menu');
    var customizeBtn = document.getElementById('erp-nav-customize-btn');
    var modal = document.getElementById('erp-nav-customize-modal');
    var list = document.getElementById('erp-nav-customize-list');
    var saveBtn = document.getElementById('erp-nav-customize-save');
    var resetBtn = document.getElementById('erp-nav-customize-reset');

    if (!modulesEl) return;

    // Every nav item that exists right now — either a plain link
    // (a.erp-topnav-module) or a grouped dropdown button (.erp-topnav-group,
    // e.g. "Offset"/"Admin Tools") — is pinnable as a single unit, keyed by
    // its stable data-nav-key.
    var CATALOG_SELECTOR = 'a.erp-topnav-module[data-nav-key], .erp-topnav-group[data-nav-key]';

    // Catalog: every nav item that exists right now, keyed by its stable
    // data-nav-key. Read from ALL THREE containers an item could be in
    // (row1, row2, More) — not just one. navbar_autofit.js's own
    // auto-overflow pass is scheduled via setTimeout(fn, 0), which on a
    // slow-enough network (this script's own <script src> fetch taking a
    // few ms longer than usual) can fire and complete BEFORE this script
    // even runs, having already moved several modules out of row1. Reading
    // only one container made those permanently invisible to this catalog —
    // and since a Save only ever persists what the catalog captured, that
    // loss compounded on every subsequent save (previously reported: 21 ->
    // 16 -> 8 modules over repeated saves, back when there were only two
    // containers to read). Reading all three keeps the catalog complete
    // regardless of which container auto-overflow had already sorted each
    // item into.
    var catalog = Array.prototype.slice.call(
        modulesEl.querySelectorAll(CATALOG_SELECTOR)
    ).concat(Array.prototype.slice.call(
        row2El ? row2El.querySelectorAll(CATALOG_SELECTOR) : []
    )).concat(Array.prototype.slice.call(
        moreMenu ? moreMenu.querySelectorAll(CATALOG_SELECTOR) : []
    ));
    var nodesByKey = {};
    var defaultOrder = [];
    catalog.forEach(function (node) {
        var key = node.getAttribute('data-nav-key');
        if (nodesByKey[key]) return; // shouldn't happen, but never double-count
        nodesByKey[key] = node;
        defaultOrder.push(key);
    });

    function getCookie(name) {
        var match = document.cookie.match(new RegExp('(^| )' + name + '=([^;]+)'));
        return match ? match[2] : '';
    }
    var csrfToken = (document.querySelector('[name=csrfmiddlewaretoken]') || {}).value || getCookie('csrftoken');

    // Normalizes any raw layout-ish object (server JSON, localStorage, or a
    // legacy {pinned, overflow} shape) into {row1, row2, overflow} arrays,
    // defaulting missing fields to empty so callers never have to null-check.
    function normalizeRaw(raw) {
        if (!raw || typeof raw !== 'object') return null;
        var row1 = raw.row1 || raw.pinned; // "pinned" is the pre-two-row key name
        var row2 = raw.row2 || [];
        var overflow = raw.overflow;
        if (!Array.isArray(row1) || !Array.isArray(row2) || !Array.isArray(overflow)) return null;
        return { row1: row1, row2: row2, overflow: overflow };
    }

    function readServerLayout() {
        var el = document.getElementById('erp-nav-layout-data');
        if (!el) return null;
        try {
            return normalizeRaw(JSON.parse(el.textContent || 'null'));
        } catch (e) { /* fall through */ }
        return null;
    }

    function readLocalFallback() {
        try {
            return normalizeRaw(JSON.parse(localStorage.getItem(LOCAL_FALLBACK_KEY) || 'null'));
        } catch (e) { /* ignore */ }
        return null;
    }

    // Guarantees every catalog key appears exactly once across row1+row2+
    // overflow — regardless of what a saved layout actually contains. Called
    // on every load, every modal open, and every save, so a layout that
    // somehow lost (or duplicated) an entry heals itself the next time it
    // passes through here instead of that module silently staying
    // unreachable/duplicated.
    function reconcile(rawRow1, rawRow2, rawOverflow) {
        var seen = {};
        var row1 = [];
        var row2 = [];
        var overflow = [];
        (rawRow1 || []).forEach(function (k) {
            if (nodesByKey[k] && !seen[k]) { seen[k] = true; row1.push(k); }
        });
        (rawRow2 || []).forEach(function (k) {
            if (nodesByKey[k] && !seen[k]) { seen[k] = true; row2.push(k); }
        });
        (rawOverflow || []).forEach(function (k) {
            if (nodesByKey[k] && !seen[k]) { seen[k] = true; overflow.push(k); }
        });
        // New modules (permission granted since last save, first ever load,
        // or recovering from a layout that had dropped one) default to row1
        // so they're never silently unreachable.
        defaultOrder.forEach(function (k) {
            if (!seen[k]) { seen[k] = true; row1.push(k); }
        });
        return { row1: row1, row2: row2, overflow: overflow };
    }

    function loadLayout() {
        var raw = readServerLayout() || readLocalFallback();
        if (!raw) return { row1: defaultOrder.slice(), row2: [], overflow: [] };
        return reconcile(raw.row1, raw.row2, raw.overflow);
    }

    function saveLayout(layout) {
        var saveUrl = customizeBtn && customizeBtn.getAttribute('data-save-url');
        if (!saveUrl) return;
        fetch(saveUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'X-CSRFToken': csrfToken },
            body: JSON.stringify(layout),
            credentials: 'same-origin',
        }).catch(function () {
            // Offline/network failure: keep the change working locally on this
            // device at least, even though it won't have synced to the account.
            try { localStorage.setItem(LOCAL_FALLBACK_KEY, JSON.stringify(layout)); } catch (e) { /* ignore */ }
        });
    }

    function applyLayout(layout) {
        // navbar_autofit.js (loaded first) may have already auto-overflowed
        // some of these same node references if they didn't fit at the
        // current width — always clear that tag here so the manual layout
        // just applied is authoritative; the next fit() pass recomputes
        // auto-overflow fresh against it rather than trusting a stale tag
        // from before this layout existed.
        layout.row1.forEach(function (key) {
            nodesByKey[key].removeAttribute('data-auto-origin');
            modulesEl.appendChild(nodesByKey[key]);
        });
        layout.row2.forEach(function (key) {
            nodesByKey[key].removeAttribute('data-auto-origin');
            row2El.appendChild(nodesByKey[key]);
        });
        layout.overflow.forEach(function (key) {
            nodesByKey[key].removeAttribute('data-auto-origin');
            moreMenu.appendChild(nodesByKey[key]);
        });
        moreBtn.style.display = layout.overflow.length ? '' : 'none';
        if (row2Wrap) row2Wrap.classList.toggle('has-items', layout.row2.length > 0);
    }

    var currentLayout = loadLayout();
    applyLayout(currentLayout);

    // "More" dropdown open/close, and the same for grouped dropdowns
    // (Offset, Admin Tools), is wired centrally in nav_groups.js.

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

    function buildDivider(text, sectionAttr) {
        var li = document.createElement('li');
        li.className = 'erp-nav-customize-divider';
        li.setAttribute('data-section-end', sectionAttr);
        li.textContent = text;
        return li;
    }

    function openCustomizeModal() {
        if (!modal || !list) return;
        // Reconcile before rendering: if currentLayout ever lost an entry
        // (e.g. a layout saved before this safety net existed), this is
        // where it gets a chance to come back instead of just never
        // appearing in the list a user is trying to fix things from.
        currentLayout = reconcile(currentLayout.row1, currentLayout.row2, currentLayout.overflow);
        list.innerHTML = '';
        currentLayout.row1.forEach(function (key) {
            list.appendChild(buildListRow(key));
        });
        list.appendChild(buildDivider('Row 2 below', 'row1'));
        currentLayout.row2.forEach(function (key) {
            list.appendChild(buildListRow(key));
        });
        list.appendChild(buildDivider('In "More" menu below', 'row2'));
        currentLayout.overflow.forEach(function (key) {
            list.appendChild(buildListRow(key));
        });
        var countEl = document.getElementById('erp-nav-customize-count');
        if (countEl) {
            var total = currentLayout.row1.length + currentLayout.row2.length + currentLayout.overflow.length;
            countEl.textContent = total + ' modules total — scroll within the list to see them all.';
        }
        modal.classList.add('open');
    }

    if (customizeBtn) {
        customizeBtn.addEventListener('click', openCustomizeModal);
    }

    // ── Drag-and-drop reordering within the single list (both dividers
    // included as drop-target boundaries between the three sections) ─────
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
            var row1 = [];
            var row2 = [];
            var overflow = [];
            var section = 'row1'; // which bucket we're currently filling, advances past each divider
            Array.prototype.forEach.call(list.children, function (child) {
                if (child.classList.contains('erp-nav-customize-divider')) {
                    section = section === 'row1' ? 'row2' : 'overflow';
                    return;
                }
                var key = child.getAttribute('data-key');
                if (section === 'row1') row1.push(key);
                else if (section === 'row2') row2.push(key);
                else overflow.push(key);
            });
            currentLayout = reconcile(row1, row2, overflow);
            saveLayout(currentLayout);
            applyLayout(currentLayout);
            modal.classList.remove('open');
        });
    }

    if (resetBtn) {
        resetBtn.addEventListener('click', function () {
            currentLayout = { row1: defaultOrder.slice(), row2: [], overflow: [] };
            saveLayout(currentLayout);
            applyLayout(currentLayout);
            modal.classList.remove('open');
        });
    }
})();
