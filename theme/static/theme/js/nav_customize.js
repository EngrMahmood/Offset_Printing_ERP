// Lets each user choose which top-nav modules stay inline vs. get tucked
// into a "More" dropdown, and reorder both — no code change needed. Layout
// is saved server-side (UserProfile.nav_layout) so it's the same on every
// device/browser the user logs into, not just the one it was set on.
(function () {
    var LOCAL_FALLBACK_KEY = 'erp_nav_layout_v1'; // used only if the save request fails

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

    // Catalog: every nav link that exists right now, keyed by its stable
    // data-nav-key. Read from BOTH .erp-topnav-modules and the "More" menu —
    // not just the former. navbar_autofit.js's own auto-overflow pass is
    // scheduled via setTimeout(fn, 0), which on a slow-enough network (this
    // script's own <script src> fetch taking a few ms longer than usual) can
    // fire and complete BEFORE this script even runs, having already moved
    // several modules into the "More" menu. Reading only .erp-topnav-modules
    // made those permanently invisible to this catalog — and since a Save
    // only ever persists what the catalog captured, that loss compounded on
    // every subsequent save (reported: 21 -> 16 -> 8 modules over repeated
    // saves). Reading both containers makes the catalog complete regardless
    // of which one auto-overflow had already sorted each item into.
    var catalog = Array.prototype.slice.call(
        modulesEl.querySelectorAll('a.erp-topnav-module[data-nav-key]')
    ).concat(Array.prototype.slice.call(
        moreMenu ? moreMenu.querySelectorAll('a.erp-topnav-module[data-nav-key]') : []
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

    function readServerLayout() {
        var el = document.getElementById('erp-nav-layout-data');
        if (!el) return null;
        try {
            var raw = JSON.parse(el.textContent || 'null');
            if (raw && Array.isArray(raw.pinned) && Array.isArray(raw.overflow)) return raw;
        } catch (e) { /* fall through */ }
        return null;
    }

    function readLocalFallback() {
        try {
            var raw = JSON.parse(localStorage.getItem(LOCAL_FALLBACK_KEY) || 'null');
            if (raw && Array.isArray(raw.pinned) && Array.isArray(raw.overflow)) return raw;
        } catch (e) { /* ignore */ }
        return null;
    }

    // Guarantees every catalog key appears exactly once across pinned+overflow
    // — regardless of what a saved layout actually contains. Called on every
    // load, every modal open, and every save, so a layout that somehow lost
    // (or duplicated) an entry heals itself the next time it passes through
    // here instead of that module silently staying unreachable/duplicated.
    function reconcile(rawPinned, rawOverflow) {
        var seen = {};
        var pinned = [];
        var overflow = [];
        (rawPinned || []).forEach(function (k) {
            if (nodesByKey[k] && !seen[k]) { seen[k] = true; pinned.push(k); }
        });
        (rawOverflow || []).forEach(function (k) {
            if (nodesByKey[k] && !seen[k]) { seen[k] = true; overflow.push(k); }
        });
        // New modules (permission granted since last save, first ever load,
        // or recovering from a layout that had dropped one) default to
        // pinned so they're never silently unreachable.
        defaultOrder.forEach(function (k) {
            if (!seen[k]) { seen[k] = true; pinned.push(k); }
        });
        return { pinned: pinned, overflow: overflow };
    }

    function loadLayout() {
        var raw = readServerLayout() || readLocalFallback();
        if (!raw) return { pinned: defaultOrder.slice(), overflow: [] };
        return reconcile(raw.pinned, raw.overflow);
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
        layout.pinned.forEach(function (key) {
            nodesByKey[key].removeAttribute('data-auto-overflow');
            modulesEl.appendChild(nodesByKey[key]);
        });
        layout.overflow.forEach(function (key) {
            nodesByKey[key].removeAttribute('data-auto-overflow');
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
        // Reconcile before rendering: if currentLayout ever lost an entry
        // (e.g. a layout saved before this safety net existed), this is
        // where it gets a chance to come back instead of just never
        // appearing in the list a user is trying to fix things from.
        currentLayout = reconcile(currentLayout.pinned, currentLayout.overflow);
        list.innerHTML = '';
        currentLayout.pinned.forEach(function (key) {
            list.appendChild(buildListRow(key));
        });
        list.appendChild(buildDivider());
        currentLayout.overflow.forEach(function (key) {
            list.appendChild(buildListRow(key));
        });
        var countEl = document.getElementById('erp-nav-customize-count');
        if (countEl) {
            var total = currentLayout.pinned.length + currentLayout.overflow.length;
            countEl.textContent = total + ' modules total — scroll within the list to see them all.';
        }
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
            currentLayout = reconcile(pinned, overflow);
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
