/**
 * Shared in-page tab component.
 *
 * Markup contract:
 *   <div class="erp-tabs">
 *     <button class="erp-tab-btn is-active" data-target="pane-a">A</button>
 *     <button class="erp-tab-btn" data-target="pane-b">B</button>
 *   </div>
 *   <div id="pane-a" class="erp-tab-content">…</div>
 *   <div id="pane-b" class="erp-tab-content" style="display:none;">…</div>
 *
 * Groups are scoped to the nearest common ancestor of the .erp-tabs bar, so
 * several independent tab sets can coexist on one page.
 */
(function () {
    function activate(btn) {
        var bar = btn.closest('.erp-tabs');
        if (!bar) return;
        var scope = bar.parentElement || document;

        bar.querySelectorAll('.erp-tab-btn').forEach(function (b) {
            b.classList.remove('is-active');
            b.setAttribute('aria-selected', 'false');
        });
        btn.classList.add('is-active');
        btn.setAttribute('aria-selected', 'true');

        var targetId = btn.getAttribute('data-target');
        bar.querySelectorAll('.erp-tab-btn').forEach(function (b) {
            var pane = scope.querySelector('#' + b.getAttribute('data-target'));
            if (pane) pane.style.display = (b === btn) ? 'block' : 'none';
        });

        // Let pages react (e.g. lazy chart sizing) without re-wiring the tabs.
        document.dispatchEvent(new CustomEvent('erp:tab-change', {
            detail: { target: targetId, name: btn.getAttribute('data-tab-name') || targetId }
        }));
    }

    // Delegated, so panes swapped in via AJAX keep working with no re-wiring.
    document.addEventListener('click', function (event) {
        var btn = event.target.closest('.erp-tab-btn');
        if (btn) activate(btn);
    });

    window.erpActivateTabByName = function (name) {
        var btn = Array.from(document.querySelectorAll('.erp-tab-btn'))
            .find(function (b) { return b.getAttribute('data-tab-name') === name; });
        if (btn) activate(btn);
        return !!btn;
    };
})();
