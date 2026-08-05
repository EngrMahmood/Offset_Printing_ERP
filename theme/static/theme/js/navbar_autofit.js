// Shrinks the top-nav module links' font size/padding just enough that every
// module stays visible on one line instead of being clipped by overflow-x.
(function () {
    var STEPS = [13, 12.5, 12, 11.5, 11, 10.5, 10, 9.5, 9];
    var PAD_STEPS = [13, 12, 11, 10, 9, 8, 7, 6, 6];

    // If even the smallest font still doesn't fit (accounts with ~20 modules
    // don't fit at any readable size, even at a 1920px-wide window — measured
    // during testing), the excess has to go somewhere reachable. Auto-move
    // trailing modules into the "More" menu until it fits, and restore them
    // first thing on every fit() pass so growing the window (or removing a
    // manual customization) brings them back instead of leaving them stuck
    // in "More" forever. Tagged data-auto-overflow so this never touches
    // items the user deliberately put in "More" via the customize modal —
    // those stay put regardless of available width.
    function restoreAutoOverflow(container, moreMenu) {
        if (!moreMenu) return;
        var auto = moreMenu.querySelectorAll('[data-auto-overflow]');
        for (var i = 0; i < auto.length; i++) {
            auto[i].removeAttribute('data-auto-overflow');
            container.appendChild(auto[i]);
        }
    }

    function autoOverflow(container, moreMenu, moreBtn) {
        if (!moreMenu || !moreBtn) return;
        var items = container.querySelectorAll('a.erp-topnav-module[data-nav-key]');
        // Always leave at least one module visible even if it still doesn't
        // fit — an empty bar would be worse than a slightly-too-wide one.
        while (items.length > 1 && container.scrollWidth > container.clientWidth + 1) {
            var last = items[items.length - 1];
            last.setAttribute('data-auto-overflow', '1');
            moreMenu.insertBefore(last, moreMenu.firstChild);
            items = container.querySelectorAll('a.erp-topnav-module[data-nav-key]');
        }
        if (moreMenu.children.length) moreBtn.style.display = '';
    }

    function fit() {
        var container = document.querySelector('.erp-topnav-modules');
        if (!container) return;
        var moreMenu = document.getElementById('erp-topnav-more-menu');
        var moreBtn = document.getElementById('erp-topnav-more-btn');

        restoreAutoOverflow(container, moreMenu);
        if (moreBtn && moreMenu && !moreMenu.children.length) moreBtn.style.display = 'none';

        container.style.setProperty('--nav-font-size', STEPS[0] + 'px');
        container.style.setProperty('--nav-pad-x', PAD_STEPS[0] + 'px');

        for (var i = 0; i < STEPS.length; i++) {
            if (container.scrollWidth <= container.clientWidth + 1) break;
            container.style.setProperty('--nav-font-size', STEPS[i] + 'px');
            container.style.setProperty('--nav-pad-x', PAD_STEPS[i] + 'px');
        }

        autoOverflow(container, moreMenu, moreBtn);
    }

    var scheduled = false;
    function scheduleFit() {
        if (scheduled) return;
        scheduled = true;
        setTimeout(function () {
            scheduled = false;
            fit();
        }, 0);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', scheduleFit);
    } else {
        scheduleFit();
    }
    window.addEventListener('resize', scheduleFit);
    if (document.fonts && document.fonts.ready) {
        document.fonts.ready.then(scheduleFit);
    }

    // Watch .erp-topnav-row1 (the whole bar), never .erp-topnav-modules
    // itself. Row1's width is driven purely by the viewport — brand/actions
    // are flex-shrink:0 and modules is flex:1, so nothing inside modules
    // (including the font-size/padding fit() just changed) can change row1's
    // own width. Watching the element fit() mutates was a real bug: any
    // sub-pixel width perturbation from its own font-size change (real
    // WebView rendering, unlike the non-compositing sandbox this was first
    // tested in) re-triggered the observer, which reset to full size before
    // shrinking again — a visible, unending pulse ("nav items jittering").
    var lastRowWidth = null;
    function checkRowWidth() {
        var row = document.querySelector('.erp-topnav-row1');
        if (!row) return;
        var w = row.clientWidth;
        if (w !== lastRowWidth) {
            lastRowWidth = w;
            fit();
        }
    }

    if (window.ResizeObserver) {
        var observed = null;
        var observer = new ResizeObserver(checkRowWidth);
        var watch = function () {
            var row = document.querySelector('.erp-topnav-row1');
            if (row && row !== observed) {
                if (observed) observer.unobserve(observed);
                observer.observe(row);
                observed = row;
            }
        };
        watch();
        document.addEventListener('DOMContentLoaded', watch);
    }

    // Fallback poll: some automation/embedding contexts suppress resize and
    // ResizeObserver notifications (they're gated behind frame compositing).
    // Same row1-only width check, so it can't self-trigger either.
    setInterval(checkRowWidth, 500);
})();
