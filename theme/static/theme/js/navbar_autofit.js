// Shrinks the top-nav module links' font size/padding just enough that every
// module stays visible on one line instead of being clipped by overflow-x,
// and — for whatever a user hasn't manually pinned via the customize modal —
// auto-cascades overflow from Row 1 into Row 2, then from Row 2 into "More".
(function () {
    var STEPS = [13, 12.5, 12, 11.5, 11, 10.5, 10, 9.5, 9];
    var PAD_STEPS = [13, 12, 11, 10, 9, 8, 7, 6, 6];

    // Every module this function has auto-relocated (out of its manually-set
    // row) is tagged data-auto-origin="row1"/"row2" with the container it
    // structurally belongs to — regardless of how many hops it took to get
    // where it currently sits (row1 -> row2 -> More is two hops but one
    // origin). Restoring by recorded origin, rather than by "undo the last
    // hop", means an item can never end up re-tagged mid-chain and
    // oscillate between containers on alternating fit() passes.
    function restoreAutoPlaced(row1, row2) {
        var tagged = document.querySelectorAll('[data-auto-origin]');
        for (var i = 0; i < tagged.length; i++) {
            var el = tagged[i];
            var origin = el.getAttribute('data-auto-origin');
            el.removeAttribute('data-auto-origin');
            var target = (origin === 'row2' && row2) ? row2 : row1;
            target.appendChild(el);
        }
    }

    function shrinkToFit(container) {
        container.style.setProperty('--nav-font-size', STEPS[0] + 'px');
        container.style.setProperty('--nav-pad-x', PAD_STEPS[0] + 'px');
        for (var i = 0; i < STEPS.length; i++) {
            if (container.scrollWidth <= container.clientWidth + 1) break;
            container.style.setProperty('--nav-font-size', STEPS[i] + 'px');
            container.style.setProperty('--nav-pad-x', PAD_STEPS[i] + 'px');
        }
    }

    // Moves trailing modules out of `container` into `target` until it fits
    // (or only one module is left — an empty row is worse than a slightly-
    // too-wide one). `originTag` records where the module structurally
    // belongs so restoreAutoPlaced can find its way back regardless of how
    // many further hops it takes later in this same pass.
    function overflowInto(container, target, originTag) {
        if (!target) return;
        var items = container.querySelectorAll('a.erp-topnav-module[data-nav-key]');
        while (items.length > 1 && container.scrollWidth > container.clientWidth + 1) {
            var last = items[items.length - 1];
            if (!last.hasAttribute('data-auto-origin')) last.setAttribute('data-auto-origin', originTag);
            target.insertBefore(last, target.firstChild);
            items = container.querySelectorAll('a.erp-topnav-module[data-nav-key]');
        }
    }

    function fit() {
        var row1 = document.getElementById('erp-topnav-modules');
        if (!row1) return;
        var row2 = document.getElementById('erp-topnav-modules-row2');
        var row2Wrap = document.getElementById('erp-topnav-row2');
        var moreMenu = document.getElementById('erp-topnav-more-menu');
        var moreBtn = document.getElementById('erp-topnav-more-btn');

        restoreAutoPlaced(row1, row2);
        if (moreBtn && moreMenu && !moreMenu.children.length) moreBtn.style.display = 'none';
        if (row2Wrap && row2 && !row2.children.length) row2Wrap.classList.remove('has-items');

        shrinkToFit(row1);
        // Row 1's overflow goes to Row 2 if it exists, else straight to More.
        overflowInto(row1, row2 || moreMenu, 'row1');

        if (row2 && row2.children.length) {
            if (row2Wrap) row2Wrap.classList.add('has-items');
            shrinkToFit(row2);
            overflowInto(row2, moreMenu, 'row2');
        }

        if (moreMenu && moreMenu.children.length && moreBtn) moreBtn.style.display = '';
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

    // Watch .erp-topnav-row1 (the whole top bar), never the modules
    // containers themselves. Row1's width is driven purely by the viewport —
    // brand/actions are flex-shrink:0 and modules is flex:1, so nothing
    // inside modules (including the font-size/padding fit() just changed)
    // can change row1's own width. Watching an element fit() mutates was a
    // real bug: any sub-pixel width perturbation from its own font-size
    // change (real WebView rendering, unlike the non-compositing sandbox
    // this was first tested in) re-triggered the observer, which reset to
    // full size before shrinking again — a visible, unending pulse ("nav
    // items jittering"). Row 2 always resizes in lockstep with Row 1 (same
    // viewport-driven width), so one watch target is enough for both.
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
