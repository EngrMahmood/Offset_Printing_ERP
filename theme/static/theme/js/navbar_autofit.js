// Shrinks the top-nav module links' font size/padding just enough that every
// module stays visible on one line instead of being clipped by overflow-x.
(function () {
    var STEPS = [13, 12.5, 12, 11.5, 11, 10.5, 10, 9.5, 9];
    var PAD_STEPS = [13, 12, 11, 10, 9, 8, 7, 6, 6];

    function fit() {
        var container = document.querySelector('.erp-topnav-modules');
        if (!container) return;

        container.style.setProperty('--nav-font-size', STEPS[0] + 'px');
        container.style.setProperty('--nav-pad-x', PAD_STEPS[0] + 'px');

        for (var i = 0; i < STEPS.length; i++) {
            if (container.scrollWidth <= container.clientWidth + 1) return;
            container.style.setProperty('--nav-font-size', STEPS[i] + 'px');
            container.style.setProperty('--nav-pad-x', PAD_STEPS[i] + 'px');
        }
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
