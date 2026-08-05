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

    // The container's own width depends on the viewport and its sibling
    // elements (brand, actions), not on its own font-size — overflow-x:auto
    // makes its flex min-width 0, so this can't feedback-loop off our own
    // shrinking. ResizeObserver catches layout changes that don't dispatch a
    // window 'resize' event (sidebar toggles, zoom, browser chrome changes).
    if (window.ResizeObserver) {
        var observed = null;
        var observer = new ResizeObserver(scheduleFit);
        var watch = function () {
            var container = document.querySelector('.erp-topnav-modules');
            if (container && container !== observed) {
                if (observed) observer.unobserve(observed);
                observer.observe(container);
                observed = container;
            }
        };
        watch();
        document.addEventListener('DOMContentLoaded', watch);
    }

    // Fallback poll: some automation/embedding contexts suppress resize and
    // ResizeObserver notifications (they're gated behind frame compositing).
    // Cheap width check, only re-fits when the available width actually moved.
    var lastWidth = null;
    setInterval(function () {
        var container = document.querySelector('.erp-topnav-modules');
        if (!container) return;
        if (container.clientWidth !== lastWidth) {
            lastWidth = container.clientWidth;
            fit();
        }
    }, 500);
})();
