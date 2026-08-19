// Click-to-toggle dropdowns for the top-nav's grouped module buttons
// (Offset, Admin Tools) and the "More" overflow menu — one shared
// implementation instead of one bespoke wiring per menu. Each menu is
// position:fixed and positioned here on open (same pattern as
// notifications.js's positionPanel()), clamped to the viewport so it can
// never run off either edge of a narrow phone screen regardless of where
// its trigger button currently sits in the bar (including after
// navbar_autofit.js has auto-shrunk it into Row 2 or "More" itself).
(function () {
    function pairs() {
        var list = [];
        var moreBtn = document.getElementById('erp-topnav-more-btn');
        var moreMenu = document.getElementById('erp-topnav-more-menu');
        if (moreBtn && moreMenu) list.push({ btn: moreBtn, menu: moreMenu });
        var groupBtns = document.querySelectorAll('.erp-topnav-group-btn');
        for (var i = 0; i < groupBtns.length; i++) {
            var btn = groupBtns[i];
            var menu = btn.parentNode.querySelector('.erp-topnav-group-menu');
            if (menu) list.push({ btn: btn, menu: menu });
        }
        return list;
    }

    function positionMenu(btn, menu) {
        var rect = btn.getBoundingClientRect();
        var width = Math.max(menu.offsetWidth, 190);
        var left = rect.left;
        if (left + width > window.innerWidth - 8) left = window.innerWidth - width - 8;
        if (left < 8) left = 8;
        menu.style.left = left + 'px';
        menu.style.top = (rect.bottom + 6) + 'px';
    }

    function openPairs() {
        return pairs().filter(function (p) { return p.menu.classList.contains('is-open'); });
    }

    // A group button that has itself been auto-overflowed into "More" (or
    // into another group's menu, in principle) lives inside another tracked
    // menu. Closing that ancestor menu before opening this one would hide
    // the button's own container out from under it — so closing here skips
    // any menu that currently contains the button just clicked, and can
    // legitimately leave two menus open at once (e.g. "More" behind
    // "Admin Tools" nested inside it).
    function closeOthers(btn) {
        pairs().forEach(function (p) {
            if (p.menu.contains(btn)) return;
            p.menu.classList.remove('is-open');
        });
    }

    pairs().forEach(function (p) {
        p.btn.addEventListener('click', function (event) {
            event.stopPropagation();
            var willOpen = !p.menu.classList.contains('is-open');
            closeOthers(p.btn);
            if (willOpen) {
                positionMenu(p.btn, p.menu);
                p.menu.classList.add('is-open');
            } else {
                p.menu.classList.remove('is-open');
            }
        });
    });

    document.addEventListener('click', function (event) {
        openPairs().forEach(function (p) {
            if (p.btn.contains(event.target) || p.menu.contains(event.target)) return;
            p.menu.classList.remove('is-open');
        });
    });
    document.addEventListener('scroll', function () {
        openPairs().forEach(function (p) { p.menu.classList.remove('is-open'); });
    }, true);
    window.addEventListener('resize', function () {
        openPairs().forEach(function (p) { positionMenu(p.btn, p.menu); });
    });
})();
