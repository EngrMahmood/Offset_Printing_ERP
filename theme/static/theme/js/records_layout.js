(function () {
    function measureStickyTop(root) {
        const topnav = document.querySelector('.erp-topnav');
        const contextNav = document.querySelector('.erp-subnav-context');
        let stickyTop = 8;
        if (topnav) {
            stickyTop = topnav.getBoundingClientRect().height + 8;
        }
        if (contextNav) {
            stickyTop += contextNav.getBoundingClientRect().height;
        }
        root.style.setProperty('--erp-records-sticky-top', stickyTop + 'px');
        return stickyTop;
    }

    function fitRecordsTable(wrap) {
        const table = wrap.querySelector('table');
        if (!table) return;

        wrap.classList.remove('is-scroll-x');
        table.style.minWidth = '0';
        table.style.width = '100%';

        const needsHorizontalScroll = table.scrollWidth > wrap.clientWidth + 2;
        if (needsHorizontalScroll) {
            wrap.classList.add('is-scroll-x');
            table.style.minWidth = table.scrollWidth + 'px';
        }
    }

    function fitAllRecordsTables() {
        document.querySelectorAll('.erp-records-table-wrap').forEach(fitRecordsTable);
    }

    function measurePaneOffsets(root) {
        document.querySelectorAll('.erp-records-data-pane').forEach(function (pane) {
            const head = pane.querySelector('.erp-records-pane-head');
            const stickyTop = parseFloat(getComputedStyle(root).getPropertyValue('--erp-records-sticky-top')) || 96;
            let paneOffset = stickyTop + 180;
            if (head) {
                paneOffset = stickyTop + head.getBoundingClientRect().height + 24;
            }
            pane.style.setProperty('--erp-records-pane-offset', paneOffset + 'px');
        });
    }

    function initFilterSidebar(sidebar) {
        const toggle = sidebar.querySelector('.erp-records-filter-toggle');
        const body = sidebar.querySelector('.erp-records-filter-body');
        if (!toggle || !body) return;

        const storageKey = sidebar.getAttribute('data-collapse-key') || 'erpRecordsFilterCollapsed';

        function setCollapsed(collapsed) {
            sidebar.classList.toggle('is-collapsed', collapsed);
            toggle.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
            toggle.textContent = collapsed ? 'Expand' : 'Collapse';
            try {
                sessionStorage.setItem(storageKey, collapsed ? '1' : '0');
            } catch (e) {}
            measurePaneOffsets(document.documentElement);
            fitAllRecordsTables();
        }

        try {
            setCollapsed(sessionStorage.getItem(storageKey) === '1');
        } catch (e) {
            setCollapsed(false);
        }

        toggle.addEventListener('click', function () {
            setCollapsed(!sidebar.classList.contains('is-collapsed'));
        });
    }

    function initRecordsLayout() {
        const root = document.documentElement;
        measureStickyTop(root);
        measurePaneOffsets(root);
        fitAllRecordsTables();

        document.querySelectorAll('.erp-records-filter-sidebar').forEach(initFilterSidebar);

        document.querySelectorAll('.erp-records-table-wrap').forEach(function (wrap) {
            wrap.addEventListener('scroll', function () {
                if (wrap.scrollTop <= 24) return;
                const layout = wrap.closest('.erp-records-layout');
                if (!layout) return;
                const sidebar = layout.querySelector('.erp-records-filter-sidebar.is-collapsible');
                if (sidebar && !sidebar.classList.contains('is-collapsed')) {
                    const toggle = sidebar.querySelector('.erp-records-filter-toggle');
                    if (toggle) toggle.click();
                }
            }, { passive: true });
        });

        window.addEventListener('resize', function () {
            measureStickyTop(root);
            measurePaneOffsets(root);
            fitAllRecordsTables();
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initRecordsLayout);
    } else {
        initRecordsLayout();
    }
})();
