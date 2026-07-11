document.addEventListener('DOMContentLoaded', function () {
    if (typeof window.jQuery === 'undefined' || typeof window.jQuery.fn.DataTable === 'undefined') {
        return;
    }

    const $ = window.jQuery;
    const table = $('#manual-working-table');
    const scroll = document.querySelector('.manual-working-table-scroll');
    const topBar = document.querySelector('.manual-working-dt-top');
    const bottomBar = document.querySelector('.manual-working-dt-bottom');

    table.find('tbody td:first-child').addClass('manual-working-sticky-col');

    // Server already paginates; DataTables operates on the current page only.
    // Full filtered Excel/CSV is available via the server Export buttons.
    table.DataTable({
        scrollX: false,
        scrollCollapse: false,
        autoWidth: false,
        paging: false,
        deferRender: true,
        info: true,
        dom: 'frtip',
        columnDefs: [
            { targets: '_all', defaultContent: '' },
            { targets: 0, width: '120px', className: 'manual-working-sticky-col' },
            { targets: '_all', width: '110px' },
        ],
        order: [[0, 'desc']],
        initComplete: function () {
            const api = this.api();
            const wrapper = table.closest('.dataTables_wrapper');
            if (!wrapper.length || !scroll) {
                return;
            }

            wrapper.find('.dataTables_filter').appendTo(topBar);
            wrapper.find('.dataTables_info').appendTo(bottomBar);
            wrapper.find('.dataTables_paginate').appendTo(bottomBar);

            table.detach().appendTo(scroll);
            wrapper.remove();

            api.columns.adjust();
        },
        drawCallback: function () {
            this.api().columns.adjust();
        },
    });
});
