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

    table.DataTable({
        scrollX: false,
        scrollCollapse: false,
        autoWidth: false,
        pageLength: 25,
        dom: 'Bfrtip',
        buttons: ['copyHtml5', 'csvHtml5', 'excelHtml5'],
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

            wrapper.find('.dt-buttons').appendTo(topBar);
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
