document.addEventListener('DOMContentLoaded', function () {
    if (typeof window.jQuery === 'undefined' || typeof window.jQuery.fn.DataTable === 'undefined') {
        return;
    }

    window.jQuery('#manual-working-table').DataTable({
        scrollX: true,
        autoWidth: false,
        pageLength: 25,
        dom: 'Bfrtip',
        buttons: ['copyHtml5', 'csvHtml5', 'excelHtml5'],
        columnDefs: [
            { targets: '_all', defaultContent: '' },
        ],
        order: [[0, 'desc']],
    });
});
