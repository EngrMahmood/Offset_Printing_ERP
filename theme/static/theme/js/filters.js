(function () {
    var filterForms = document.querySelectorAll('.erp-filter-bar');
    filterForms.forEach(function (form) {
        var reset = form.querySelector('[data-filter-reset]');
        if (!reset) return;
        reset.addEventListener('click', function () {
            form.querySelectorAll('input, select').forEach(function (field) {
                if (field.type === 'hidden') return;
                field.value = '';
            });
            form.submit();
        });
    });
})();
