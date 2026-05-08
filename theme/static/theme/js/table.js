(function () {
    document.querySelectorAll('[data-check-all]').forEach(function (master) {
        master.addEventListener('change', function () {
            var target = master.getAttribute('data-check-all');
            if (!target) return;
            document.querySelectorAll(target).forEach(function (cb) {
                cb.checked = master.checked;
            });
        });
    });
})();
