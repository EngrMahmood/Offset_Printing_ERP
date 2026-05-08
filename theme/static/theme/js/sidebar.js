(function () {
    var sidebar = document.getElementById('erp-sidebar');
    var toggle = document.getElementById('erp-sidebar-toggle');
    if (!sidebar || !toggle) return;
    toggle.addEventListener('click', function () {
        sidebar.classList.toggle('collapsed');
    });
})();
