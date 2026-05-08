(function () {
    document.addEventListener('click', function (event) {
        var openTarget = event.target.closest('[data-modal-open]');
        if (openTarget) {
            var id = openTarget.getAttribute('data-modal-open');
            var modal = document.getElementById(id);
            if (modal) modal.classList.add('open');
        }

        var closeTarget = event.target.closest('[data-modal-close]');
        if (closeTarget) {
            var modalNode = closeTarget.closest('.erp-modal');
            if (modalNode) modalNode.classList.remove('open');
        }
    });
})();
