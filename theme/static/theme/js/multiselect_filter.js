(function () {
    function closeAll(except) {
        document.querySelectorAll('.erp-msf.is-open').forEach(function (el) {
            if (el !== except) {
                el.classList.remove('is-open');
                var panel = el.querySelector('[data-msf-panel]');
                if (panel) panel.hidden = true;
            }
        });
    }

    function updateLabel(root) {
        var labelEl = root.querySelector('[data-msf-label]');
        var placeholder = root.getAttribute('data-msf-placeholder') || 'All';
        var checked = root.querySelectorAll('.erp-msf-options input[type="checkbox"]:checked');
        if (!checked.length) {
            labelEl.textContent = placeholder;
        } else if (checked.length === 1) {
            var lbl = checked[0].closest('.erp-msf-option').querySelector('span');
            labelEl.textContent = lbl ? lbl.textContent : placeholder;
        } else {
            labelEl.textContent = checked.length + ' selected';
        }
    }

    function init(root) {
        var toggle = root.querySelector('[data-msf-toggle]');
        var panel = root.querySelector('[data-msf-panel]');
        var search = root.querySelector('[data-msf-search]');
        var selectAllBtn = root.querySelector('[data-msf-all]');
        var clearBtn = root.querySelector('[data-msf-clear]');
        var options = Array.prototype.slice.call(root.querySelectorAll('.erp-msf-option'));

        updateLabel(root);

        toggle.addEventListener('click', function (e) {
            e.stopPropagation();
            var isOpen = root.classList.contains('is-open');
            closeAll(root);
            root.classList.toggle('is-open', !isOpen);
            panel.hidden = isOpen;
            if (!isOpen && search) {
                search.value = '';
                options.forEach(function (opt) { opt.style.display = ''; });
                search.focus();
            }
        });

        if (search) {
            search.addEventListener('input', function () {
                var term = search.value.trim().toLowerCase();
                options.forEach(function (opt) {
                    var text = opt.textContent.trim().toLowerCase();
                    opt.style.display = !term || text.indexOf(term) !== -1 ? '' : 'none';
                });
            });
        }

        if (selectAllBtn) {
            selectAllBtn.addEventListener('click', function () {
                options.forEach(function (opt) {
                    if (opt.style.display !== 'none') {
                        opt.querySelector('input[type="checkbox"]').checked = true;
                    }
                });
                updateLabel(root);
            });
        }

        if (clearBtn) {
            clearBtn.addEventListener('click', function () {
                options.forEach(function (opt) {
                    if (opt.style.display !== 'none') {
                        opt.querySelector('input[type="checkbox"]').checked = false;
                    }
                });
                updateLabel(root);
            });
        }

        root.querySelectorAll('.erp-msf-options input[type="checkbox"]').forEach(function (cb) {
            cb.addEventListener('change', function () { updateLabel(root); });
        });
    }

    document.querySelectorAll('.erp-msf').forEach(init);

    document.addEventListener('click', function (e) {
        if (!e.target.closest('.erp-msf')) {
            closeAll(null);
        }
    });
})();
