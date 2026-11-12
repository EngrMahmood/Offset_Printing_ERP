/* When the server sends a form page back with an error:
   1. refill the form with what the user typed, so nothing has to be entered twice;
   2. mark only the fields that need attention in red, with the message under each one.
   The data comes from the context processor form_repost; only forms posting to this same page are touched. */
(function () {
    var node = document.getElementById('repost-data');
    if (!node) { return; }
    var repost;
    try { repost = JSON.parse(node.textContent); if (typeof repost === 'string') { repost = JSON.parse(repost); } } catch (e) { return; }
    if (!repost || !repost.data) { return; }
    var errors = repost.errors || {};

    var style = document.createElement('style');
    style.textContent = '.field-has-error{border:2px solid #c0392b !important;background:#fdeceb !important;outline:none}' +
        '.field-error-msg{color:#c0392b;font-size:12px;font-weight:600;margin:3px 0 0;display:block}';
    document.head.appendChild(style);

    function fire(el, type) {
        try { el.dispatchEvent(new Event(type, { bubbles: true })); } catch (e) { /* ignore */ }
    }

    function myForms() {
        return Array.prototype.filter.call(document.forms, function (f) {
            if ((f.method || '').toLowerCase() !== 'post') { return false; }
            try { return new URL(f.action, location.href).pathname === repost.path; } catch (e) { return false; }
        });
    }

    function refill() {
        myForms().forEach(function (form) {
            Object.keys(repost.data).forEach(function (name) {
                var wanted = repost.data[name];
                var fields = form.querySelectorAll('[name="' + name.replace(/"/g, '\\"') + '"]');
                if (!fields.length) { return; }
                var list = Array.isArray(wanted) ? wanted : [wanted];
                Array.prototype.forEach.call(fields, function (el, index) {
                    if (el.disabled || el.type === 'file' || el.type === 'hidden' && name === 'edit_id') { return; }
                    if (el.type === 'checkbox' || el.type === 'radio') {
                        el.checked = list.indexOf(el.value) !== -1;
                        return;
                    }
                    var value = Array.isArray(wanted) && name.slice(-2) === '[]' ? (list[index] !== undefined ? list[index] : '') : list[0];
                    if (el.tagName === 'SELECT') {
                        var has = Array.prototype.some.call(el.options, function (o) { return o.value === value; });
                        if (!has) { return; }
                        if (el.value !== value) { el.value = value; fire(el, 'change'); }
                    } else if (el.value !== value) {
                        el.value = value;
                        fire(el, 'input');
                    }
                });
            });
        });
    }

    function clearMark(el) {
        el.classList.remove('field-has-error');
        var next = el.parentNode && el.parentNode.querySelector('.field-error-msg[data-for="' + (el.id || el.name) + '"]');
        if (next) { next.remove(); }
    }

    function mark(el, message) {
        el.classList.add('field-has-error');
        var key = el.id || el.name;
        if (!el.parentNode.querySelector('.field-error-msg[data-for="' + key + '"]')) {
            var note = document.createElement('span');
            note.className = 'field-error-msg';
            note.setAttribute('data-for', key);
            note.textContent = message;
            (el.closest('.erp-form-row, .form-group, .erp-field') || el.parentNode).appendChild(note);
        }
        var drop = function () { clearMark(el); el.removeEventListener('input', drop); el.removeEventListener('change', drop); };
        el.addEventListener('input', drop);
        el.addEventListener('change', drop);
    }

    function showErrors() {
        var first = null;
        myForms().forEach(function (form) {
            Object.keys(errors).forEach(function (key) {
                var parts = key.split('#');
                var fields = form.querySelectorAll('[name="' + parts[0].replace(/"/g, '\\"') + '"]');
                var el = parts.length > 1 ? fields[parseInt(parts[1], 10)] : fields[0];
                if (!el || el.type === 'hidden') { return; }
                mark(el, errors[key]);
                if (!first) { first = el; }
            });
        });
        if (first) {
            try { first.scrollIntoView({ block: 'center', behavior: 'smooth' }); first.focus({ preventScroll: true }); } catch (e) { /* ignore */ }
        }
    }

    // The page's own scripts set their defaults first (job card details, machine defaults); refill after them,
    // and again a little later in case they load details asynchronously.
    window.addEventListener('load', function () {
        setTimeout(refill, 400);
        setTimeout(function () { refill(); showErrors(); }, 1500);
    });
})();
