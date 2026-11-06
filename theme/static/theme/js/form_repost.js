/* Refills a form with what the user typed when the server sent the page back with an error, so nothing is typed twice.
   The data comes from the context processor form_repost; only forms posting to this same page are touched. */
(function () {
    var node = document.getElementById('repost-data');
    if (!node) { return; }
    var repost;
    try { repost = JSON.parse(node.textContent); } catch (e) { return; }
    if (!repost || !repost.data) { return; }

    function fire(el, type) {
        try { el.dispatchEvent(new Event(type, { bubbles: true })); } catch (e) { /* ignore */ }
    }

    function refill() {
        var forms = Array.prototype.filter.call(document.forms, function (f) {
            if ((f.method || '').toLowerCase() !== 'post') { return false; }
            try { return new URL(f.action, location.href).pathname === repost.path; } catch (e) { return false; }
        });
        forms.forEach(function (form) {
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

    // The page's own scripts set their defaults first (job card details, machine defaults); refill after them.
    window.addEventListener('load', function () { setTimeout(refill, 500); setTimeout(refill, 1500); });
})();
