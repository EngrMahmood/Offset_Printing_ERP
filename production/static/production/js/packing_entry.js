document.addEventListener('DOMContentLoaded', function () {
    const isViewMode = new URLSearchParams(window.location.search).has('view');
    if (isViewMode) {
        return;
    }
    const jobSelect = document.getElementById('job_card');
    const jobSearchInput = document.getElementById('job_card_search');
    const jobCardResults = document.getElementById('job_card_results');
    const jobCardSearchMeta = document.getElementById('job_card_search_meta');
    const jobCardClear = document.getElementById('job_card_clear');
    const jobCardSelectedChip = document.getElementById('job_card_selected_chip');
    const jobCardSelectedLabel = document.getElementById('job_card_selected_label');
    const jobInfoCard = document.getElementById('job_info_card');
    const historyCard = document.getElementById('history_card');

    let searchTimer = null;
    let activeSearchRequest = null;

    function debounce(fn, delay) {
        return function (...args) {
            clearTimeout(searchTimer);
            searchTimer = setTimeout(() => fn.apply(this, args), delay);
        };
    }

    function parseNumber(value) {
        const n = parseInt(String(value || '').replace(/,/g, ''), 10);
        return Number.isNaN(n) ? 0 : n;
    }

    function getInfoMap() {
        return window.PACKING_JOB_INFO_MAP || {};
    }

    function getInfo() {
        return getInfoMap()[jobSelect?.value] || null;
    }

    function ensureJobOption(jobId, label) {
        let option = Array.from(jobSelect.options).find((opt) => opt.value === String(jobId));
        if (!option) {
            option = document.createElement('option');
            option.value = jobId;
            option.textContent = label;
            jobSelect.appendChild(option);
        }
        return option;
    }

    function setSelectedChip(label) {
        if (!label) {
            jobCardSelectedChip?.classList.add('is-hidden');
            if (jobCardSelectedLabel) jobCardSelectedLabel.textContent = '-';
            return;
        }
        jobCardSelectedChip?.classList.remove('is-hidden');
        if (jobCardSelectedLabel) jobCardSelectedLabel.textContent = label;
    }

    function populateJobInfo() {
        const info = getInfo();
        if (!info) {
            jobInfoCard?.classList.add('is-hidden');
            historyCard?.classList.add('is-hidden');
            return;
        }
        jobInfoCard?.classList.remove('is-hidden');
        historyCard?.classList.remove('is-hidden');
        const set = (id, val) => {
            const el = document.getElementById(id);
            if (el) el.textContent = val || '-';
        };
        set('ji_job_card', info.job_card_no);
        set('ji_sku', info.sku);
        set('ji_product', info.product);
        set('ji_process', info.process_type);
        set('ji_order_qty', info.order_qty);
        set('ji_produced_pcs', info.produced_pcs);
        set('ji_pack_limit', info.pack_limit);
        set('ji_packed', info.already_packed);
        set('ji_sort_waste', info.already_sort_waste);
        set('ji_used', info.already_used);
        set('ji_remaining', info.remaining_allowed);
        set('ji_dispatched', info.dispatched);
        set('ji_delivery', info.delivery_date);
        set('ji_material', info.material);
        set('ji_destination', info.destination);

        const historyBody = document.getElementById('history_body');
        if (!historyBody) return;
        historyBody.innerHTML = '';
        if (info.history && info.history.length) {
            info.history.forEach((row) => {
                const tr = document.createElement('tr');
                tr.innerHTML = `<td>${row.date}</td><td>${row.shift}</td><td>${row.packing_qty}</td><td>${row.sorting_waste}</td><td>${row.sorter}</td><td>${row.entered_by}</td>`;
                historyBody.appendChild(tr);
            });
        } else {
            historyBody.innerHTML = '<tr><td colspan="6" style="text-align:center;">No packing history</td></tr>';
        }
        updateSummary();
    }

    function updateSummary() {
        const info = getInfo();
        if (!info) return;
        const pack = parseNumber(document.getElementById('packing_qty')?.value);
        const waste = parseNumber(document.getElementById('sorting_waste_qty')?.value);
        const alreadyPacked = parseNumber(info.already_packed);
        const alreadyWaste = parseNumber(info.already_sort_waste);
        const remaining = parseNumber(info.remaining_allowed);
        const entryTotal = pack + waste;

        const sumEntry = document.getElementById('sum_entry_total');
        const sumPacked = document.getElementById('sum_total_packed');
        const sumWaste = document.getElementById('sum_total_waste');
        const sumRemaining = document.getElementById('sum_remaining');
        const warnings = document.getElementById('summary_warnings');

        if (sumEntry) sumEntry.textContent = entryTotal.toLocaleString();
        if (sumPacked) sumPacked.textContent = (alreadyPacked + pack).toLocaleString();
        if (sumWaste) sumWaste.textContent = (alreadyWaste + waste).toLocaleString();
        if (sumRemaining) sumRemaining.textContent = Math.max(0, remaining - entryTotal).toLocaleString();

        if (warnings) {
            warnings.innerHTML = '';
            if (entryTotal > remaining) {
                warnings.innerHTML = `<div class="alert alert-error">Pack + waste (${entryTotal.toLocaleString()}) exceeds remaining allowance (${remaining.toLocaleString()} pcs).</div>`;
            }
        }
    }

    function selectJobCard(jobId, label, info) {
        if (info) {
            window.PACKING_JOB_INFO_MAP = window.PACKING_JOB_INFO_MAP || {};
            window.PACKING_JOB_INFO_MAP[String(jobId)] = info;
        }
        const option = ensureJobOption(jobId, label);
        jobSelect.value = String(jobId);
        jobSearchInput.value = label || option.text;
        jobCardResults.style.display = 'none';
        jobCardSearchMeta.textContent = 'Selected job card loaded.';
        setSelectedChip(label || option.text);
        populateJobInfo();
    }

    function renderSearchResults(results) {
        jobCardResults.innerHTML = '';
        if (!results.length) {
            jobCardResults.style.display = 'block';
            jobCardSearchMeta.textContent = 'No matching job cards found.';
            const noRow = document.createElement('div');
            noRow.className = 'dispatch-search-result-item';
            noRow.textContent = 'No results';
            noRow.style.cursor = 'default';
            jobCardResults.appendChild(noRow);
            return;
        }
        results.forEach((item) => {
            const row = document.createElement('button');
            row.type = 'button';
            row.className = 'dispatch-search-result-item';
            row.innerHTML = `
                <div class="result-main">${item.job_card_no} · ${item.sku}</div>
                <div class="result-meta">${item.customer} · Remaining: ${item.remaining_display} pcs</div>
            `;
            row.addEventListener('click', () => selectJobCard(item.id, item.label, item.info));
            jobCardResults.appendChild(row);
        });
        jobCardResults.style.display = 'block';
        jobCardSearchMeta.textContent = `Showing ${results.length} result(s). Click one to select.`;
    }

    function filterPreloadedJobCards(query) {
        const options = Array.from(jobSelect.options).filter((opt, idx) => idx > 0);
        return options
            .filter((opt) => !query || (opt.text || '').toLowerCase().includes(query))
            .slice(0, 30)
            .map((opt) => {
                const info = getInfoMap()[opt.value] || null;
                return {
                    id: opt.value,
                    label: opt.text,
                    job_card_no: info ? info.job_card_no : opt.text,
                    sku: info ? info.sku : '-',
                    customer: info ? info.customer : '-',
                    remaining_display: info ? info.remaining_display : '-',
                    info,
                };
            });
    }

    function runJobCardSearch() {
        const query = (jobSearchInput.value || '').trim();
        if (!query) {
            jobCardResults.style.display = 'none';
            jobCardSearchMeta.textContent = 'Start typing to search job cards with printing entry (or cut & pack jobs).';
            return;
        }
        if (query.length < 2) {
            renderSearchResults(filterPreloadedJobCards(query.toLowerCase()));
            return;
        }
        if (activeSearchRequest) activeSearchRequest.abort();
        activeSearchRequest = new AbortController();
        const params = new URLSearchParams({ q: query });
        if (window.CURRENT_RECORD_ID) params.set('edit_id', window.CURRENT_RECORD_ID);
        jobCardSearchMeta.textContent = 'Searching...';
        fetch(`${window.PACKING_JOB_SEARCH_URL}?${params.toString()}`, {
            headers: { 'X-Requested-With': 'XMLHttpRequest' },
            signal: activeSearchRequest.signal,
        })
            .then((response) => (response.ok ? response.json() : Promise.reject()))
            .then((data) => renderSearchResults(data.results || []))
            .catch((err) => {
                if (err.name === 'AbortError') return;
                renderSearchResults(filterPreloadedJobCards(query.toLowerCase()));
            })
            .finally(() => {
                activeSearchRequest = null;
            });
    }

    function clearJobCardSelection() {
        jobSelect.value = '';
        jobSearchInput.value = '';
        jobCardResults.style.display = 'none';
        jobCardSearchMeta.textContent = 'Start typing to search job cards with printing entry (or cut & pack jobs).';
        setSelectedChip('');
        populateJobInfo();
    }

    jobSearchInput?.addEventListener('input', debounce(runJobCardSearch, 300));
    jobSearchInput?.addEventListener('focus', runJobCardSearch);
    jobCardClear?.addEventListener('click', clearJobCardSelection);

    document.addEventListener('click', (event) => {
        if (!jobCardResults?.contains(event.target) && event.target !== jobSearchInput) {
            jobCardResults.style.display = 'none';
        }
    });

    ['packing_qty', 'sorting_waste_qty'].forEach((id) => {
        document.getElementById(id)?.addEventListener('input', updateSummary);
    });

    if (jobSelect?.value) {
        const selectedOpt = jobSelect.options[jobSelect.selectedIndex];
        if (selectedOpt) {
            jobSearchInput.value = selectedOpt.text;
            jobCardSearchMeta.textContent = 'Selected job card loaded.';
            setSelectedChip(selectedOpt.text);
        }
        populateJobInfo();
    }

    if (window.IS_PACKING_VIEW_MODE) {
        document.querySelectorAll('#packing_form input, #packing_form select, #packing_form textarea, #packing_form button').forEach((el) => {
            if (el.type !== 'hidden') el.disabled = true;
        });
    }
});

function openSorterModal() {
    document.getElementById('sorter_modal')?.classList.remove('is-hidden');
}
function closeSorterModal() {
    document.getElementById('sorter_modal')?.classList.add('is-hidden');
}
function saveSorterModal() {
    const name = document.getElementById('sorter_modal_name')?.value?.trim();
    if (!name) return alert('Sorter name is required.');
    const code = document.getElementById('sorter_modal_code')?.value?.trim() || '';
    const csrf = document.querySelector('[name=csrfmiddlewaretoken]')?.value;
    const body = new FormData();
    body.append('name', name);
    body.append('employee_code', code);
    fetch('/create-sorter/', { method: 'POST', headers: { 'X-CSRFToken': csrf }, body })
        .then((r) => r.json())
        .then((data) => {
            if (data.error) {
                alert(data.error);
                return;
            }
            const select = document.getElementById('sorter');
            if (!select) return;
            const opt = document.createElement('option');
            opt.value = data.id;
            opt.textContent = data.display_name || data.name;
            select.appendChild(opt);
            select.value = data.id;
            closeSorterModal();
        })
        .catch(() => alert('Failed to add sorter.'));
}
