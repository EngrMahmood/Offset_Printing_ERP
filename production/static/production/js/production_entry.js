document.addEventListener('DOMContentLoaded', function(){
    const isViewMode = Boolean(
        window.IS_VIEW_MODE || new URLSearchParams(window.location.search).has('view')
    );
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
    const ji = id => document.getElementById(id);
    window.__productionEntryLoaded = true;

    let searchTimer = null;
    let activeSearchRequest = null;

    function debounce(fn, delay) {
        return function (...args) {
            clearTimeout(searchTimer);
            searchTimer = setTimeout(() => fn.apply(this, args), delay);
        };
    }
    function parseNumber(v){
        if(!v) return 0;
        const n = parseFloat(v.toString().replace(/,/g, ''));
        return isNaN(n)?0:n;
    }

    const initialImpressions = parseNumber(ji('impressions')?.value);

    function applyViewModeRestrictions(){
        Array.from(document.querySelectorAll('#production_form input, #production_form select, #production_form textarea, #production_form button')).forEach(el => {
            if(el.type === 'hidden') return;
            if(el.id === 'submit_production') return;
            el.disabled = true;
        });
        document.getElementById('submit_production')?.remove();
        const masterButtons = document.querySelectorAll('.master-add');
        masterButtons.forEach(btn => btn.disabled = true);
        const form = document.getElementById('production_form');
        if(form){
            form.addEventListener('submit', function(e){
                if(isViewMode){
                    e.preventDefault();
                }
            });
        }
    }

    function populateJobInfo(){
        const jobId = jobSelect.value;
        const info = window.JOB_INFO_MAP ? window.JOB_INFO_MAP[jobId] : null;

        if(!info){
            jobInfoCard.classList.add('is-hidden');
            historyCard.classList.add('is-hidden');
            return;
        }

        jobInfoCard.classList.remove('is-hidden');
        historyCard.classList.remove('is-hidden');

        ji('ji_job_card').textContent = info.job_card_no;
        ji('ji_customer').textContent = info.customer;
        ji('ji_product').textContent = info.product;
        ji('ji_machine').textContent = info.machine;
        ji('ji_paper').textContent = info.paper;
        ji('ji_gsm').textContent = info.gsm;
        ji('ji_colors').textContent = info.colors;
        ji('ji_job_type').textContent = info.job_type;
        ji('ji_due_date').textContent = info.due_date;
        ji('ji_order_qty').textContent = info.order_qty;
        ji('ji_required_sheets').textContent = info.required_sheets;
        ji('ji_produced').textContent = info.produced_qty;
        ji('ji_remaining').textContent = info.remaining_qty;
        ji('sum_pass_type').textContent = info.pass_type || 'Single Pass';
        renderPassMeters(info);
        syncPrintPassControls(info);
        syncOutputSheetsForPass(info);

        const waitingPlateNotice = document.getElementById('ji_waiting_plate_notice');
        if (waitingPlateNotice) {
            waitingPlateNotice.style.display = info.waiting_for_plate ? 'block' : 'none';
        }

        const mergeNotice = document.getElementById('ji_merge_notice');
        if (mergeNotice) {
            const merge = info.merge;
            if (merge && merge.is_lead) {
                mergeNotice.style.display = 'block';
                mergeNotice.textContent = 'Smart merge ' + merge.code + ': this is the LEAD job. '
                    + 'Enter the combined run once here (' + merge.run_sheets + ' sheets). '
                    + 'The system splits produced pieces to every SKU on the sheet automatically.';
            } else if (merge) {
                mergeNotice.style.display = 'block';
                mergeNotice.textContent = 'Smart merge ' + merge.code + ': this SKU prints on the combined sheet led by '
                    + merge.lead_jc + '. Record printing on the lead job — production is filled in here automatically.';
            } else {
                mergeNotice.style.display = 'none';
                mergeNotice.textContent = '';
            }
        }

        // Machine auto-display + fallback selection logic
        try{
            const machineAuto = document.getElementById('machine_auto_display');
            const machineSelect = document.getElementById('machine');
            const machineOverride = document.getElementById('machine_override');
            const machineMap = window.JOB_MACHINE_MAP || {};
            const mapped = machineMap[jobId] || {};
            // Show both Job Card machine and mapped machine if different
            const jcName = mapped.job_card_machine_name || info.machine || '';
            const mappedName = mapped.mapped_machine_name || '';
            if(jcName && mappedName && jcName !== mappedName){
                machineAuto.textContent = jcName + ' (mapped: ' + mappedName + ')';
            } else if(jcName){
                machineAuto.textContent = jcName;
            } else if(mappedName){
                machineAuto.textContent = mappedName;
            } else {
                machineAuto.textContent = 'No machine mapped in Job Card';
            }
            // Set fallback select to mapped machine if not overriding
            if(!machineOverride || !machineOverride.checked){
                if(mapped.machine_id){
                    machineSelect.value = mapped.machine_id;
                } else {
                    machineSelect.value = '';
                }
                machineSelect.disabled = true;
            }
        } catch(err){ console.warn('Machine mapping update failed', err); }

        // History
        const historyBody = document.getElementById('history_body');
        historyBody.innerHTML = '';
        if(info.history && info.history.length > 0){
            info.history.forEach(h => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${h.date}</td>
                    <td>${h.shift}</td>
                    <td>${h.pass_label || '-'}</td>
                    <td>${h.impressions}</td>
                    <td>${h.output}</td>
                    <td>${h.waste}</td>
                    <td>${h.runtime}</td>
                    <td>${h.make_ready}</td>
                    <td>${h.downtime}</td>
                    <td>${h.status}</td>
                `;
                historyBody.appendChild(tr);
            });
        } else {
            historyBody.innerHTML = '<tr><td colspan="10" style="text-align:center;">No previous entries found</td></tr>';
        }

        updateSummary();
    }

    function getSelectedPassNumber() {
        const passCount = parseNumber(window.JOB_INFO_MAP?.[jobSelect.value]?.pass_count) || 1;
        if (passCount <= 1) {
            return 1;
        }
        return parseInt(document.getElementById('print_pass_number')?.value || '1', 10);
    }

    function isFinalPassSelected(info) {
        const passCount = parseNumber(info?.pass_count) || 1;
        if (passCount <= 1) {
            return true;
        }
        return getSelectedPassNumber() >= passCount;
    }

    function renderPassMeters(info) {
        const panel = document.getElementById('pass_tracking_panel');
        const meters = document.getElementById('ji_pass_meters');
        const passType = document.getElementById('ji_pass_type');
        const totalLine = document.getElementById('ji_total_impressions');
        const legacyNotice = document.getElementById('ji_legacy_notice');
        if (!panel || !meters || !info) {
            return;
        }
        const passCount = parseNumber(info.pass_count) || 1;
        if (passCount <= 1 && !info.legacy_notice) {
            panel.style.display = 'none';
            return;
        }
        panel.style.display = 'block';
        if (passCount <= 1) {
            if (passType) {
                passType.textContent = info.pass_type || 'Single-pass';
            }
            meters.innerHTML = '';
            if (totalLine) {
                totalLine.textContent = `Total impressions: ${info.total_impressions_used_display || info.used_impressions} / ${info.total_impressions_allowed_display || info.allowed_impressions}`;
            }
            if (legacyNotice) {
                legacyNotice.style.display = 'block';
                legacyNotice.textContent = info.legacy_notice;
            }
            return;
        }
        if (passType) {
            passType.textContent = info.pass_type || `${passCount}-pass job`;
        }
        meters.innerHTML = '';
        (info.pass_rows || []).forEach((row) => {
            const line = document.createElement('div');
            line.textContent = `${row.label}: ${row.used_display} / ~${row.budget_display} impressions`;
            meters.appendChild(line);
        });
        if (totalLine) {
            totalLine.textContent = `Total impressions: ${info.total_impressions_used_display || info.used_impressions} / ${info.total_impressions_allowed_display || info.allowed_impressions}`;
        }
        if (legacyNotice) {
            if (info.legacy_notice) {
                legacyNotice.style.display = 'block';
                legacyNotice.textContent = info.legacy_notice;
            } else {
                legacyNotice.style.display = 'none';
                legacyNotice.textContent = '';
            }
        }
    }

    function syncPrintPassControls(info) {
        const field = document.getElementById('print_pass_field');
        const select = document.getElementById('print_pass_number');
        const single = document.getElementById('print_pass_number_single');
        const passCount = parseNumber(info?.pass_count) || 1;
        if (!field || !select || !single) {
            return;
        }
        if (passCount <= 1) {
            field.style.display = 'none';
            select.disabled = true;
            select.removeAttribute('name');
            single.disabled = false;
            single.setAttribute('name', 'print_pass_number');
            single.value = '1';
            return;
        }
        field.style.display = '';
        select.disabled = false;
        select.setAttribute('name', 'print_pass_number');
        single.disabled = true;
        single.removeAttribute('name');
        const suggested = info.suggested_pass || 1;
        const currentValue = window.EDIT_RECORD_PASS || suggested;
        // In edit mode we never disable passes (the operator may be correcting an
        // existing entry on any pass). For new entries, a non-final pass whose
        // per-pass budget is used up is disabled so wrong data can't be logged.
        const isEditMode = window.CURRENT_RECORD_ID !== null && window.CURRENT_RECORD_ID !== undefined;
        const passRows = Array.isArray(info.pass_rows) ? info.pass_rows : [];
        const usedByPass = {};
        passRows.forEach(r => { usedByPass[parseNumber(r.pass_number)] = { used: parseNumber(r.used), budget: parseNumber(r.budget) }; });
        select.innerHTML = '';
        for (let passNo = 1; passNo <= passCount; passNo += 1) {
            const option = document.createElement('option');
            option.value = String(passNo);
            const isFinal = passNo >= passCount;
            const row = usedByPass[passNo];
            const complete = !isFinal && row && row.budget > 0 && row.used >= row.budget;
            option.textContent = isFinal
                ? `Pass ${passNo} (final)`
                : (complete ? `Pass ${passNo} — complete` : `Pass ${passNo}`);
            if (complete && !isEditMode) {
                option.disabled = true;
            }
            select.appendChild(option);
        }
        // Don't land on a disabled option — fall back to the suggested pass.
        const target = select.querySelector(`option[value="${currentValue}"]`);
        select.value = (target && !target.disabled) ? String(currentValue) : String(suggested);
        window.EDIT_RECORD_PASS = null;
    }

    function syncOutputSheetsForPass(info) {
        const outputEl = ji('output_sheets');
        const help = document.getElementById('print_pass_help');
        if (!outputEl || !info) {
            return;
        }
        const passCount = parseNumber(info.pass_count) || 1;
        const passNo = getSelectedPassNumber();
        const finalPass = isFinalPassSelected(info);
        if (isViewMode) {
            return;
        }
        if (finalPass) {
            outputEl.disabled = false;
            outputEl.setAttribute('required', 'required');
            if (help) {
                help.textContent = 'Final pass — enter good sheets and impressions.';
            }
        } else {
            outputEl.value = '';
            outputEl.disabled = true;
            outputEl.removeAttribute('required');
            if (help) {
                help.textContent = `Pass ${passNo} of ${passCount} — impressions and waste only. You can log another Pass ${passNo} entry later if the run resumes.`;
            }
        }
    }

    function updateSummary(){
        const jobId = jobSelect.value;
        const info = window.JOB_INFO_MAP ? window.JOB_INFO_MAP[jobId] : null;
        if(!info) return;

        const orderQtyRaw = info.order_qty;
        const producedBeforeRaw = info.produced_qty;
        const orderQty = parseNumber(orderQtyRaw);
        const producedBefore = parseNumber(producedBeforeRaw);
        const good = parseNumber(ji('output_sheets').value);
        const waste = parseNumber(ji('waste_sheets').value);
        const impressions = parseNumber(ji('impressions').value);
        const makeReady = parseNumber(ji('make_ready_time').value);
        const downtime = parseNumber(ji('downtime_minutes').value);
        const finalPass = isFinalPassSelected(info);
        const currentEntryOutput = parseNumber(ji('output_sheets').value);

        const displayProducedBefore = isViewMode ? Math.max(0, producedBefore - currentEntryOutput) : producedBefore;
        const displayAfterSave = isViewMode ? producedBefore : producedBefore + (finalPass ? currentEntryOutput : 0);

        ji('sum_order_qty').textContent = orderQtyRaw ? orderQty.toLocaleString() : '-';
        ji('sum_produced_before').textContent = displayProducedBefore ? displayProducedBefore.toLocaleString() : '-';

        // Net Impressions
        let netImpressions = impressions;
        ji('sum_net_impressions').textContent = netImpressions.toLocaleString();

        // Runtime
        const runtimeMin = parseNumber(ji('run_time').value);
        ji('sum_runtime').textContent = runtimeMin.toLocaleString() + ' min';
        ji('sum_make_ready').textContent = makeReady.toLocaleString() + ' min';
        ji('sum_downtime').textContent = downtime.toLocaleString() + ' min';

        const totalHandled = good + waste;
        const wastePct = totalHandled > 0 ? (waste / totalHandled * 100) : 0;
        ji('sum_waste_pct').textContent = wastePct.toFixed(1) + '%';

        const basePassCount = parseNumber(info.pass_count) || 1;
        const selectedPass = getSelectedPassNumber();
        const effectivePassCount = finalPass ? basePassCount : selectedPass;
        const minImpressions = totalHandled > 0 ? totalHandled * effectivePassCount : 0;
        ji('sum_pass_type').textContent = finalPass ? `${basePassCount}-pass (final)` : `Pass ${selectedPass} of ${basePassCount}`;
        ji('sum_min_impressions').textContent = minImpressions.toLocaleString();

        const allowedImpressions = parseNumber(info.allowed_impressions);
        const remainingAllowedBefore = parseNumber(info.remaining_impressions);
        const currentRecordImpressions = parseNumber(window.CURRENT_RECORD_IMPRESSIONS);
        const hasCurrentRecord = window.CURRENT_RECORD_ID !== null && window.CURRENT_RECORD_ID !== undefined;
        const remainingAllowedAfter = hasCurrentRecord
            ? Math.max(0, remainingAllowedBefore + currentRecordImpressions - impressions)
            : Math.max(0, remainingAllowedBefore - impressions);

        ji('sum_allowed_impressions').textContent = allowedImpressions.toLocaleString();
        ji('sum_remaining_impressions').textContent = remainingAllowedAfter.toLocaleString();

        const progressGood = finalPass ? good : 0;
        const currentEntry = progressGood;
        const afterSave = isViewMode ? producedBefore : producedBefore + progressGood;
        const remaining = Math.max(0, orderQty - afterSave);

        const warnings = [];
        if(impressions > 0 && totalHandled > 0 && impressions < minImpressions){
            warnings.push(`<div class="alert alert-warning">Impressions are below expected minimum for ${effectivePassCount}-pass output; you may continue saving.</div>`);
        }
        if(wastePct > 10){
            warnings.push(`<div class="alert alert-warning">High waste detected (${wastePct.toFixed(1)}%).</div>`);
            ji('sum_waste_pct').style.color = '#b91c1c';
        } else {
            ji('sum_waste_pct').style.color = '';
        }

        if(afterSave > orderQty){
            const diff = afterSave - orderQty;
            warnings.push(`<div class="alert alert-info">Production exceeds order quantity by ${diff.toLocaleString()} sheets.</div>`);
            ji('sum_after_save').style.color = '#b91c1c';
        } else {
            ji('sum_after_save').style.color = '';
        }

        ji('sum_current_entry').textContent = currentEntry.toLocaleString();
        ji('sum_after_save').textContent = afterSave.toLocaleString();
        ji('sum_remaining').textContent = remaining.toLocaleString();

        const rawPct = orderQty > 0 ? Math.round((afterSave / orderQty) * 100) : 0;
        const pct = rawPct;
        ji('progress_fill').style.width = Math.min(100, pct) + '%';
        ji('progress_label').textContent = pct + '%';

        // Status Suggestion
        if(afterSave >= orderQty){
            ji('status').value = 'completed';
        } else {
            ji('status').value = 'in_progress';
        }

        ji('summary_warnings').innerHTML = warnings.join('');

        updatePassOverridePanel(jobId, info);
    }

    function updatePassOverridePanel(jobId, info) {
        const card = document.getElementById('pass_override_card');
        if (!card || !info) return;
        card.classList.remove('is-hidden');
        const jobIdInput = document.getElementById('pass_override_job_id');
        if (jobIdInput) jobIdInput.value = jobId;

        const hintEl = document.getElementById('pass_override_hint');
        const countInput = document.getElementById('pass_override_count');
        const parts = [];

        if (info.pass_override) {
            parts.push(`<div class="alert alert-info" style="margin:0 0 6px;">Active override: <strong>${info.pass_override} passes</strong>${info.pass_override_reason ? ' — ' + info.pass_override_reason : ''}. Submit with the field blank to clear it.</div>`);
        }
        const hint = info.machine_pass_hint;
        if (hint) {
            parts.push(`<div class="alert alert-warning" style="margin:0 0 6px;">Assigned machine <strong>${hint.machine_name}</strong> is running ${hint.effective_colors} of ${hint.default_colors} colours. This ${hint.colors_per_pass}-colour job needs about <strong>${hint.suggested_passes} passes</strong>. Suggested value pre-filled.</div>`);
            if (countInput && !countInput.value && !info.pass_override) {
                countInput.value = hint.suggested_passes;
            }
        }
        if (hintEl) hintEl.innerHTML = parts.join('');
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

    function selectJobCard(jobId, label, info, machine, plan) {
        if (info) {
            window.JOB_INFO_MAP = window.JOB_INFO_MAP || {};
            window.JOB_INFO_MAP[String(jobId)] = info;
        }
        if (machine) {
            window.JOB_MACHINE_MAP = window.JOB_MACHINE_MAP || {};
            window.JOB_MACHINE_MAP[String(jobId)] = machine;
        }
        if (plan) {
            window.JOB_PLAN_MAP = window.JOB_PLAN_MAP || {};
            window.JOB_PLAN_MAP[String(jobId)] = plan;
        }
        const option = ensureJobOption(jobId, label);
        jobSelect.value = String(jobId);
        jobSearchInput.value = label || option.text;
        jobCardResults.style.display = 'none';
        jobCardSearchMeta.textContent = 'Selected job card loaded.';
        setSelectedChip(label || option.text);
        populateJobInfo();
    }

    function renderSearchResults(results, completedMatches) {
        if (!jobCardResults) return;
        jobCardResults.innerHTML = '';
        if (!results.length) {
            jobCardResults.style.display = 'block';
            if (completedMatches && completedMatches.length) {
                const names = completedMatches.map((m) => `${m.job_card_no} (${m.sku})`).join(', ');
                jobCardSearchMeta.textContent = `${names} already Completed — no new printing entries needed.`;
                const doneRow = document.createElement('div');
                doneRow.className = 'dispatch-search-result-item';
                doneRow.style.cursor = 'default';
                doneRow.innerHTML = `<div class="result-main">✅ ${names}</div><div class="result-meta">Already Completed — see Production Records for its history.</div>`;
                jobCardResults.appendChild(doneRow);
                return;
            }
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
            const merge = item.info && item.info.merge;
            const mergeTag = merge
                ? `<span style="margin-left:6px;padding:1px 6px;border-radius:999px;font-size:11px;font-weight:700;background:${merge.is_lead ? '#ffe08a' : '#ffb3b3'};color:#5a3d00;">`
                    + (merge.is_lead ? `MERGE LEAD ${merge.code}` : `MERGED — see ${merge.lead_jc}`) + '</span>'
                : '';
            row.innerHTML = `
                <div class="result-main">${item.job_card_no} · ${item.sku}${mergeTag}</div>
                <div class="result-meta">${item.customer} · Remaining: ${item.remaining_display} pcs</div>
            `;
            row.addEventListener('click', () => selectJobCard(item.id, item.label, item.info, item.machine, item.plan));
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
                const info = window.JOB_INFO_MAP ? window.JOB_INFO_MAP[opt.value] : null;
                return {
                    id: opt.value,
                    label: opt.text,
                    job_card_no: info ? info.job_card_no : opt.text,
                    sku: info ? info.product : '-',
                    customer: info ? info.customer : '-',
                    remaining_display: info ? info.remaining_qty : '-',
                    info,
                    machine: window.JOB_MACHINE_MAP ? window.JOB_MACHINE_MAP[opt.value] : null,
                    plan: window.JOB_PLAN_MAP ? window.JOB_PLAN_MAP[opt.value] : null,
                };
            });
    }

    function runJobCardSearch() {
        if (!jobSearchInput || !jobSelect) return;
        const query = (jobSearchInput.value || '').trim();
        if (!query) {
            jobCardResults.style.display = 'none';
            jobCardSearchMeta.textContent = 'Start typing to search job cards.';
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
        fetch(`${window.PRINTING_JOB_SEARCH_URL}?${params.toString()}`, {
            headers: { 'X-Requested-With': 'XMLHttpRequest' },
            signal: activeSearchRequest.signal,
        })
            .then((response) => (response.ok ? response.json() : Promise.reject()))
            .then((data) => renderSearchResults(data.results || [], data.completed_matches || []))
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
        jobCardSearchMeta.textContent = 'Start typing to search job cards.';
        setSelectedChip('');
        populateJobInfo();
    }

    jobSelect && jobSelect.addEventListener('change', populateJobInfo);
    jobSearchInput && jobSearchInput.addEventListener('input', debounce(runJobCardSearch, 300));
    jobSearchInput && jobSearchInput.addEventListener('focus', runJobCardSearch);
    jobCardClear && jobCardClear.addEventListener('click', clearJobCardSelection);

    document.addEventListener('click', (event) => {
        if (!jobCardResults?.contains(event.target) && event.target !== jobSearchInput) {
            jobCardResults.style.display = 'none';
        }
    });

    // Machine override toggle
    const machineOverrideEl = document.getElementById('machine_override');
    if(machineOverrideEl){
        machineOverrideEl.addEventListener('change', function(){
            const machineSel = document.getElementById('machine');
            if(this.checked){
                machineSel.disabled = false;
            } else {
                machineSel.disabled = true;
                // when turning off override, re-populate machine from job info
                populateJobInfo();
            }
        });
    }
    if(isViewMode){
        applyViewModeRestrictions();
    }

    if(jobSelect && jobSelect.value){
        const selectedOpt = jobSelect.options[jobSelect.selectedIndex];
        if (selectedOpt && jobSearchInput) {
            jobSearchInput.value = selectedOpt.text;
            jobCardSearchMeta.textContent = 'Selected job card loaded.';
            setSelectedChip(selectedOpt.text);
        }
        populateJobInfo();
    }

    ji('print_pass_number')?.addEventListener('change', function(){
        const info = window.JOB_INFO_MAP ? window.JOB_INFO_MAP[jobSelect.value] : null;
        syncOutputSheetsForPass(info);
        updateSummary();
    });

    ['output_sheets','waste_sheets','impressions','run_time','make_ready_time','downtime_minutes'].forEach(id=>{
        const el = ji(id);
        el && el.addEventListener('input', updateSummary);
    });

    // waste/downtime 'Other' toggles
    ji('waste_reason')?.addEventListener('change', function(){
        if(this.value === 'other') ji('waste_reason_other').classList.remove('is-hidden'); else ji('waste_reason_other').classList.add('is-hidden');
    });
    ji('downtime_category')?.addEventListener('change', function(){
        if(this.value === 'other') ji('downtime_category_other').classList.remove('is-hidden'); else ji('downtime_category_other').classList.add('is-hidden');
    });

    // Submit Production
    ji('submit_production')?.addEventListener('click', function(){
        const form = ji('production_form');
        const actionStatus = ji('action_status');
        const impressionsVal = parseNumber(ji('impressions').value);
        const runTimeVal = parseNumber(ji('run_time').value);
        const makeReadyVal = parseNumber(ji('make_ready_time').value);
        const downtimeVal = parseNumber(ji('downtime_minutes').value);
        const finalPass = isFinalPassSelected(window.JOB_INFO_MAP?.[jobSelect.value] || {});
        const outputVal = parseNumber(ji('output_sheets').value);

        if(impressionsVal < 0){
            actionStatus.textContent = 'Impressions cannot be negative.';
            return;
        }
        if(runTimeVal < 0){
            actionStatus.textContent = 'Run time cannot be negative.';
            return;
        }
        if(makeReadyVal < 0){
            actionStatus.textContent = 'Make ready time cannot be negative.';
            return;
        }
        if(downtimeVal < 0){
            actionStatus.textContent = 'Downtime cannot be negative.';
            return;
        }
        if(runTimeVal === 0){
            actionStatus.textContent = 'Run time should be entered before saving.';
            return;
        }
        const passCount = parseNumber(window.JOB_INFO_MAP?.[jobSelect.value]?.pass_count) || 1;
        const allowedImpressions = parseNumber(window.JOB_INFO_MAP?.[jobSelect.value]?.allowed_impressions) || 0;
        const usedImpressions = parseNumber(window.JOB_INFO_MAP?.[jobSelect.value]?.used_impressions) || 0;
        const remainingAllowed = Math.max(0, allowedImpressions - usedImpressions);

        if(impressionsVal === 0){
            actionStatus.textContent = 'Impressions should be entered before saving.';
            return;
        }
        if(!finalPass && outputVal > 0){
            actionStatus.textContent = 'Good sheets are only allowed on the final print pass.';
            return;
        }
        // Block logging onto a non-final pass that is already complete (new entries only).
        const isEditMode = window.CURRENT_RECORD_ID !== null && window.CURRENT_RECORD_ID !== undefined;
        if(!isEditMode && !finalPass){
            const selPass = getSelectedPassNumber();
            const info = window.JOB_INFO_MAP?.[jobSelect.value];
            const row = (info?.pass_rows || []).find(r => parseNumber(r.pass_number) === selPass);
            if(row && parseNumber(row.budget) > 0 && parseNumber(row.used) >= parseNumber(row.budget)){
                actionStatus.textContent = `Pass ${selPass} is already complete — select the next pass.`;
                return;
            }
        }
        if(finalPass && outputVal <= 0){
            actionStatus.textContent = 'Enter good output sheets for the final print pass.';
            return;
        }
        if(jobSelect.value && impressionsVal > remainingAllowed){
            actionStatus.textContent = `Impressions exceed remaining allowed limit (${remainingAllowed.toLocaleString()}).`;
            return;
        }

        this.disabled = true;
        this.textContent = 'Saving...';
        form.submit();
    });

    /* Master-data modal handling */
    function getCookie(name){
        let v = document.cookie.match('(^|;)\\s*' + name + '\\s*=\\s*([^;]+)');
        return v ? v.pop() : '';
    }

    const masterModal = document.getElementById('master_modal');
    const masterFields = document.getElementById('master_fields');
    const masterTitle = document.getElementById('master_modal_title');
    let currentMasterType = null;

    window.openMasterModal = function(type){
        currentMasterType = type;
        masterFields.innerHTML = '';
        if(type === 'operator'){
            masterTitle.textContent = 'Add Operator';
            masterFields.insertAdjacentHTML('beforeend', '<input name="name" placeholder="Operator name" required class="erp-input">');
            masterFields.insertAdjacentHTML('beforeend', '<input name="employee_code" placeholder="Employee code (optional)" class="erp-input">');
        } else if(type === 'machine'){
            masterTitle.textContent = 'Add Machine';
            masterFields.insertAdjacentHTML('beforeend', '<input name="name" placeholder="Machine name" required class="erp-input">');
            masterFields.insertAdjacentHTML('beforeend', '<input name="standard_impressions_per_hour" placeholder="Impressions/hour (optional)" class="erp-input">');
            masterFields.insertAdjacentHTML('beforeend', '<input name="standard_setup_minutes_per_color" placeholder="Setup minutes per color (optional)" class="erp-input">');
        } else if(type === 'supervisor'){
            masterTitle.textContent = 'Add Supervisor';
            masterFields.insertAdjacentHTML('beforeend', '<input name="name" placeholder="Supervisor name" required class="erp-input">');
            masterFields.insertAdjacentHTML('beforeend', '<input name="employee_code" placeholder="Employee code (optional)" class="erp-input">');
        }
        masterModal.classList.remove('is-hidden');
        masterModal.style.display = 'flex';
        masterModal.setAttribute('aria-hidden','false');
    }

    window.closeMasterModal = function(){
        masterModal.classList.add('is-hidden');
        masterModal.style.display = 'none';
        masterModal.setAttribute('aria-hidden','true');
        masterFields.innerHTML = '';
        currentMasterType = null;
    }

    document.querySelectorAll('.master-add').forEach(btn=>{
        btn.addEventListener('click', function(e){
            const t = this.getAttribute('data-type');
            openMasterModal(t);
        });
    });

    document.getElementById('master_cancel').addEventListener('click', function(){ closeMasterModal(); });

    document.getElementById('master_modal_form').addEventListener('submit', function(e){
        e.preventDefault();
        if(!currentMasterType) return;
        const formData = new FormData(this);
        let url = '/create-operator/';
        if(currentMasterType === 'machine') url = '/create-machine/';
        if(currentMasterType === 'supervisor') url = '/create-supervisor/';
        fetch(url, {
            method: 'POST',
            headers: { 'X-CSRFToken': getCookie('csrftoken') },
            body: formData
        }).then(r=>r.json()).then(data=>{
            if(data && data.error){
                alert('Error: ' + data.error);
                return;
            }
            // add option to relevant select and select it
            if(currentMasterType === 'operator' && data.id){
                const sel = document.getElementById('operator');
                const opt = document.createElement('option'); opt.value = data.id; opt.text = data.name; opt.selected = true; sel.appendChild(opt);
            }
            if(currentMasterType === 'machine' && data.id){
                const sel = document.getElementById('machine');
                const opt = document.createElement('option'); opt.value = data.id; opt.text = data.name; opt.selected = true; sel.appendChild(opt);
                sel.disabled = false; // enable if previously disabled
            }
            if(currentMasterType === 'supervisor' && data.id){
                const sel = document.getElementById('supervisor');
                const opt = document.createElement('option'); opt.value = data.id; opt.text = data.display_name || data.name || 'Supervisor'; opt.selected = true; sel.appendChild(opt);
            }
            closeMasterModal();
        }).catch(err=>{ alert('Save failed'); console.error(err); });
    });
});