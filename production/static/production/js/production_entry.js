document.addEventListener('DOMContentLoaded', function(){
    const jobSelect = document.getElementById('job_card');
    const jobSearch = document.getElementById('job_card_search');
    const jobInfoCard = document.getElementById('job_info_card');
    const historyCard = document.getElementById('history_card');
    const ji = id => document.getElementById(id);
    const originalJobCardOptions = jobSelect ? Array.from(jobSelect.options).map(opt => ({ value: opt.value, text: opt.text, selected: opt.selected })) : [];

    function parseNumber(v){
        if(!v) return 0;
        const n = parseFloat(v.toString().replace(/,/g, ''));
        return isNaN(n)?0:n;
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
                    <td>${h.impressions}</td>
                    <td>${h.output}</td>
                    <td>${h.waste}</td>
                    <td>${h.intermediate}</td>
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

    function updateSummary(){
        const jobId = jobSelect.value;
        const info = window.JOB_INFO_MAP ? window.JOB_INFO_MAP[jobId] : null;
        if(!info) return;

        const orderQty = parseNumber(info.order_qty);
        const producedBefore = parseNumber(info.produced_qty);
        const good = parseNumber(ji('output_sheets').value);
        const waste = parseNumber(ji('waste_sheets').value);
        const impressions = parseNumber(ji('impressions').value);
        const makeReady = parseNumber(ji('make_ready_time').value);
        const downtime = parseNumber(ji('downtime_minutes').value);
        const intermediatePass = ji('intermediate_pass')?.checked;

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
        const effectivePassCount = intermediatePass ? Math.max(basePassCount, 2) : basePassCount;
        const minImpressions = totalHandled > 0 ? totalHandled * effectivePassCount : 0;
        ji('sum_pass_type').textContent = intermediatePass ? `Intermediate pass (${effectivePassCount}-pass)` : `${effectivePassCount}-pass`;
        ji('sum_min_impressions').textContent = minImpressions.toLocaleString();
        ji('sum_allowed_impressions').textContent = parseNumber(info.allowed_impressions).toLocaleString();
        ji('sum_remaining_impressions').textContent = parseNumber(info.remaining_impressions).toLocaleString();

        const progressGood = intermediatePass ? 0 : good;
        const currentEntry = progressGood;
        const afterSave = producedBefore + progressGood;
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

        const pct = orderQty > 0 ? Math.round((afterSave / orderQty) * 100) : 0;
        ji('progress_fill').style.width = Math.min(100, pct) + '%';
        ji('progress_label').textContent = pct + '%';

        // Status Suggestion
        if(afterSave >= orderQty){
            ji('status').value = 'completed';
        } else {
            ji('status').value = 'in_progress';
        }

        ji('summary_warnings').innerHTML = warnings.join('');
    }

    function filterJobCards(){
        if(!jobSearch || !jobSelect) return;
        const query = jobSearch.value.trim().toLowerCase();
        const selectedValue = jobSelect.value;
        jobSelect.innerHTML = '';
        originalJobCardOptions.forEach(opt => {
            const matches = !opt.value || opt.text.toLowerCase().includes(query) || opt.value === selectedValue;
            if(matches){
                const elem = document.createElement('option');
                elem.value = opt.value;
                elem.text = opt.text;
                if(opt.value === selectedValue) elem.selected = true;
                jobSelect.appendChild(elem);
            }
        });
    }

    jobSelect && jobSelect.addEventListener('change', populateJobInfo);
    jobSearch && jobSearch.addEventListener('input', filterJobCards);

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
    ['output_sheets','waste_sheets','impressions','run_time','make_ready_time','downtime_minutes'].forEach(id=>{
        const el = ji(id);
        el && el.addEventListener('input', updateSummary);
    });

    ji('intermediate_pass')?.addEventListener('change', function(){
        const outputEl = ji('output_sheets');
        if(this.checked){
            outputEl.removeAttribute('required');
        } else {
            outputEl.setAttribute('required', 'required');
        }
        updateSummary();
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
        const intermediatePass = ji('intermediate_pass')?.checked;
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
        if(!intermediatePass && outputVal <= 0){
            actionStatus.textContent = 'Enter good output sheets or mark this as an intermediate pass.';
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

    // initial populate if preselected
    if(jobSelect && jobSelect.value) populateJobInfo();

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