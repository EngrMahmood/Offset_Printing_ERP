(function () {
    'use strict';

    var ROTATE_INTERVAL_MS = 18000;
    var POLL_INTERVAL_MS = 45000;
    var DATA_URL = 'api/data/';

    var screens = Array.prototype.slice.call(document.querySelectorAll('.fd-screen'));
    var dotsContainer = document.getElementById('fd-dots');
    var alertTicker = document.getElementById('fd-alert-ticker');
    var currentIndex = 0;
    var previousData = null;

    function initialData() {
        var el = document.getElementById('fd-initial-data');
        if (!el) return null;
        try {
            return JSON.parse(el.textContent);
        } catch (e) {
            return null;
        }
    }

    function buildDots() {
        screens.forEach(function (_, i) {
            var dot = document.createElement('div');
            dot.className = 'fd-dot' + (i === 0 ? ' active' : '');
            dotsContainer.appendChild(dot);
        });
    }

    function rotate() {
        screens[currentIndex].classList.remove('active');
        dotsContainer.children[currentIndex].classList.remove('active');
        currentIndex = (currentIndex + 1) % screens.length;
        screens[currentIndex].classList.add('active');
        dotsContainer.children[currentIndex].classList.add('active');
    }

    function setText(root, field, value) {
        var el = root.querySelector('[data-field="' + field + '"]');
        if (el) el.textContent = value;
    }

    function renderLeaderboard(root, field, rows, renderRow) {
        var el = root.querySelector('[data-field="' + field + '"]');
        if (!el) return;
        el.innerHTML = '';
        rows.forEach(function (row, i) {
            var li = document.createElement('li');
            li.innerHTML = renderRow(row, i + 1);
            el.appendChild(li);
        });
    }

    function rowMarkup(rank, name, value) {
        return '<span><span class="fd-rank">' + rank + '</span>' + name + '</span><span>' + value + '</span>';
    }

    function toast(message) {
        var div = document.createElement('div');
        div.className = 'fd-alert-toast';
        div.textContent = message;
        alertTicker.appendChild(div);
        setTimeout(function () {
            div.remove();
        }, 6000);
    }

    function checkForAlerts(data) {
        if (!previousData) return;
        var prevOverview = previousData.plant_overview || {};
        var overview = data.plant_overview || {};
        if (overview.completed_jobs > prevOverview.completed_jobs) {
            toast('✅ Job completed');
        }
        var prevDispatch = (previousData.dispatch_performance || {}).dispatch_qty || 0;
        var dispatch = (data.dispatch_performance || {}).dispatch_qty || 0;
        if (dispatch > prevDispatch) {
            toast('🚚 Dispatch completed');
        }
        checkMilestone(prevOverview.printed_pcs || 0, overview.printed_pcs || 0, 100000, '🎉 100,000 pcs milestone reached!');
    }

    function checkMilestone(before, after, threshold, message) {
        if (Math.floor(before / threshold) < Math.floor(after / threshold)) {
            toast(message);
            burstConfetti();
        }
    }

    function burstConfetti() {
        for (var i = 0; i < 24; i++) {
            var piece = document.createElement('div');
            piece.style.position = 'fixed';
            piece.style.top = '-10px';
            piece.style.left = Math.random() * 100 + 'vw';
            piece.style.width = '10px';
            piece.style.height = '10px';
            piece.style.background = ['#34d399', '#fbbf24', '#60a5fa', '#f87171'][i % 4];
            piece.style.zIndex = '999';
            piece.style.borderRadius = '2px';
            piece.style.transition = 'transform 2.5s linear, opacity 2.5s linear';
            document.body.appendChild(piece);
            setTimeout(function (el) {
                el.style.transform = 'translateY(100vh) rotate(360deg)';
                el.style.opacity = '0';
            }, 20, piece);
            setTimeout(function (el) {
                el.remove();
            }, 2600, piece);
        }
    }

    function renderPlantOverview(root, d) {
        setText(root, 'time', d.time);
        setText(root, 'running_jobs', d.running_jobs);
        setText(root, 'completed_jobs', d.completed_jobs);
        setText(root, 'pending_jobs', d.pending_jobs);
        setText(root, 'efficiency_pct', d.efficiency_pct + '%');
        setText(root, 'printed_pcs', d.printed_pcs);
        setText(root, 'dispatched_pcs', d.dispatched_pcs);
        setText(root, 'process_wastage_pct', d.process_wastage_pct + '%');
        setText(root, 'plant_utilization_pct', d.plant_utilization_pct + '%');
        setText(root, 'machines_sub', d.active_machines + '/' + d.total_machines + ' machines');
        setText(root, 'target_progress_pct', d.target_progress_pct + '%');
        setText(root, 'printed_of_target', d.packed_pcs);
        setText(root, 'target_qty', d.target_qty);
        setText(root, 'target_remaining', d.target_remaining);
        setText(root, 'target_estimated_tag', d.target_is_estimated ? '(estimated)' : '');
        var bar = root.querySelector('[data-field="target_progress_bar"]');
        if (bar) bar.style.width = Math.min(d.target_progress_pct, 100) + '%';
    }

    function renderPrinting(root, d) {
        setText(root, 'impressions_total', d.impressions_total);
        setText(root, 'output_sheets_total', d.output_sheets_total);
        setText(root, 'waste_pct', d.waste_pct + '%');
        setText(root, 'avg_run_speed', d.avg_run_speed);
        renderLeaderboard(root, 'top_machines', d.top_machines, function (m, rank) {
            return rowMarkup(rank, m.name, m.efficiency_pct + '%');
        });
        renderLeaderboard(root, 'top_operators', d.top_operators, function (o, rank) {
            return rowMarkup(rank, o.name, o.efficiency_pct + '%');
        });
    }

    function renderPacking(root, d) {
        setText(root, 'packed_pcs', d.packed_pcs);
        setText(root, 'sorting_waste_pcs', d.sorting_waste_pcs);
        setText(root, 'target_progress_pct', d.target_progress_pct + '%');
        renderLeaderboard(root, 'top_sorters', d.top_sorters, function (s, rank) {
            return rowMarkup(rank, s.name, s.packed_qty + ' pcs · ' + s.waste_pct + '% waste');
        });
    }

    function renderDispatch(root, d) {
        setText(root, 'dispatch_qty', d.dispatch_qty);
        setText(root, 'dc_count', d.dc_count);
        setText(root, 'completion_pct', d.completion_pct + '%');
    }

    function renderShiftComparison(root, d) {
        var container = root.querySelector('[data-field="shifts"]');
        if (container) {
            container.innerHTML = '';
            d.shifts.forEach(function (s) {
                var div = document.createElement('div');
                div.className = 'fd-card';
                div.innerHTML = '<div class="fd-card-label">' + s.label + '</div>' +
                    '<div class="fd-card-value">' + s.efficiency_pct + '%</div>' +
                    '<div class="fd-card-sub">Output ' + s.output_sheets + ' sheets · Waste ' + s.waste_pct + '%</div>';
                container.appendChild(div);
            });
        }
        setText(root, 'winning_shift', d.winning_shift ? ('Winning Shift: ' + d.winning_shift) : '');
    }

    function renderPeriodSummary(root, d) {
        var container = root.querySelector('[data-field="periods"]');
        if (!container) return;
        container.innerHTML = '';
        d.periods.forEach(function (p) {
            var div = document.createElement('div');
            div.className = 'fd-card';
            div.innerHTML = '<div class="fd-card-label">' + p.label + '<br><span style="font-size:14px;">' + p.range_label + '</span></div>' +
                '<div class="fd-card-value fd-blue">' + p.printed_pcs + '</div>' +
                '<div class="fd-card-sub">printed pcs</div>' +
                '<div class="fd-card-sub">Packed ' + p.packed_pcs + ' · Dispatched ' + p.dispatched_pcs + '</div>' +
                '<div class="fd-card-sub">Wastage ' + p.process_wastage_pct + '% · Efficiency ' + p.efficiency_pct + '%</div>';
            container.appendChild(div);
        });
    }

    function renderMachineLeaderboard(root, d) {
        renderLeaderboard(root, 'machines', d.machines, function (m, rank) {
            return rowMarkup(rank, m.name, m.output_sheets + ' sheets · ' + m.efficiency_pct + '% eff');
        });
        setText(root, 'machine_of_the_day', d.machine_of_the_day ? ('🏆 Machine of the Day: ' + d.machine_of_the_day) : '');
    }

    function renderOperatorLeaderboard(root, d) {
        renderLeaderboard(root, 'operators', d.operators, function (o, rank) {
            return rowMarkup(rank, o.name, o.output_sheets + ' sheets');
        });
        renderLeaderboard(root, 'sorters', d.sorters, function (s, rank) {
            return rowMarkup(rank, s.name, s.packed_qty + ' pcs');
        });
    }

    function renderWastageQuality(root, d) {
        setText(root, 'printing_waste_pcs', d.printing_waste_pcs);
        setText(root, 'sorting_waste_pcs', d.sorting_waste_pcs);
        setText(root, 'quality_score', d.quality_score);
        setText(root, 'flagged_job_count', d.flagged_job_count);
    }

    function renderTargetAchievement(root, d) {
        setText(root, 'target_qty', d.target_qty);
        setText(root, 'achieved_qty', d.achieved_qty);
        setText(root, 'remaining_qty', d.remaining_qty);
        setText(root, 'prediction', d.prediction);
        setText(root, 'forecast_qty', d.forecast_qty);
        setText(root, 'hourly_rate', d.hourly_rate);
        setText(root, 'target_estimated_tag', d.target_is_estimated ? '(estimated)' : '');
    }

    function renderRecognition(root, d) {
        var container = root.querySelector('[data-field="badges"]');
        if (container) {
            container.innerHTML = '';
            d.badges.forEach(function (b) {
                var div = document.createElement('div');
                div.className = 'fd-badge';
                div.innerHTML = '<div class="fd-badge-title">' + b.title + '</div>' +
                    '<div class="fd-badge-name">' + b.name + '</div>' +
                    '<div class="fd-card-sub">' + (b.detail || '') + '</div>';
                container.appendChild(div);
            });
        }
        setText(root, 'motivational_message', '"' + d.motivational_message + '"');
    }

    var RENDERERS = {
        plant_overview: renderPlantOverview,
        period_summary: renderPeriodSummary,
        printing_performance: renderPrinting,
        packing_performance: renderPacking,
        dispatch_performance: renderDispatch,
        shift_comparison: renderShiftComparison,
        machine_leaderboard: renderMachineLeaderboard,
        operator_leaderboard: renderOperatorLeaderboard,
        wastage_quality: renderWastageQuality,
        target_achievement: renderTargetAchievement,
        recognition: renderRecognition,
    };

    function renderAll(data) {
        screens.forEach(function (screen) {
            var key = screen.getAttribute('data-screen');
            var renderer = RENDERERS[key];
            if (renderer && data[key]) {
                renderer(screen, data[key]);
            }
        });
        var overviewScreen = document.querySelector('[data-screen="plant_overview"]');
        if (overviewScreen) {
            setText(overviewScreen, 'display_date_note', data.is_fallback_date
                ? ('No activity logged yet today — showing the most recent active day: ' + data.display_date)
                : ('Live data for ' + data.display_date));
        }
    }

    function poll() {
        fetch(DATA_URL)
            .then(function (resp) { return resp.json(); })
            .then(function (data) {
                checkForAlerts(data);
                renderAll(data);
                previousData = data;
            })
            .catch(function () {
                // Network hiccup on a kiosk TV — keep showing the last good data,
                // next poll will retry automatically.
            });
    }

    document.addEventListener('DOMContentLoaded', function () {
        buildDots();
        previousData = initialData();
        setInterval(rotate, ROTATE_INTERVAL_MS);
        setInterval(poll, POLL_INTERVAL_MS);
    });
})();
