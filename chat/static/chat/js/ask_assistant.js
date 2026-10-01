(function () {
    'use strict';

    const ASK_URL = '/api/chat/assistant/ask/';
    const POSITION_KEY = 'ask_ai_launcher_position';
    const DRAG_THRESHOLD_PX = 6;

    function getCookie(name) {
        const match = document.cookie.match(new RegExp('(^| )' + name + '=([^;]+)'));
        return match ? match[2] : '';
    }
    const csrfToken = getCookie('csrftoken');

    function clamp(value, min, max) {
        return Math.max(min, Math.min(max, value));
    }

    // Per-viewer convenience only (where the user last dragged the button on
    // THIS browser) — never required for the widget to work, so every read/
    // write is wrapped and a failure (private window, blocked storage) just
    // means the button falls back to its default corner.
    function loadSavedPosition() {
        try {
            const raw = localStorage.getItem(POSITION_KEY);
            if (!raw) return null;
            const parsed = JSON.parse(raw);
            if (typeof parsed.left === 'number' && typeof parsed.top === 'number') return parsed;
        } catch (e) { /* ignore */ }
        return null;
    }
    function savePosition(left, top) {
        try {
            localStorage.setItem(POSITION_KEY, JSON.stringify({ left: left, top: top }));
        } catch (e) { /* ignore */ }
    }

    function buildWidget() {
        const launcher = document.createElement('button');
        launcher.id = 'ask-ai-launcher';
        launcher.type = 'button';
        launcher.title = 'Ask AI — drag to move';
        launcher.innerHTML = '<i class="fas fa-wand-magic-sparkles"></i><span>Ask AI</span>';
        document.body.appendChild(launcher);

        const overlay = document.createElement('div');
        overlay.id = 'ask-ai-overlay';
        overlay.className = 'is-hidden';
        overlay.innerHTML =
            '<div id="ask-ai-panel">' +
            '<div id="ask-ai-panel__header">' +
            '<h3>Ask about a JC, PO, SKU, dispatch, task, or report</h3>' +
            '<button id="ask-ai-panel__close" type="button" aria-label="Close">&times;</button>' +
            '</div>' +
            '<div id="ask-ai-panel__body">' +
            '<p id="ask-ai-panel__hint">e.g. "status of JC-07-26-PP-0701", "dc no DC-1023", "stock report"</p>' +
            '<div id="ask-ai-answer"></div>' +
            '</div>' +
            '<div id="ask-ai-panel__footer">' +
            '<textarea id="ask-ai-input" rows="2" placeholder="Ask a question..."></textarea>' +
            '<button id="ask-ai-submit" type="button">Ask</button>' +
            '</div>' +
            '</div>';
        document.body.appendChild(overlay);

        const panel = overlay.querySelector('#ask-ai-panel');
        const input = overlay.querySelector('#ask-ai-input');
        const submitBtn = overlay.querySelector('#ask-ai-submit');
        const answerEl = overlay.querySelector('#ask-ai-answer');
        const closeBtn = overlay.querySelector('#ask-ai-panel__close');

        // ---- Positioning (default corner, or the user's last drag) --------

        function applyLauncherPosition(left, top) {
            const rect = launcher.getBoundingClientRect();
            const maxLeft = Math.max(0, window.innerWidth - rect.width);
            const maxTop = Math.max(0, window.innerHeight - rect.height);
            left = clamp(left, 0, maxLeft);
            top = clamp(top, 0, maxTop);
            launcher.style.left = left + 'px';
            launcher.style.top = top + 'px';
            launcher.style.right = 'auto';
            launcher.style.bottom = 'auto';
            return { left: left, top: top };
        }

        const saved = loadSavedPosition();
        if (saved) {
            applyLauncherPosition(saved.left, saved.top);
        } else {
            // Default: bottom-left, deliberately opposite the chat dock
            // (bottom-right) so the two never overlap out of the box.
            const rect = launcher.getBoundingClientRect();
            applyLauncherPosition(16, window.innerHeight - rect.height - 16);
        }

        function positionPanelNearLauncher() {
            const btnRect = launcher.getBoundingClientRect();
            const panelWidth = Math.min(360, window.innerWidth - 32);
            // Prefer opening above the button with room, else below.
            const spaceAbove = btnRect.top;
            const openAbove = spaceAbove > 320;
            let left = btnRect.left;
            let top = openAbove ? null : btnRect.bottom + 8;
            const maxLeft = Math.max(8, window.innerWidth - panelWidth - 8);
            left = clamp(left, 8, maxLeft);
            panel.style.left = left + 'px';
            if (openAbove) {
                panel.style.bottom = (window.innerHeight - btnRect.top + 8) + 'px';
                panel.style.top = 'auto';
            } else {
                panel.style.top = top + 'px';
                panel.style.bottom = 'auto';
            }
        }

        // ---- Drag (pointer events cover mouse + touch) ---------------------

        let dragState = null;

        launcher.addEventListener('pointerdown', function (e) {
            const rect = launcher.getBoundingClientRect();
            dragState = {
                startX: e.clientX,
                startY: e.clientY,
                originLeft: rect.left,
                originTop: rect.top,
                moved: false,
                pointerId: e.pointerId,
            };
            launcher.setPointerCapture(e.pointerId);
        });

        launcher.addEventListener('pointermove', function (e) {
            if (!dragState || dragState.pointerId !== e.pointerId) return;
            const dx = e.clientX - dragState.startX;
            const dy = e.clientY - dragState.startY;
            if (!dragState.moved && Math.hypot(dx, dy) < DRAG_THRESHOLD_PX) return;
            dragState.moved = true;
            launcher.classList.add('is-dragging');
            applyLauncherPosition(dragState.originLeft + dx, dragState.originTop + dy);
        });

        function endDrag(e) {
            if (!dragState || dragState.pointerId !== e.pointerId) return;
            launcher.classList.remove('is-dragging');
            if (dragState.moved) {
                const rect = launcher.getBoundingClientRect();
                savePosition(rect.left, rect.top);
            } else {
                open();
            }
            dragState = null;
        }
        launcher.addEventListener('pointerup', endDrag);
        launcher.addEventListener('pointercancel', endDrag);

        window.addEventListener('resize', function () {
            const rect = launcher.getBoundingClientRect();
            applyLauncherPosition(rect.left, rect.top);
        });

        // ---- Modal open/close/ask -------------------------------------------

        function open() {
            positionPanelNearLauncher();
            overlay.classList.remove('is-hidden');
            input.focus();
        }
        function close() {
            overlay.classList.add('is-hidden');
        }
        function ask() {
            const question = input.value.trim();
            if (!question) return;
            submitBtn.disabled = true;
            answerEl.className = 'is-visible';
            answerEl.textContent = 'Thinking… this can take up to a minute on the local model.';

            fetch(ASK_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-CSRFToken': csrfToken,
                    'X-Requested-With': 'XMLHttpRequest',
                },
                body: JSON.stringify({ question: question }),
            })
                .then(function (r) {
                    return r.json().then(function (data) { return { ok: r.ok, data: data }; });
                })
                .then(function (result) {
                    if (!result.ok) {
                        answerEl.className = 'is-visible is-error';
                        answerEl.textContent = result.data.detail || 'The AI Assistant could not answer that right now.';
                        return;
                    }
                    answerEl.className = 'is-visible';
                    answerEl.textContent = result.data.answer || '(no answer)';
                })
                .catch(function () {
                    answerEl.className = 'is-visible is-error';
                    answerEl.textContent = 'Could not reach the server. Please try again.';
                })
                .finally(function () {
                    submitBtn.disabled = false;
                });
        }

        closeBtn.addEventListener('click', close);
        overlay.addEventListener('click', function (e) {
            if (e.target === overlay) close();
        });
        submitBtn.addEventListener('click', ask);
        input.addEventListener('keydown', function (e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                ask();
            }
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', buildWidget);
    } else {
        buildWidget();
    }
})();
