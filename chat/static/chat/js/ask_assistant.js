(function () {
    'use strict';

    const ASK_URL = '/api/chat/assistant/ask/';

    function getCookie(name) {
        const match = document.cookie.match(new RegExp('(^| )' + name + '=([^;]+)'));
        return match ? match[2] : '';
    }
    const csrfToken = getCookie('csrftoken');

    function buildWidget() {
        const launcher = document.createElement('button');
        launcher.id = 'ask-ai-launcher';
        launcher.type = 'button';
        launcher.innerHTML = '<i class="fas fa-wand-magic-sparkles"></i><span>Ask AI</span>';

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

        document.body.appendChild(launcher);
        document.body.appendChild(overlay);

        const panel = overlay.querySelector('#ask-ai-panel');
        const input = overlay.querySelector('#ask-ai-input');
        const submitBtn = overlay.querySelector('#ask-ai-submit');
        const answerEl = overlay.querySelector('#ask-ai-answer');
        const closeBtn = overlay.querySelector('#ask-ai-panel__close');

        function open() {
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

        launcher.addEventListener('click', open);
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
