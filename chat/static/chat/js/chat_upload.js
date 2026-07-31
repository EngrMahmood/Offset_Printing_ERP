(function () {
    const root = document.getElementById('chat-app-root');
    if (!root) return;

    const attachBtn = document.getElementById('chat-attach-btn');
    const fileInput = document.getElementById('chat-file-input');
    const progressEl = document.getElementById('chat-upload-progress');

    attachBtn.addEventListener('click', function () {
        fileInput.click();
    });

    fileInput.addEventListener('change', function () {
        const file = fileInput.files[0];
        fileInput.value = '';
        if (!file) return;
        uploadFile(file);
    });

    const messageInput = document.getElementById('chat-message-input');
    if (messageInput) {
        messageInput.addEventListener('paste', function (event) {
            const items = (event.clipboardData || window.clipboardData || {}).items || [];
            for (let i = 0; i < items.length; i++) {
                const item = items[i];
                if (item.type && item.type.indexOf('image/') === 0) {
                    event.preventDefault();
                    const file = item.getAsFile();
                    if (!file) return;
                    const ext = item.type.split('/')[1] || 'png';
                    const named = new File([file], 'pasted-image-' + Date.now() + '.' + ext, { type: item.type });
                    uploadFile(named);
                    return;
                }
            }
            // No image in the clipboard — let the default text-paste behavior proceed.
        });
    }

    function uploadFile(file) {
        const app = window.ChatApp;
        const roomId = app.currentRoomId;
        if (!roomId) return;

        const url = app.urls.attachments.replace('/0/', '/' + roomId + '/');
        const formData = new FormData();
        formData.append('file', file);

        const xhr = new XMLHttpRequest();
        xhr.open('POST', url, true);
        xhr.setRequestHeader('X-CSRFToken', app.csrfToken);
        xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');

        progressEl.hidden = false;
        progressEl.textContent = 'Uploading ' + file.name + '… 0%';

        xhr.upload.onprogress = function (event) {
            if (!event.lengthComputable) return;
            const pct = Math.round((event.loaded / event.total) * 100);
            progressEl.textContent = 'Uploading ' + file.name + '… ' + pct + '%';
        };

        xhr.onload = function () {
            progressEl.hidden = true;
            if (xhr.status >= 200 && xhr.status < 300) {
                const message = JSON.parse(xhr.responseText);
                app.appendMessage(message, false);
                app.scrollToBottom();
                app.loadRoomList();
            } else {
                let detail = 'Upload failed.';
                try { detail = JSON.parse(xhr.responseText).detail || detail; } catch (e) { /* ignore */ }
                alert(detail);
            }
        };

        xhr.onerror = function () {
            progressEl.hidden = true;
            alert('Upload failed — network error.');
        };

        xhr.send(formData);
    }

    // Exposed so voice_message.js can reuse this exact upload pipeline for
    // recorded audio blobs, without duplicating the XHR/progress logic.
    window.ChatApp.uploadFile = uploadFile;
})();
