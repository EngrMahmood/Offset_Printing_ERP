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
})();
