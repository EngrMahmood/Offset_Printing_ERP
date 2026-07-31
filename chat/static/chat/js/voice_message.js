(function () {
    const root = document.getElementById('chat-app-root');
    if (!root) return;

    const voiceBtn = document.getElementById('chat-voice-btn');
    const composer = document.getElementById('chat-composer');
    const recordingBar = document.getElementById('chat-voice-recording');
    const timerEl = document.getElementById('chat-voice-timer');
    const cancelBtn = document.getElementById('chat-voice-cancel-btn');
    const stopBtn = document.getElementById('chat-voice-stop-btn');

    const MAX_DURATION_MS = 120000; // 2 minutes, matches the existing attachment size cap in spirit
    const CANDIDATE_MIME_TYPES = ['audio/webm;codecs=opus', 'audio/ogg;codecs=opus'];

    const rec = {
        stream: null,
        recorder: null,
        chunks: [],
        mimeType: '',
        startedAt: 0,
        timerIntervalId: null,
        maxDurationTimeoutId: null,
        cancelled: false,
    };

    function pickMimeType() {
        if (window.MediaRecorder && MediaRecorder.isTypeSupported) {
            for (let i = 0; i < CANDIDATE_MIME_TYPES.length; i++) {
                if (MediaRecorder.isTypeSupported(CANDIDATE_MIME_TYPES[i])) return CANDIDATE_MIME_TYPES[i];
            }
        }
        return ''; // let the browser pick its own default (e.g. Safari's audio/mp4)
    }

    function formatTimer(ms) {
        const totalSeconds = Math.floor(ms / 1000);
        const minutes = Math.floor(totalSeconds / 60);
        const seconds = totalSeconds % 60;
        return minutes + ':' + (seconds < 10 ? '0' : '') + seconds;
    }

    function setRecordingUI(active) {
        composer.hidden = active;
        recordingBar.hidden = !active;
    }

    async function startRecording() {
        if (!window.isSecureContext || !navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
            alert('This page must be loaded over HTTPS to record voice messages — ask an admin for the https:// link to this ERP.');
            return;
        }
        let stream;
        try {
            stream = await navigator.mediaDevices.getUserMedia({ audio: true });
        } catch (e) {
            const name = e && e.name;
            if (name === 'NotAllowedError' || name === 'PermissionDeniedError') {
                alert('Microphone access was denied. Check this site\'s permissions in your browser settings and try again.');
            } else if (name === 'NotFoundError') {
                alert('No microphone was found on this device.');
            } else {
                alert('Could not access the microphone.');
            }
            return;
        }

        rec.stream = stream;
        rec.chunks = [];
        rec.cancelled = false;
        rec.mimeType = pickMimeType();
        rec.recorder = rec.mimeType ? new MediaRecorder(stream, { mimeType: rec.mimeType }) : new MediaRecorder(stream);

        rec.recorder.addEventListener('dataavailable', function (event) {
            if (event.data && event.data.size > 0) rec.chunks.push(event.data);
        });
        rec.recorder.addEventListener('stop', onRecorderStop);

        rec.recorder.start();
        rec.startedAt = Date.now();
        setRecordingUI(true);
        timerEl.textContent = '0:00';
        rec.timerIntervalId = window.setInterval(function () {
            timerEl.textContent = formatTimer(Date.now() - rec.startedAt);
        }, 250);
        rec.maxDurationTimeoutId = window.setTimeout(function () {
            stopRecording(false);
        }, MAX_DURATION_MS);
    }

    function releaseStream() {
        if (rec.stream) {
            rec.stream.getTracks().forEach(function (t) { t.stop(); });
            rec.stream = null;
        }
    }

    function stopRecording(cancelled) {
        rec.cancelled = cancelled;
        if (rec.timerIntervalId) { window.clearInterval(rec.timerIntervalId); rec.timerIntervalId = null; }
        if (rec.maxDurationTimeoutId) { window.clearTimeout(rec.maxDurationTimeoutId); rec.maxDurationTimeoutId = null; }
        if (rec.recorder && rec.recorder.state !== 'inactive') {
            rec.recorder.stop();
        } else {
            releaseStream();
            setRecordingUI(false);
        }
    }

    function onRecorderStop() {
        releaseStream();
        setRecordingUI(false);
        if (rec.cancelled || !rec.chunks.length) {
            rec.chunks = [];
            return;
        }
        const blob = new Blob(rec.chunks, { type: rec.mimeType || 'audio/webm' });
        rec.chunks = [];
        const ext = (rec.mimeType || 'audio/webm').indexOf('ogg') !== -1 ? 'ogg' : 'webm';
        const file = new File([blob], 'voice-message-' + Date.now() + '.' + ext, { type: rec.mimeType || 'audio/webm' });
        window.ChatApp.uploadFile(file);
    }

    voiceBtn.addEventListener('click', startRecording);
    stopBtn.addEventListener('click', function () { stopRecording(false); });
    cancelBtn.addEventListener('click', function () { stopRecording(true); });
})();
