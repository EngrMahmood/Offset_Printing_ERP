// Site-wide notification sounds — synthesized via the Web Audio API, no
// binary audio assets to source/commit. Loaded on every authenticated page
// (see theme/base.html) so both the full /chat/ page and the docked-popup
// script can call into one shared instance via window.ChatSound.
(function () {
    const STORAGE_KEY = 'chat_sound_muted';
    const MAX_RING_MS = 30000; // safety cap so a missed stopRingtone() can't ring forever

    let audioCtx = null;
    let ringIntervalId = null;
    let ringStopTimeoutId = null;

    function getContext() {
        if (!audioCtx) {
            const Ctx = window.AudioContext || window.webkitAudioContext;
            if (!Ctx) return null;
            audioCtx = new Ctx();
        }
        return audioCtx;
    }

    function unlock() {
        const ctx = getContext();
        if (ctx && ctx.state === 'suspended') ctx.resume().catch(function () { /* ignore */ });
    }

    function isMuted() {
        try { return localStorage.getItem(STORAGE_KEY) === '1'; } catch (e) { return false; }
    }

    function setMuted(muted) {
        try { localStorage.setItem(STORAGE_KEY, muted ? '1' : '0'); } catch (e) { /* ignore */ }
        if (muted) stopRingtone();
        document.dispatchEvent(new CustomEvent('chatsound:mutechange', { detail: { muted: muted } }));
    }

    function toggle() {
        const next = !isMuted();
        setMuted(next);
        return next;
    }

    function playTone(freq, startTime, duration, gainPeak) {
        const ctx = getContext();
        if (!ctx) return;
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sine';
        osc.frequency.value = freq;
        gain.gain.setValueAtTime(0, startTime);
        gain.gain.linearRampToValueAtTime(gainPeak, startTime + 0.02);
        gain.gain.linearRampToValueAtTime(0, startTime + duration);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(startTime);
        osc.stop(startTime + duration + 0.02);
    }

    function playMessageDing() {
        if (isMuted()) return;
        const ctx = getContext();
        if (!ctx) return;
        const now = ctx.currentTime;
        playTone(880, now, 0.12, 0.15);
        playTone(1318.5, now + 0.09, 0.16, 0.12);
    }

    function ringBurst() {
        const ctx = getContext();
        if (!ctx) return;
        const now = ctx.currentTime;
        // Classic two-tone phone ring (~440Hz + 480Hz), two short bursts.
        [now, now + 0.45].forEach(function (t) {
            playTone(440, t, 0.35, 0.14);
            playTone(480, t, 0.35, 0.12);
        });
    }

    function playRingtone() {
        if (isMuted()) return;
        stopRingtone();
        ringBurst();
        ringIntervalId = window.setInterval(ringBurst, 1600);
        ringStopTimeoutId = window.setTimeout(stopRingtone, MAX_RING_MS);
    }

    function stopRingtone() {
        if (ringIntervalId) { window.clearInterval(ringIntervalId); ringIntervalId = null; }
        if (ringStopTimeoutId) { window.clearTimeout(ringStopTimeoutId); ringStopTimeoutId = null; }
    }

    function playBuzz() {
        // Yahoo-Messenger-style "buzz": a short, loud, rattly burst — a fast
        // sawtooth trill rather than the smooth sine tones used elsewhere,
        // so it reads as an attention-grabbing nudge, not a normal chime.
        if (isMuted()) return;
        const ctx = getContext();
        if (!ctx) return;
        const now = ctx.currentTime;
        for (let i = 0; i < 6; i++) {
            const t = now + i * 0.07;
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.type = 'sawtooth';
            osc.frequency.value = i % 2 === 0 ? 180 : 140;
            gain.gain.setValueAtTime(0, t);
            gain.gain.linearRampToValueAtTime(0.2, t + 0.01);
            gain.gain.linearRampToValueAtTime(0, t + 0.06);
            osc.connect(gain);
            gain.connect(ctx.destination);
            osc.start(t);
            osc.stop(t + 0.07);
        }
    }

    document.addEventListener('click', unlock, { once: true });
    document.addEventListener('keydown', unlock, { once: true });

    window.ChatSound = {
        isMuted: isMuted,
        setMuted: setMuted,
        toggle: toggle,
        unlock: unlock,
        playMessageDing: playMessageDing,
        playRingtone: playRingtone,
        stopRingtone: stopRingtone,
        playBuzz: playBuzz,
    };

    // ---- Mute toggle button (navbar) --------------------------------------

    const toggleBtn = document.getElementById('chat-sound-toggle-btn');
    if (toggleBtn) {
        function syncIcon(muted) {
            toggleBtn.innerHTML = muted ? '<i class="fas fa-volume-mute"></i>' : '<i class="fas fa-volume-up"></i>';
            toggleBtn.title = muted ? 'Unmute chat sounds' : 'Mute chat sounds';
        }
        syncIcon(isMuted());
        toggleBtn.addEventListener('click', function () {
            syncIcon(toggle());
        });
        document.addEventListener('chatsound:mutechange', function (event) {
            syncIcon(event.detail.muted);
        });
    }
})();
