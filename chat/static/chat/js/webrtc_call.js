(function () {
    const root = document.getElementById('chat-app-root');
    if (!root) return;

    const canInitiateCall = root.dataset.canInitiateCall === '1';
    const currentUserId = parseInt(root.dataset.currentUserId, 10);

    const overlay = document.getElementById('chat-call-overlay');
    const titleEl = document.getElementById('chat-call-title');
    const videosEl = document.getElementById('chat-call-videos');
    const acceptBtn = document.getElementById('chat-call-accept-btn');
    const declineBtn = document.getElementById('chat-call-decline-btn');
    const declineLabel = document.getElementById('chat-call-decline-label');
    const muteBtn = document.getElementById('chat-call-mute-btn');
    const screenshareBtn = document.getElementById('chat-call-screenshare-btn');
    const audioBtn = document.getElementById('chat-call-audio-btn');
    const videoBtn = document.getElementById('chat-call-video-btn');
    const minimizeBtn = document.getElementById('chat-call-minimize-btn');
    const overlayBox = document.querySelector('#chat-call-overlay .chat-call-overlay__box');

    let iceServers = [{ urls: ['stun:stun.l.google.com:19302'] }];
    fetch(root.dataset.iceConfigUrl, { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
        .then(function (r) { return r.json(); })
        .then(function (data) { if (data.ice_servers) iceServers = data.ice_servers; })
        .catch(function () { /* fall back to default STUN */ });

    const call = {
        active: false,
        roomId: null,
        callId: null,
        callType: 'audio',
        socket: null,
        localStream: null,
        peers: {}, // userId -> RTCPeerConnection
        isIncoming: false,
        isMuted: false,
        connectedAt: null,
        timerInterval: null,
    };

    function resetCallUI() {
        overlay.hidden = true;
        overlay.classList.remove('is-minimized');
        videosEl.innerHTML = '';
        acceptBtn.hidden = true;
        muteBtn.hidden = true;
        screenshareBtn.hidden = true;
        screenshareBtn.classList.remove('is-active');
        declineLabel.textContent = 'Decline';
        stopCallTimer();
    }

    function setMinimized(minimized) {
        overlay.classList.toggle('is-minimized', minimized);
        minimizeBtn.innerHTML = minimized
            ? '<i class="fas fa-window-restore"></i>'
            : '<i class="fas fa-window-minimize"></i>';
    }

    minimizeBtn.addEventListener('click', function (event) {
        event.stopPropagation();
        setMinimized(!overlay.classList.contains('is-minimized'));
    });
    // While minimized, the whole pill acts as a "restore" button — except its
    // own action buttons (mute/decline/etc.), which should keep working
    // without expanding the overlay first.
    overlayBox.addEventListener('click', function (event) {
        if (overlay.classList.contains('is-minimized') && !event.target.closest('button')) {
            setMinimized(false);
        }
    });

    function showOverlay(title) {
        overlay.hidden = false;
        overlay.classList.remove('is-minimized');
        minimizeBtn.innerHTML = '<i class="fas fa-window-minimize"></i>';
        titleEl.textContent = title;
    }

    function formatDuration(totalSeconds) {
        const m = Math.floor(totalSeconds / 60);
        const s = totalSeconds % 60;
        return (m < 10 ? '0' : '') + m + ':' + (s < 10 ? '0' : '') + s;
    }

    function startCallTimer() {
        if (call.timerInterval) return; // already running
        call.connectedAt = Date.now();
        titleEl.textContent = '00:00';
        call.timerInterval = setInterval(function () {
            titleEl.textContent = formatDuration(Math.floor((Date.now() - call.connectedAt) / 1000));
        }, 1000);
    }

    function stopCallTimer() {
        if (call.timerInterval) { clearInterval(call.timerInterval); call.timerInterval = null; }
        call.connectedAt = null;
    }

    // Any single peer actually reaching 'connected' is enough to consider the
    // call live — a group call's other legs may still be negotiating, but the
    // user watching this overlay already has working audio/video with at
    // least one participant.
    function onPeerConnectionStateChange(pc) {
        if (pc.connectionState === 'connected' && !call.timerInterval) {
            startCallTimer();
        }
    }

    function addVideoEl(userId, stream, isLocal) {
        let video = document.getElementById('chat-call-video-' + userId);
        if (!video) {
            video = document.createElement('video');
            video.id = 'chat-call-video-' + userId;
            video.autoplay = true;
            video.playsInline = true;
            video.title = 'Double-click to enlarge';
            if (isLocal) video.muted = true;
            videosEl.appendChild(video);
        }
        video.srcObject = stream;
    }

    // Double-click any video (own camera, remote camera, or a shared screen)
    // to view it fullscreen — the small grid inside the call box otherwise
    // has no way to actually see shared-screen detail properly.
    videosEl.addEventListener('dblclick', function (event) {
        const video = event.target.closest('video');
        if (!video) return;
        if (document.fullscreenElement) {
            document.exitFullscreen();
        } else if (video.requestFullscreen) {
            video.requestFullscreen();
        }
    });

    function connectCallSocket(roomId) {
        const socket = new WebSocket(window.ChatApp.wsUrl('/ws/chat/call/' + roomId + '/'));
        socket.onmessage = function (event) {
            handleSignal(JSON.parse(event.data));
        };
        call.socket = socket;
        return new Promise(function (resolve) {
            socket.onopen = function () { resolve(socket); };
        });
    }

    function send(payload) {
        if (call.socket && call.socket.readyState === WebSocket.OPEN) {
            call.socket.send(JSON.stringify(payload));
        }
    }

    function describeMediaError(err) {
        if (err && err.code === 'insecure-context') {
            return 'This page must be loaded over HTTPS to use the camera/microphone. Ask an admin for the https:// link to this ERP.';
        }
        const name = err && err.name;
        if (name === 'NotAllowedError' || name === 'PermissionDeniedError') {
            return 'Camera/microphone access was denied. Check this site\'s permissions in your browser settings and try again.';
        }
        if (name === 'NotFoundError' || name === 'OverconstrainedError') {
            return 'No camera/microphone was found on this device.';
        }
        if (name === 'NotReadableError' || name === 'TrackStartError') {
            return 'Your camera/microphone is already in use by another application.';
        }
        return 'Could not access microphone/camera.';
    }

    async function getLocalStream(callType) {
        if (!window.isSecureContext || !navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
            const err = new Error('Insecure context: camera/microphone unavailable.');
            err.code = 'insecure-context';
            throw err;
        }
        if (callType === 'video') {
            try {
                call.localStream = await navigator.mediaDevices.getUserMedia({ audio: true, video: true });
            } catch (e) {
                // No camera on this particular machine shouldn't kill the whole
                // call — the other participant's camera still works fine and
                // this side can still see them; fall back to audio-only rather
                // than failing outright. A truly missing microphone (the only
                // hard requirement) still throws normally below.
                call.localStream = await navigator.mediaDevices.getUserMedia({ audio: true, video: false });
                alert('No camera found on this device — joining with audio only. The other participant\'s video will still work.');
            }
        } else {
            call.localStream = await navigator.mediaDevices.getUserMedia({ audio: true, video: false });
        }
        addVideoEl(currentUserId, call.localStream, true);
        return call.localStream;
    }

    function createPeerConnection(remoteUserId) {
        const pc = new RTCPeerConnection({ iceServers: iceServers });
        call.localStream.getTracks().forEach(function (track) {
            pc.addTrack(track, call.localStream);
        });
        pc.ontrack = function (event) {
            addVideoEl(remoteUserId, event.streams[0], false);
        };
        pc.onicecandidate = function (event) {
            if (event.candidate) {
                send({ event: 'ice-candidate', to_user_id: remoteUserId, candidate: event.candidate, call_id: call.callId });
            }
        };
        pc.onconnectionstatechange = function () { onPeerConnectionStateChange(pc); };
        // Fires whenever a track is added/removed mid-call (e.g. starting
        // screen share on a call that began audio-only) — without this, the
        // new track is attached locally but the already-negotiated remote
        // side never learns about it, so nothing shows up on their end.
        pc.onnegotiationneeded = function () { renegotiate(remoteUserId, pc); };
        call.peers[remoteUserId] = pc;
        return pc;
    }

    async function renegotiate(remoteUserId, pc) {
        // Only for mid-call renegotiation (e.g. adding a screen-share track).
        // The initial offer/answer handshake is driven explicitly by
        // offerTo()/the 'offer' handler below; onnegotiationneeded also fires
        // for that first exchange (adding the initial tracks triggers it),
        // so skip until the call has actually connected once, and only when
        // no offer/answer is already in flight, to avoid a duplicate/racing
        // offer colliding with the explicit initial handshake.
        if (!call.connectedAt || pc.signalingState !== 'stable') return;
        try {
            const offer = await pc.createOffer();
            await pc.setLocalDescription(offer);
            send({ event: 'offer', to_user_id: remoteUserId, sdp: offer, call_id: call.callId, call_type: call.callType });
        } catch (e) { /* ignore — a subsequent negotiationneeded will retry */ }
    }

    async function offerTo(remoteUserId) {
        const pc = createPeerConnection(remoteUserId);
        const offer = await pc.createOffer();
        await pc.setLocalDescription(offer);
        send({ event: 'offer', to_user_id: remoteUserId, sdp: offer, call_id: call.callId, call_type: call.callType });
    }

    async function handleSignal(payload) {
        const fromUserId = payload.from_user_id;

        if (payload.event === 'offer' && (payload.to_user_id === undefined || payload.to_user_id === currentUserId)) {
            call.callId = payload.call_id || call.callId;
            // Reuse the existing connection for this peer if one's already up
            // (a mid-call renegotiation, e.g. screen share starting) — creating
            // a fresh RTCPeerConnection here would tear down the working call
            // instead of just adding the new track to it.
            const pc = call.peers[fromUserId] || createPeerConnection(fromUserId);
            await pc.setRemoteDescription(new RTCSessionDescription(payload.sdp));
            const answer = await pc.createAnswer();
            await pc.setLocalDescription(answer);
            send({ event: 'answer', to_user_id: fromUserId, sdp: answer, call_id: call.callId });
        } else if (payload.event === 'answer' && payload.to_user_id === currentUserId) {
            const pc = call.peers[fromUserId];
            if (pc) await pc.setRemoteDescription(new RTCSessionDescription(payload.sdp));
        } else if (payload.event === 'ice-candidate' && payload.to_user_id === currentUserId) {
            const pc = call.peers[fromUserId];
            if (pc && payload.candidate) {
                try { await pc.addIceCandidate(new RTCIceCandidate(payload.candidate)); } catch (e) { /* ignore */ }
            }
        } else if (payload.event === 'call-ready' && !call.isIncoming && payload.to_user_id === currentUserId) {
            // A callee has joined the call socket and is ready to receive our offer.
            // v1 limitation: only the original caller initiates offers, so group
            // calls form a star of connections through the caller rather than a
            // full mesh between every pair — acceptable at the ~3-4 participant
            // scale expected on this LAN deployment.
            if (window.ChatSound) window.ChatSound.stopRingtone();
            showOverlay('Connecting…');
            offerTo(fromUserId);
        } else if (payload.event === 'hangup' || payload.event === 'call-decline') {
            endCall(false);
        }
    }

    async function startCall(roomId, callType) {
        if (!canInitiateCall) {
            alert('You do not have permission to start calls.');
            return;
        }
        call.active = true;
        call.roomId = roomId;
        call.callType = callType;
        call.isIncoming = false;

        // For a DM, the other participant being online means they'll actually
        // hear this ring live ("Ringing…"); offline just means the call
        // request is queued for whenever they next open the app ("Calling…").
        // Group calls have no single peer to check, so keep the generic label.
        const roomDetail = window.ChatApp.getCurrentRoomDetail && window.ChatApp.getCurrentRoomDetail();
        let statusLabel = 'Calling…';
        if (roomDetail && roomDetail.room_type === 'dm' && roomDetail.other_user_id) {
            statusLabel = window.ChatApp.isUserOnline(roomDetail.other_user_id) ? 'Ringing…' : 'Calling…';
        }
        showOverlay(statusLabel);
        declineLabel.textContent = 'Cancel';
        muteBtn.hidden = false;
        screenshareBtn.hidden = false;

        try {
            await getLocalStream(callType);
        } catch (e) {
            alert(describeMediaError(e));
            endCall(false);
            return;
        }

        await connectCallSocket(roomId);
        send({ event: 'call-invite', call_type: callType });
        if (window.ChatSound) window.ChatSound.playRingtone();

        // Offer to any participant who's already on the room's call socket
        // (best-effort mesh join; primary target is the direct-message peer).
    }

    async function acceptIncomingCall() {
        if (!call.isIncoming) return;
        if (window.ChatSound) window.ChatSound.stopRingtone();
        showOverlay('Connecting…');
        acceptBtn.hidden = true;
        muteBtn.hidden = false;
        screenshareBtn.hidden = false;

        try {
            await getLocalStream(call.callType);
        } catch (e) {
            alert(describeMediaError(e));
            endCall(true);
            return;
        }

        await connectCallSocket(call.roomId);
        send({ event: 'call-ready', call_id: call.callId, to_user_id: call.fromUserId });
    }

    function endCall(sendDecline) {
        if (window.ChatSound) window.ChatSound.stopRingtone();
        if (call.socket) {
            send({ event: sendDecline ? 'call-decline' : 'hangup', call_id: call.callId });
            call.socket.close();
        }
        Object.values(call.peers).forEach(function (pc) { pc.close(); });
        if (call.localStream) {
            call.localStream.getTracks().forEach(function (t) { t.stop(); });
        }
        if (call.screenStream) {
            call.screenStream.getTracks().forEach(function (t) { t.stop(); });
        }
        call.active = false;
        call.roomId = null;
        call.callId = null;
        call.socket = null;
        call.peers = {};
        call.localStream = null;
        call.screenStream = null;
        call.cameraTrack = null;
        resetCallUI();
    }

    // ---- Screen sharing ----------------------------------------------------

    function replaceOutgoingVideoTrack(newTrack) {
        Object.values(call.peers).forEach(function (pc) {
            const sender = pc.getSenders().find(function (s) { return s.track && s.track.kind === 'video'; });
            if (sender) {
                sender.replaceTrack(newTrack);
            } else if (newTrack) {
                pc.addTrack(newTrack, call.localStream);
            }
        });
    }

    async function startScreenShare() {
        if (!window.isSecureContext || !navigator.mediaDevices || !navigator.mediaDevices.getDisplayMedia) {
            alert('This page must be loaded over HTTPS to share your screen.');
            return;
        }
        let stream;
        try {
            stream = await navigator.mediaDevices.getDisplayMedia({ video: true });
        } catch (e) {
            return; // user cancelled the browser's share picker
        }
        const screenTrack = stream.getVideoTracks()[0];
        call.screenStream = stream;
        call.cameraTrack = call.localStream.getVideoTracks()[0] || null; // may be null if call started audio-only

        replaceOutgoingVideoTrack(screenTrack);
        addVideoEl(currentUserId, stream, true);
        screenshareBtn.classList.add('is-active');
        // The whole point of sharing your screen is showing something else —
        // a full-screen modal sitting on top of it defeats that. Shrink to a
        // small corner pill automatically so the rest of the page (or
        // whatever else is on screen) stays visible and usable.
        setMinimized(true);

        screenTrack.addEventListener('ended', function () {
            stopScreenShare();
        });
    }

    function stopScreenShare() {
        if (!call.screenStream) return;
        call.screenStream.getTracks().forEach(function (t) { t.stop(); });
        call.screenStream = null;
        screenshareBtn.classList.remove('is-active');

        replaceOutgoingVideoTrack(call.cameraTrack || null);
        if (call.localStream) addVideoEl(currentUserId, call.localStream, true);
    }

    audioBtn.addEventListener('click', function () {
        if (window.ChatApp.currentRoomId) startCall(window.ChatApp.currentRoomId, 'audio');
    });
    videoBtn.addEventListener('click', function () {
        if (window.ChatApp.currentRoomId) startCall(window.ChatApp.currentRoomId, 'video');
    });
    declineBtn.addEventListener('click', function () { endCall(true); });
    acceptBtn.addEventListener('click', acceptIncomingCall);
    screenshareBtn.addEventListener('click', function () {
        if (call.screenStream) stopScreenShare();
        else startScreenShare();
    });
    muteBtn.addEventListener('click', function () {
        if (!call.localStream) return;
        call.isMuted = !call.isMuted;
        call.localStream.getAudioTracks().forEach(function (t) { t.enabled = !call.isMuted; });
        muteBtn.innerHTML = call.isMuted ? '<i class="fas fa-microphone-slash"></i>' : '<i class="fas fa-microphone"></i>';
    });

    window.ChatApp.onIncomingCall = function (payload) {
        if (call.active) return; // already on a call — a full implementation would queue/reject
        call.active = true;
        call.isIncoming = true;
        call.roomId = payload.room_id;
        call.callId = payload.call_id;
        call.callType = payload.call_type || 'audio';
        call.fromUserId = payload.from_user_id;

        showOverlay('Incoming ' + call.callType + ' call…');
        acceptBtn.hidden = false;
        declineLabel.textContent = 'Decline';
        if (window.ChatSound) window.ChatSound.playRingtone();
    };
})();
