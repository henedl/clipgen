/* clipgen Composer — hub.
 *
 * Owns the shared state (participant/parts/cuts/markers/zoom), the API client,
 * participant selection, the <video> player with multi-part source switching
 * (forked from transcripts-video.js), the keyboard handler + shortcuts popover,
 * the cut list panel, and clip generation through Studio's streaming
 * /api/generate-intake endpoint. Satellites (composer-markers.js,
 * composer-timeline.js) share state through window.ClipgenComposer (CO) and
 * publish their entry points back; the hub calls them through same-named
 * guarded delegators so its own call sites stay plain. Ambient utils.js
 * globals (qs/el/formatTime/formatDuration/clamp/showToast/clipgenApplyConfig/
 * getCanvasThemeColors) are reached via the scope chain.
 */
(function () {
  "use strict";

  var audioPanel = null; // ClipgenVideoControls audio-popover controller

  // Playback-speed cycle. Coarse like Screenspace's (not Transcripts' fine
  // 0.75–2x list): Composer plays whole session recordings, where the point of
  // the control is skimming rather than close listening.
  var VIDEO_SPEEDS = [0.5, 1, 2, 3, 5];

  var state = {
    participants: [],       // [{id, parts:[{name,duration,offset}], total_duration}]
    hasSheet: false,        // a spreadsheet is loaded — gates the off-sheet suffix
    participant: null,      // active participant id
    parts: [],              // active participant's parts
    duration: 0,            // stitched total (seconds)
    activePart: 0,
    playhead: 0,            // global seconds (kept current by timeupdate)
    playing: false,
    videoMuted: false,      // player mute state (icon + video.muted)
    videoPlaybackRate: 1,   // playback speed (session-only, as in Screenspace/Transcripts)
    cuts: [],               // all cuts from the composer manifest (all participants)
    trims: {},              // marker key → {start, end} span overrides
    annotations: [],        // all annotation records (all participants)
    selectedAnnotationIds: [], // multi-select: ids of the selected annotations
    annTool: "select",      // "select" | "text" | "draw"
    annColor: "",           // filled from CLIPGEN_CONFIG at boot
    annStrokeWidth: 0,      // default stroke width for new shapes/strokes (boot)
    annStrokeStyle: "solid", // default stroke style: solid | dashed | dotted (boot)
    annHidden: false,       // hide the annotation layer (B: hold to peek, tap to toggle)
    selectedCutId: null,
    pendingIn: null,        // in-point awaiting its out-point (global seconds)
    markers: { sheet: [], screenspace: [], transcript: [] },
    sourceToggles: { sheet: true, screenspace: true, transcript: true },
    laneFolds: { sheet: true, screenspace: true, transcript: true, annotations: true },
    markerThumbnails: false, // thumbnail strips in marker/cut bars (persisted ui)
    markerAudioScrub: false, // hover audio scrub + waveform on bars (persisted ui)
    followPlayhead: true,   // pan the zoomed timeline to keep the playhead in view (persisted ui)
    sidebarTab: "cuts",     // "cuts" | "sheet" | "screenspace" | "transcript"
    sidebarOpen: true,      // right timelines panel collapsed/expanded (localStorage)
    zoom: 1,
    offset: 0,              // timeline pan offset (global seconds)
    dragging: false,        // timeline satellite sets during pan/edge drags
    generating: false,
    artifactLog: [],        // session record of generate results (TopNav log)
  };

  var CO = { state: state };
  window.ClipgenComposer = CO;

  // ---- Satellite delegators (guarded; satellites load after the hub) ----

  function renderTimeline() { return CO.renderTimeline && CO.renderTimeline.apply(null, arguments); }
  function renderPlayhead() { return CO.renderPlayhead && CO.renderPlayhead.apply(null, arguments); }
  function revealTime() { return CO.revealTime && CO.revealTime.apply(null, arguments); }
  function initTimeline() { return CO.initTimeline && CO.initTimeline.apply(null, arguments); }
  function initMarkerToggles() { return CO.initMarkerToggles && CO.initMarkerToggles.apply(null, arguments); }
  function loadMarkers() { return CO.loadMarkers && CO.loadMarkers.apply(null, arguments); }
  function initAnnotate() { return CO.initAnnotate && CO.initAnnotate.apply(null, arguments); }
  function renderAnnotations() { return CO.renderAnnotations && CO.renderAnnotations.apply(null, arguments); }
  function setAnnotateTool() { return CO.setAnnotateTool && CO.setAnnotateTool.apply(null, arguments); }
  function initMarkerScrub() { return CO.initMarkerScrub && CO.initMarkerScrub.apply(null, arguments); }

  // ---- Multi-part video (forked from transcripts-video.js) ----
  // The <video> plays one part at a time; these helpers present a single
  // GLOBAL timeline so the playhead, cuts, and markers all use global seconds.

  function partForGlobal(g) {
    var parts = state.parts;
    for (var i = 0; i < parts.length; i++) {
      if (g >= parts[i].offset && g < parts[i].offset + parts[i].duration) return i;
    }
    return Math.max(0, parts.length - 1);
  }

  function videoGlobalTime() {
    var video = qs("#coVideo");
    var part = state.parts[state.activePart];
    return (video ? video.currentTime : 0) + (part ? part.offset : 0);
  }
  CO.videoGlobalTime = videoGlobalTime;

  // Deferred/coalesced seek scaffolding (createSeekCoalescer, utils.js).
  var _seek = createSeekCoalescer(
    function () { return qs("#coVideo"); },
    function (t) { seekLocal(t); },
    function (video, t) { video.currentTime = t; }
  );

  function cancelPendingSeek() {
    _seek.cancel();
  }

  function seekVideo(time) {
    // *time* is GLOBAL; switch parts when needed, then seek the local offset.
    if (!state.parts.length) return;
    var g = clamp(time, 0, Math.max(0, state.duration - 0.001));
    var i = partForGlobal(g);
    var local = g - state.parts[i].offset;
    state.playhead = g;
    if (i !== state.activePart) {
      switchToPart(i, local, state.playing);
    } else {
      seekLocal(local);
    }
    revealTime(g); // pan the zoomed camera to the seek target before drawing
    renderPlayhead();
    renderAnnotations();
    updateTimeLabel();
  }
  CO.seekVideo = seekVideo;

  function seekLocal(time) {
    _seek.seek(time);
  }

  function switchToPart(i, localTime, resume) {
    var video = qs("#coVideo");
    if (!video) return;
    // A same-part seek may still be deferred on loadedmetadata; drop it so it
    // can't fire after this part loads and clobber this cross-part seek.
    cancelPendingSeek();
    state.activePart = i;
    video.src = "media/" + encodeURIComponent(state.parts[i].name);
    video.load();
    var onMeta = function () {
      video.removeEventListener("loadedmetadata", onMeta);
      video.currentTime = localTime;
      applyPlaybackRate(); // load() reset the element to its default rate
      if (resume) video.play();
    };
    video.addEventListener("loadedmetadata", onMeta);
  }

  function togglePlay() {
    var video = qs("#coVideo");
    if (!video || !video.src) return;
    if (video.paused) video.play();
    else video.pause();
  }

  function updatePlayButton() {
    var icon = qs("#coPlayIcon");
    if (icon) icon.className = "co-btn-icon " + (state.playing ? "co-icon-pause" : "co-icon-play");
  }

  function updateMuteButton() {
    var icon = qs("#coMuteIcon");
    if (icon) icon.className = "co-btn-icon " + (state.videoMuted ? "co-icon-mute-off" : "co-icon-mute");
    // Accent tint while sound is on — matches Screenspace/Transcripts mute btns.
    var btn = qs("#coMuteBtn");
    if (btn) btn.classList.toggle("active", !state.videoMuted);
  }

  function applyPlaybackRate() {
    window.ClipgenVideoControls.applyPlaybackRate(qs("#coVideo"), state.videoPlaybackRate);
  }

  function updateSpeedButton() {
    var btn = qs("#coSpeedBtn");
    if (!btn) return;
    btn.textContent = state.videoPlaybackRate + "x";
    btn.classList.toggle("active", state.videoPlaybackRate !== 1);
  }

  function updateTimeLabel() {
    var label = qs("#coTimeLabel");
    if (!label) return;
    label.textContent = formatTime(state.playhead, { decimals: 0 }) +
      " / " + formatTime(state.duration, { decimals: 0 });
  }

  // The footer hint advertises double-click cuts only while the setting is on
  // (config at boot; the Settings modal's onSave re-syncs it live).
  function updateTimelineHint() {
    var hint = qs(".co-timeline-hint");
    if (!hint) return;
    hint.textContent = "scroll to zoom · drag to pan · drag cut edges to trim" +
      (CLIPGEN_CONFIG.composerDoubleClickCuts
        ? " · double-click to set in/out"
        : "");
  }

  // Subheader source readout: duration · resolution · fps (Screenspace's
  // video-info format). Resolution/fps ride along on the participant record;
  // duration may firm up later for single-part participants (loadedmetadata).
  function updateVideoInfo() {
    var elInfo = qs("#coVideoInfo");
    if (!elInfo) return;
    var p = findParticipant(state.participant);
    if (!p) { elInfo.textContent = ""; return; }
    var parts = [];
    if (state.duration) parts.push(formatDuration(state.duration));
    if (p.width && p.height) parts.push(p.width + "x" + p.height);
    if (p.fps) parts.push(Math.round(p.fps) + "fps");
    elInfo.textContent = parts.join(" · ");
  }

  function initVideo() {
    var video = qs("#coVideo");
    video.addEventListener("play", function () {
      state.playing = true;
      updatePlayButton();
    });
    video.addEventListener("pause", function () {
      state.playing = false;
      updatePlayButton();
    });
    video.addEventListener("ended", function () {
      // A part can end before timeupdate ever lands inside the hand-off window
      // below — timeupdate fires ~4x/s, so at 5x the element jumps ~1.25 s per
      // event and steps clean over it. Hand off here too, or fast playback
      // stalls at every part boundary.
      if (state.activePart < state.parts.length - 1) {
        switchToPart(state.activePart + 1, 0.001, true);
        return;
      }
      state.playing = false;
      updatePlayButton();
    });
    var _playheadRaf = 0;
    video.addEventListener("timeupdate", function () {
      // Hand off to the next part as playback nears the boundary.
      var i = state.activePart;
      var parts = state.parts;
      if (parts.length > 1 && i < parts.length - 1 &&
          video.currentTime >= parts[i].duration - 0.05) {
        switchToPart(i + 1, 0.001, !video.paused);
      }
      state.playhead = videoGlobalTime();
      updateTimeLabel();
      if (_playheadRaf) return;
      _playheadRaf = requestAnimationFrame(function () {
        _playheadRaf = 0;
        if (state.followPlayhead) revealTime(state.playhead);
        renderPlayhead();
        renderAnnotations(); // spans gate visibility against the playhead
      });
    });
    // Single-part participants may have no probed duration (server couldn't
    // ffprobe) — fall back to the element's own metadata.
    video.addEventListener("loadedmetadata", function () {
      if (!state.duration && state.parts.length === 1 && isFinite(video.duration)) {
        state.duration = video.duration;
        state.parts[0].duration = video.duration;
        updateTimeLabel();
        updateVideoInfo();
        renderTimeline();
      }
    });
    qs("#coPlayBtn").addEventListener("click", togglePlay);

    qs("#coSpeedBtn").addEventListener("click", function () {
      state.videoPlaybackRate = window.ClipgenVideoControls.nextSpeed(
        VIDEO_SPEEDS, state.videoPlaybackRate);
      applyPlaybackRate();
      updateSpeedButton();
    });

    // Hover the mute button for a glassy 0–200% volume popover (click still
    // mutes). getTracks reads the active participant's probed audio layout;
    // trackAudioUrl enables per-track mixing for single-file participants.
    var muteBtn = qs("#coMuteBtn");
    audioPanel = window.ClipgenVideoControls.attachAudioPanel({
      video: video,
      button: muteBtn,
      getTracks: function () {
        var p = findParticipant(state.participant);
        return (p && p.audio_tracks) || [];
      },
      trackAudioUrl: function (idx) {
        if (!state.participant || state.parts.length > 1) return null;
        return "api/audio-track/" + encodeURIComponent(state.participant) + "/" + idx;
      },
    });
    muteBtn.addEventListener("click", function () {
      state.videoMuted = !state.videoMuted;
      if (audioPanel) audioPanel.setMuted(state.videoMuted);
      else video.muted = state.videoMuted;
      updateMuteButton();
    });
    updateMuteButton();
    updateSpeedButton();
  }

  // ---- Participant selection ----

  function findParticipant(pid) {
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) return state.participants[i];
    }
    return null;
  }

  function selectParticipant(pid) {
    var p = findParticipant(pid);
    if (!p) return;
    // Fragmented-MP4 warning + remux action, above the timeline.
    if (window.clipgenMediaBanner) {
      window.clipgenMediaBanner.show(qs("#coTimelineSection"), p);
    }
    cancelPendingSeek();
    state.participant = pid;
    setStoredUIStateField("composer", "participant", pid);
    state.parts = p.parts || [];
    state.duration = p.total_duration || 0;
    state.activePart = 0;
    state.playhead = 0;
    state.playing = false;
    state.pendingIn = null;
    state.selectedCutId = null;
    state.selectedAnnotationIds = [];
    state.zoom = 1;
    state.offset = 0;
    state.markers = { sheet: [], screenspace: [], transcript: [] };
    // Drop cached sprites/audio + stop any playing snippet: they belong to the
    // previous participant's video.
    if (CO.resetScrubMedia) CO.resetScrubMedia();
    // Undo ops reference the previous participant's cuts — an undo fired after
    // a switch would invisibly mutate that other timeline. Drop the stacks.
    _undoStack.length = 0;
    _redoStack.length = 0;
    syncUndoButtons();

    var video = qs("#coVideo");
    video.pause();
    if (state.parts.length) {
      video.src = "media/" + encodeURIComponent(state.parts[0].name);
      video.load();
      applyPlaybackRate(); // load() reset the element to its default rate
      qs("#coVideoFrame").classList.add("has-video");
    } else {
      video.removeAttribute("src");
      qs("#coVideoFrame").classList.remove("has-video");
    }
    // Reconfigure the audio popover (per-track mixer vs master slider) and tear
    // down the previous participant's track mix for this new participant.
    if (audioPanel) audioPanel.refresh();

    ["#coPlayBtn", "#coMuteBtn", "#coSpeedBtn", "#coSetInBtn", "#coSetOutBtn"].forEach(function (sel) {
      qs(sel).disabled = false;
    });
    qs("#coAnnotateBar").classList.remove("hidden");
    state.annTool = "select";
    setAnnotateTool("select"); // also closes a pending text input
    updatePendingInfo();
    updatePlayButton();
    updateTimeLabel();
    updateVideoInfo();
    updateGenerateButton();
    renderSidebar();
    if (CO.updateTimelineHeight) CO.updateTimelineHeight();
    renderTimeline();
    renderAnnotations();
    loadMarkers(pid);
  }

  function initParticipantSelect() {
    var select = qs("#coParticipantSelect");
    select.addEventListener("change", function () {
      if (select.value) selectParticipant(select.value);
    });
  }

  function populateParticipantSelect() {
    var select = qs("#coParticipantSelect");
    select.innerHTML = "";
    var placeholder = el("option", "", "Select…");
    placeholder.value = "";
    select.appendChild(placeholder);
    state.participants.forEach(function (p) {
      // An <option> can't hold a badge element, so an off-sheet participant —
      // video on disk, no column in the loaded sheet — gets a label suffix.
      // Only when a sheet is loaded: without one everything is off-sheet.
      var label = state.hasSheet && p.in_sheet === false ? p.id + " (off-sheet)" : p.id;
      var opt = el("option", "", label);
      opt.value = p.id;
      select.appendChild(opt);
    });
    select.disabled = state.participants.length === 0;
  }

  // ---- Cuts ----

  function participantCuts() {
    return state.cuts.filter(function (c) { return c.participant === state.participant; });
  }
  CO.participantCuts = participantCuts;

  // Chronological order — the cut list renders in this order and position+1 is
  // the cut's index badge (list and timeline both number from here).
  function sortedCuts() {
    return participantCuts().slice().sort(function (a, b) { return a.start - b.start; });
  }
  CO.sortedCuts = sortedCuts;

  function findCut(id) {
    for (var i = 0; i < state.cuts.length; i++) {
      if (state.cuts[i].id === id) return state.cuts[i];
    }
    return null;
  }
  CO.findCut = findCut;

  function updatePendingInfo() {
    var info = qs("#coPendingInfo");
    info.textContent = state.pendingIn === null
      ? ""
      : "In: " + formatTime(state.pendingIn, { decimals: 1 }) + ", press O";
  }

  function setInPoint() {
    if (!state.participant) return;
    state.pendingIn = state.playhead;
    updatePendingInfo();
    renderPlayhead();
  }
  CO.setInPoint = setInPoint;

  // Raw appliers — perform the API call + local state update, no undo
  // recording. User actions wrap these and record an op; undo/redo replay
  // them directly (recording again would corrupt the stacks).

  function refreshCutViews() {
    updateGenerateButton();
    renderCutList();
    renderTimeline();
  }

  function applyCreate(cutData) {
    return apiPost("api/cuts", {
      participant: cutData.participant,
      start: cutData.start,
      end: cutData.end,
      label: cutData.label || "",
    }).then(function (data) {
      if (!data.ok) throw new Error(data.error || "Could not save cut");
      state.cuts.push(data.cut);
      refreshCutViews();
      return data.cut;
    });
  }

  function applyDelete(id) {
    return apiDelete("api/cuts/" + encodeURIComponent(id)).then(function (data) {
      if (!data.ok) throw new Error(data.error || "Could not delete cut");
      state.cuts = state.cuts.filter(function (c) { return c.id !== id; });
      if (state.selectedCutId === id) state.selectedCutId = null;
      refreshCutViews();
    });
  }

  function applyTimes(id, times) {
    return apiPatch("api/cuts/" + encodeURIComponent(id), {
      start: times.start,
      end: times.end,
    }).then(function (data) {
      if (!data.ok) throw new Error(data.error || "Could not save cut");
      var cut = findCut(id);
      if (cut) {
        cut.start = data.cut.start;
        cut.end = data.cut.end;
      }
      refreshCutViews();
      return data.cut;
    });
  }

  function opFailed(err) {
    showToast(err && err.message ? err.message : "Cut update failed");
  }

  // Find a loaded marker by key across the three lanes (markers are rebuilt on
  // participant switch, so a trim op can outlive its marker — that's fine, the
  // manifest still updates and the next lane load reflects it).
  function findMarker(key) {
    var sources = Object.keys(state.markers);
    for (var s = 0; s < sources.length; s++) {
      var lane = state.markers[sources[s]];
      for (var i = 0; i < lane.length; i++) {
        if (lane[i].key === key) return lane[i];
      }
    }
    return null;
  }

  function refreshMarkerViews() {
    if (CO.updateTimelineHeight) CO.updateTimelineHeight();
    renderTimeline();
    renderSidebar();
  }

  // Raw trim applier: *values* = {start, end} sets the override, null resets
  // it. Updates the manifest, the local trims map, and the loaded marker.
  // Marker metadata rides along when the marker is loaded so Studio's
  // Composer Intake can render the trim as a card; the server preserves
  // previously stored metadata when a re-PUT (undo/redo) omits it.
  function applyTrim(key, values, sourceSpan) {
    var payload = values ? { start: values.start, end: values.end } : null;
    var meta = values && findMarker(key);
    if (meta) {
      payload.participant = state.participant;
      payload.label = meta.label || meta.eventType || "";
      payload.source = meta.source;
    }
    var call = payload
      ? apiPut("api/trims/" + encodeURIComponent(key), payload)
      : apiDelete("api/trims/" + encodeURIComponent(key));
    return call.then(function (data) {
      if (!data.ok) throw new Error(data.error || "Could not save trim");
      var marker = findMarker(key);
      if (values) {
        state.trims[key] = { start: data.trim.start, end: data.trim.end };
        if (marker) {
          if (!marker.trimmed) {
            // A drag-trim mutates marker.start/end live, so the current span is
            // already the trimmed value; *sourceSpan* (the pre-drag span passed
            // by commitMarkerTrim) is the true original. Non-drag callers omit it
            // and fall back to the current span, which hasn't been mutated.
            var src = sourceSpan || { start: marker.start, end: marker.end };
            marker.origStart = src.start;
            marker.origEnd = src.end;
          }
          marker.trimmed = true;
          marker.start = data.trim.start;
          marker.end = data.trim.end;
        }
      } else {
        delete state.trims[key];
        if (marker && marker.trimmed) {
          marker.start = marker.origStart;
          marker.end = marker.origEnd;
          marker.trimmed = false;
        }
      }
      refreshMarkerViews();
    });
  }

  // ---- Undo / redo (cut create / delete / edit) ----

  var _undoStack = [];
  var _redoStack = [];
  var UNDO_LIMIT = 100;

  function syncUndoButtons() {
    qs("#coUndoBtn").disabled = _undoStack.length === 0;
    qs("#coRedoBtn").disabled = _redoStack.length === 0;
  }

  function recordOp(op) {
    _undoStack.push(op);
    if (_undoStack.length > UNDO_LIMIT) _undoStack.shift();
    _redoStack.length = 0;
    syncUndoButtons();
  }

  // Apply *op* in the given direction. A re-created cut gets a fresh server id,
  // so the op's stored cut is swapped for the new one — the paired redo/undo
  // then targets the id that actually exists.
  function applyOp(op, isUndo) {
    if (op.type === "create") {
      return isUndo
        ? applyDelete(op.cut.id)
        : applyCreate(op.cut).then(function (cut) { op.cut = cut; });
    }
    if (op.type === "delete") {
      return isUndo
        ? applyCreate(op.cut).then(function (cut) { op.cut = cut; })
        : applyDelete(op.cut.id);
    }
    if (op.type === "trim") {
      return applyTrim(op.key, isUndo ? op.before : op.after);
    }
    if (op.type === "ann-create") {
      return isUndo
        ? applyAnnDelete(op.annotation.id)
        : applyAnnCreate(op.annotation).then(function (ann) { op.annotation = ann; });
    }
    if (op.type === "ann-delete") {
      return isUndo
        ? applyAnnCreate(op.annotation).then(function (ann) { op.annotation = ann; })
        : applyAnnDelete(op.annotation.id);
    }
    if (op.type === "ann-edit") {
      var payload = {};
      payload[op.field] = isUndo ? op.before : op.after;
      return applyAnnPatch(op.id, payload);
    }
    if (op.type === "ann-group") {
      // Sub-ops are independent per-annotation edits/deletes; apply them all in
      // the same direction (order-independent) and resolve when the last lands.
      return Promise.all(op.ops.map(function (sub) { return applyOp(sub, isUndo); }));
    }
    // edit
    return applyTimes(op.id, isUndo ? op.before : op.after);
  }

  // Peek-apply-pop: the op moves between stacks only once the server has
  // accepted it, so a failed request keeps the op available for retry (the
  // raw appliers never mutate local state on failure). The busy flag stops a
  // rapid second ⌘Z from re-applying the still-peeked op.
  var _historyBusy = false;

  function shiftHistory(fromStack, toStack, isUndo) {
    if (_historyBusy) return;
    var op = fromStack[fromStack.length - 1];
    if (!op) return;
    _historyBusy = true;
    applyOp(op, isUndo).then(function () {
      fromStack.pop();
      toStack.push(op);
    }, opFailed).then(function () {
      _historyBusy = false;
      syncUndoButtons();
    });
  }

  function undo() { shiftHistory(_undoStack, _redoStack, true); }

  function redo() { shiftHistory(_redoStack, _undoStack, false); }

  // ---- User-facing cut actions (record undo ops) ----

  function setOutPoint() {
    if (!state.participant || state.pendingIn === null) {
      showToast("Set an in point first (I)");
      return;
    }
    var start = state.pendingIn;
    var end = state.playhead;
    if (end < start) { var tmp = start; start = end; end = tmp; }
    if (end - start < 0.2) {
      showToast("Cut too short. Move the playhead past the in point.");
      return;
    }
    applyCreate({ participant: state.participant, start: start, end: end })
      .then(function (cut) {
        state.pendingIn = null;
        state.selectedCutId = cut.id;
        updatePendingInfo();
        recordOp({ type: "create", cut: cut });
        refreshCutViews();
      })
      .catch(opFailed);
  }
  CO.setOutPoint = setOutPoint;

  function deleteCut(id) {
    var cut = findCut(id);
    if (!cut) return;
    var snapshot = { id: cut.id, participant: cut.participant, start: cut.start, end: cut.end, label: cut.label };
    applyDelete(id).then(function () {
      recordOp({ type: "delete", cut: snapshot });
    }).catch(opFailed);
  }
  CO.deleteCut = deleteCut;

  // Persist a cut's edited span (timeline edge/body drags land here on drag
  // end, with *before* = the pre-drag times so the edit is undoable). On
  // failure the optimistic drag is rolled back so the view matches the server.
  function commitCutTimes(cut, before) {
    var changed = !before || before.start !== cut.start || before.end !== cut.end;
    applyTimes(cut.id, { start: cut.start, end: cut.end }).then(function (saved) {
      if (before && changed) {
        recordOp({
          type: "edit",
          id: cut.id,
          before: before,
          after: { start: saved.start, end: saved.end },
        });
      }
    }).catch(function (error) {
      if (before) {
        cut.start = before.start;
        cut.end = before.end;
        refreshCutViews();
      }
      opFailed(error);
    });
  }
  CO.commitCutTimes = commitCutTimes;

  function selectCut(id) {
    state.selectedCutId = id;
    renderCutList();
    renderTimeline();
  }
  CO.selectCut = selectCut;

  // ---- Annotations (raw appliers + recorded actions) ----

  function participantAnnotations() {
    return state.annotations.filter(function (a) {
      return a.participant === state.participant;
    });
  }
  CO.participantAnnotations = participantAnnotations;

  function findAnnotation(id) {
    for (var i = 0; i < state.annotations.length; i++) {
      if (state.annotations[i].id === id) return state.annotations[i];
    }
    return null;
  }
  CO.findAnnotation = findAnnotation;

  function refreshAnnotationViews() {
    if (CO.updateTimelineHeight) CO.updateTimelineHeight();
    renderTimeline();
    renderAnnotations();
    renderCutList(); // cut items badge overlapping annotations
  }
  CO.refreshAnnotationViews = refreshAnnotationViews;

  function applyAnnCreate(record) {
    return apiPost("api/annotations", record).then(function (data) {
      if (!data.ok) throw new Error(data.error || "Could not save annotation");
      state.annotations.push(data.annotation);
      refreshAnnotationViews();
      return data.annotation;
    });
  }

  function applyAnnDelete(id) {
    return apiDelete("api/annotations/" + encodeURIComponent(id))
      .then(function (data) {
        if (!data.ok) throw new Error(data.error || "Could not delete annotation");
        state.annotations = state.annotations.filter(function (a) { return a.id !== id; });
        state.selectedAnnotationIds = state.selectedAnnotationIds.filter(
          function (sid) { return sid !== id; });
        refreshAnnotationViews();
      });
  }

  function applyAnnPatch(id, fields) {
    return apiPatch("api/annotations/" + encodeURIComponent(id), fields)
      .then(function (data) {
        if (!data.ok) throw new Error(data.error || "Could not save annotation");
        var idx = state.annotations.findIndex(function (a) { return a.id === id; });
        if (idx >= 0) state.annotations[idx] = data.annotation;
        refreshAnnotationViews();
        return data.annotation;
      });
  }

  // Recorded user actions (undo/redo integrated like cuts/trims).

  function createAnnotation(record) {
    return applyAnnCreate(record).then(function (ann) {
      recordOp({ type: "ann-create", annotation: ann });
      state.selectedAnnotationIds = [ann.id];
      renderAnnotations();
      return ann;
    }).catch(opFailed);
  }
  CO.createAnnotation = createAnnotation;

  function deleteAnnotation(id) {
    var ann = findAnnotation(id);
    if (!ann) return;
    var snapshot = JSON.parse(JSON.stringify(ann));
    applyAnnDelete(id).then(function () {
      recordOp({ type: "ann-delete", annotation: snapshot });
    }).catch(opFailed);
  }
  CO.deleteAnnotation = deleteAnnotation;

  // Persist an edited field ("span" or "geometry"); *before* is the pre-edit
  // value for undo. The record was already mutated locally by the caller —
  // rolled back to *before* if the server rejects the edit.
  function commitAnnotationField(ann, field, before) {
    var payload = {};
    payload[field] = ann[field];
    applyAnnPatch(ann.id, payload).then(function (saved) {
      recordOp({
        type: "ann-edit",
        id: ann.id,
        field: field,
        before: before,
        after: JSON.parse(JSON.stringify(saved[field])),
      });
    }).catch(function (error) {
      if (before) {
        ann[field] = JSON.parse(JSON.stringify(before));
        refreshAnnotationViews();
      }
      opFailed(error);
    });
  }
  CO.commitAnnotationField = commitAnnotationField;

  // Batch commit of *field* ("style" or "geometry") across many annotations, each
  // already mutated, with *before* the pre-edit value. One undo step, so a group
  // style change or move undoes atomically; every edit rolls back if any patch is
  // rejected. Each op's *after* comes from the PATCH response (server-sanitized),
  // as in the single-annotation path — a partial local object can't reset
  // backfilled defaults.
  function commitAnnotationFieldGroup(field, edits) {
    if (!edits.length) return Promise.resolve();
    var patches = edits.map(function (e) {
      var payload = {};
      payload[field] = e.ann[field];
      return applyAnnPatch(e.ann.id, payload).then(function (saved) {
        return {
          type: "ann-edit",
          id: e.ann.id,
          field: field,
          before: e.before,
          after: JSON.parse(JSON.stringify(saved[field])),
        };
      });
    });
    return Promise.all(patches).then(function (ops) {
      recordOp(ops.length === 1 ? ops[0] : { type: "ann-group", ops: ops });
    }, function (error) {
      edits.forEach(function (e) {
        var ann = findAnnotation(e.ann.id);
        if (ann) ann[field] = JSON.parse(JSON.stringify(e.before));
      });
      refreshAnnotationViews();
      opFailed(error);
    });
  }
  CO.commitAnnotationFieldGroup = commitAnnotationFieldGroup;

  // Delete every selected annotation as a single undo step.
  function deleteSelectedAnnotations() {
    var snapshots = selectedAnnotations().map(function (ann) {
      return JSON.parse(JSON.stringify(ann));
    });
    if (!snapshots.length) return;
    Promise.all(snapshots.map(function (snap) {
      return applyAnnDelete(snap.id);
    })).then(function () {
      var ops = snapshots.map(function (snap) {
        return { type: "ann-delete", annotation: snap };
      });
      recordOp(ops.length === 1 ? ops[0] : { type: "ann-group", ops: ops });
    }).catch(opFailed);
  }
  CO.deleteSelectedAnnotations = deleteSelectedAnnotations;

  // ---- Selection (multi-select) ----

  function isAnnotationSelected(id) {
    return state.selectedAnnotationIds.indexOf(id) !== -1;
  }
  CO.isAnnotationSelected = isAnnotationSelected;

  // Live records for the current selection, dropping any stale ids.
  function selectedAnnotations() {
    var out = [];
    state.selectedAnnotationIds.forEach(function (id) {
      var ann = findAnnotation(id);
      if (ann) out.push(ann);
    });
    return out;
  }
  CO.selectedAnnotations = selectedAnnotations;

  // The record iff exactly one annotation is selected — gates the single-shape
  // resize/rotate handles (group resize is not supported).
  function singleSelectedAnnotation() {
    var sel = selectedAnnotations();
    return sel.length === 1 ? sel[0] : null;
  }
  CO.singleSelectedAnnotation = singleSelectedAnnotation;

  function setAnnotationSelection(ids) {
    state.selectedAnnotationIds = (ids || []).slice();
    renderTimeline();
    renderAnnotations();
  }
  CO.setAnnotationSelection = setAnnotationSelection;

  // Single-select convenience (kept for click-to-select / clear callers).
  function selectAnnotation(id) {
    setAnnotationSelection(id ? [id] : []);
  }
  CO.selectAnnotation = selectAnnotation;

  function toggleAnnotationSelection(id) {
    if (!id) return;
    setAnnotationSelection(
      isAnnotationSelected(id)
        ? state.selectedAnnotationIds.filter(function (s) { return s !== id; })
        : state.selectedAnnotationIds.concat([id])
    );
  }
  CO.toggleAnnotationSelection = toggleAnnotationSelection;

  // Hide/reveal the whole annotation layer: the overlay skips drawing and the
  // timeline lane dims. The #coToolHide button mirrors the state; its icon
  // shows the action (eye-slash = will hide, eye = will reveal).
  function setAnnotationsHidden(hidden) {
    state.annHidden = !!hidden;
    var btn = qs("#coToolHide");
    if (btn) {
      btn.setAttribute("aria-pressed", state.annHidden ? "true" : "false");
      // No key hint here — the hotkey registry renders its own chip for
      // [data-hotkey] controls on Alt-hold, and a hand-written one goes stale
      // the moment the user rebinds.
      var hideLabel = (state.annHidden ? "Show" : "Hide") + " annotations";
      btn.setAttribute("aria-label", hideLabel);
      btn.setAttribute("data-tooltip", hideLabel + " (hold to peek, tap to toggle)");
      var icon = btn.querySelector(".co-btn-icon");
      if (icon) {
        icon.classList.toggle("co-icon-eye", state.annHidden);
        icon.classList.toggle("co-icon-eye-slash", !state.annHidden);
      }
    }
    renderAnnotations();
    renderTimeline();
  }
  CO.setAnnotationsHidden = setAnnotationsHidden;

  // ---- Annotated exports (server PIL + ffmpeg overlay) ----

  var _exporting = false;

  function exportSpan() {
    // Burn/GIF need a span: the selected cut wins, else the selected
    // annotation's own visibility span.
    var cut = state.selectedCutId && findCut(state.selectedCutId);
    if (cut) return { start: cut.start, end: cut.end };
    var ann = selectedAnnotations()[0];
    if (ann) return { start: ann.span.start, end: ann.span.end };
    return null;
  }

  function runExport(path, body, busyLabel) {
    if (_exporting) { showToast("An export is already running"); return; }
    _exporting = true;
    showToast(busyLabel + "…");
    apiPost(path, body).then(function (data) {
      _exporting = false;
      if (!data.ok) { showToast(data.error || "Export failed"); return; }
      logArtifactResult({ ok: true, artifact: data.artifact }, null);
      showToast("Exported " + (data.artifact.file || ""));
    }).catch(function (err) {
      _exporting = false;
      // A 4xx envelope now rejects; err.message carries the server's text
      // ("An export is already running", "The span is outside the recording").
      showToast(err && err.message ? err.message : "Export failed");
    });
  }

  function exportScreenshot() {
    if (!state.participant) return;
    runExport("api/export/screenshot", {
      participant: state.participant,
      time: state.playhead,
    }, "Exporting screenshot");
  }

  function exportBurn(gif) {
    if (!state.participant) return;
    var span = exportSpan();
    if (!span) {
      showToast("Select a cut (or an annotation) to define the export span");
      return;
    }
    runExport(gif ? "api/export/gif" : "api/export/burn", {
      participant: state.participant,
      start: span.start,
      end: span.end,
    }, gif ? "Exporting GIF" : "Burning clip");
  }

  // ---- Marker trims (user actions; non-destructive span overrides) ----

  // Persist a marker's dragged span. The timeline mutated the marker live;
  // *before* is the trim that was in force pre-drag (null = untrimmed) and
  // *origSpan* the marker's pre-drag visual span (for failure rollback).
  function commitMarkerTrim(marker, before, origSpan) {
    var after = { start: marker.start, end: marker.end };
    applyTrim(marker.key, after, origSpan).then(function () {
      recordOp({ type: "trim", key: marker.key, before: before, after: after });
    }).catch(function (error) {
      if (origSpan) {
        marker.start = origSpan.start;
        marker.end = origSpan.end;
        renderTimeline();
      }
      opFailed(error);
    });
  }
  CO.commitMarkerTrim = commitMarkerTrim;

  function resetTrim(marker) {
    var before = state.trims[marker.key];
    if (!before) return;
    applyTrim(marker.key, null).then(function () {
      recordOp({
        type: "trim",
        key: marker.key,
        before: { start: before.start, end: before.end },
        after: null,
      });
    }).catch(opFailed);
  }
  CO.resetTrim = resetTrim;

  // Promote a marker's (possibly trimmed) span to a Composer cut so it feeds
  // generation without new plumbing.
  function copyMarkerToCut(marker) {
    applyCreate({
      participant: state.participant,
      start: marker.start,
      end: marker.end,
      label: marker.label || marker.eventType || "",
    }).then(function (cut) {
      recordOp({ type: "create", cut: cut });
      state.selectedCutId = cut.id;
      showToast("Copied to cuts");
      refreshCutViews();
    }).catch(opFailed);
  }
  CO.copyMarkerToCut = copyMarkerToCut;

  function nudgeSelectedCut(deltaSeconds) {
    var cut = state.selectedCutId && findCut(state.selectedCutId);
    if (!cut) return;
    var before = { start: cut.start, end: cut.end };
    // Nudge the out edge; hold the in edge fixed (the common trim gesture).
    cut.end = clamp(cut.end + deltaSeconds, cut.start + 0.2, state.duration || cut.end + deltaSeconds);
    renderTimeline();
    commitCutTimes(cut, before);
  }

  function commitCutLabel(cut, label) {
    if (label === (cut.label || "")) return;
    apiPatch("api/cuts/" + encodeURIComponent(cut.id), { label: label })
      .then(function (data) {
        if (!data.ok) throw new Error(data.error || "Could not save name");
        cut.label = data.cut.label;
      })
      .catch(opFailed);
  }

  function renderCutList() {
    if (state.sidebarTab !== "cuts") return;
    var list = qs("#coCutList");
    list.innerHTML = "";
    var cuts = sortedCuts();
    if (!cuts.length) {
      var hint = el("p", "co-empty-hint");
      hint.innerHTML = "Press <kbd>I</kbd> then <kbd>O</kbd> to add a cut.";
      list.appendChild(hint);
      return;
    }
    var anns = participantAnnotations();
    var frag = document.createDocumentFragment();
    cuts.forEach(function (cut, i) {
      var item = el("div", "co-cut-item" + (cut.id === state.selectedCutId ? " selected" : ""));

      // Row 1: chronological index + editable name (carried into generated
      // clips as the event type; also feeds the Studio Composer-intake tab's
      // cards) + delete button.
      var nameRow = el("div", "co-cut-row");
      nameRow.appendChild(el("span", "co-cut-index", String(i + 1)));
      var name = el("input", "co-cut-name");
      name.type = "text";
      name.autocomplete = "off";
      name.placeholder = "Name this cut…";
      name.value = cut.label || "";
      name.setAttribute("aria-label", "Cut name");
      name.addEventListener("click", function (e) { e.stopPropagation(); });
      name.addEventListener("keydown", function (e) {
        if (e.key === "Enter") name.blur();
        e.stopPropagation();
      });
      name.addEventListener("blur", function () {
        commitCutLabel(cut, name.value.trim());
      });
      nameRow.appendChild(name);
      var hasAnn = anns.some(function (a) {
        return a.span.start < cut.end && a.span.end > cut.start;
      });
      if (hasAnn) {
        var annBadge = el("span", "co-btn-icon co-icon-draw co-cut-ann-badge");
        annBadge.setAttribute("data-tooltip", "Has annotations");
        nameRow.appendChild(annBadge);
      }
      if (cut._genStatus) {
        var okStatus = cut._genStatus === "ok";
        var statusIcon = el("span", "co-cut-status co-btn-icon " + cut._genStatus +
          (okStatus ? " co-icon-check" : " co-icon-x-mark"));
        statusIcon.setAttribute("data-tooltip",
          okStatus ? "Generated" : "Generation failed");
        nameRow.appendChild(statusIcon);
      }
      var del = el("button", "co-cut-delete");
      del.type = "button";
      del.setAttribute("data-tooltip", "Delete cut");
      del.setAttribute("aria-label", "Delete cut");
      del.appendChild(el("span", "co-btn-icon co-icon-trash"));
      del.addEventListener("click", function (e) {
        e.stopPropagation();
        deleteCut(cut.id);
      });
      nameRow.appendChild(del);
      item.appendChild(nameRow);

      // Row 2: span + duration.
      var timeRow = el("div", "co-cut-row");
      timeRow.appendChild(el("span", "co-cut-times",
        formatTime(cut.start, { decimals: 1 }) + " – " + formatTime(cut.end, { decimals: 1 })));
      timeRow.appendChild(el("span", "co-cut-dur", formatDuration(cut.end - cut.start)));
      item.appendChild(timeRow);

      item.addEventListener("click", function () {
        selectCut(cut.id);
        seekVideo(cut.start);
      });
      frag.appendChild(item);
    });
    list.appendChild(frag);
  }
  CO.renderCutList = renderCutList;

  // ---- Sidebar tabs (cuts + one list per marker source) ----

  function renderMarkerList(source) {
    var list = qs("#coCutList");
    list.innerHTML = "";
    var markers = (state.markers[source] || []).slice()
      .sort(function (a, b) { return a.start - b.start; });
    if (!markers.length) {
      list.appendChild(el("p", "co-empty-hint",
        "No " + source + " markers for this participant."));
      return;
    }
    var color = getCSSVar("--stream-" + source, "#888");
    var frag = document.createDocumentFragment();
    markers.forEach(function (m) {
      var item = el("div", "co-cut-item");

      var labelRow = el("div", "co-cut-row");
      var dot = el("span", "co-marker-dot");
      dot.style.background = color;
      labelRow.appendChild(dot);
      labelRow.appendChild(el("span", "co-marker-label",
        m.label || m.eventType || source));
      var copyBtn = el("button", "co-cut-delete");
      copyBtn.type = "button";
      copyBtn.setAttribute("data-tooltip", "Copy to cuts");
      copyBtn.setAttribute("aria-label", "Copy to cuts");
      copyBtn.appendChild(el("span", "co-btn-icon co-icon-scissors"));
      copyBtn.addEventListener("click", function (e) {
        e.stopPropagation();
        copyMarkerToCut(m);
      });
      labelRow.appendChild(copyBtn);
      item.appendChild(labelRow);

      var timeRow = el("div", "co-cut-row");
      timeRow.appendChild(el("span", "co-cut-times",
        formatTime(m.start, { decimals: 1 }) +
        (m.end > m.start ? " – " + formatTime(m.end, { decimals: 1 }) : "")));
      if (m.end > m.start) {
        timeRow.appendChild(el("span", "co-cut-dur", formatDuration(m.end - m.start)));
      }
      if (m.trimmed) {
        var trimBadge = el("span", "co-trim-badge", "trimmed");
        trimBadge.setAttribute("data-tooltip",
          "Original: " + formatTime(m.origStart, { decimals: 1 }) +
          " – " + formatTime(m.origEnd, { decimals: 1 }));
        timeRow.appendChild(trimBadge);
        var resetBtn = el("button", "co-cut-delete");
        resetBtn.type = "button";
        resetBtn.setAttribute("data-tooltip", "Reset trim to the source span");
        resetBtn.setAttribute("aria-label", "Reset trim");
        resetBtn.appendChild(el("span", "co-btn-icon co-icon-undo"));
        resetBtn.addEventListener("click", function (e) {
          e.stopPropagation();
          resetTrim(m);
        });
        timeRow.appendChild(resetBtn);
      }
      item.appendChild(timeRow);

      // Set-or-remove: an empty data-tooltip would still make the row a hover
      // anchor, swallowing the child buttons' tooltips on the way past.
      if (m.label) item.setAttribute("data-tooltip", m.label);
      else item.removeAttribute("data-tooltip");
      item.addEventListener("click", function () {
        seekVideo(m.start);
      });
      frag.appendChild(item);
    });
    list.appendChild(frag);
  }

  function renderSidebar() {
    if (state.sidebarTab === "cuts") renderCutList();
    else renderMarkerList(state.sidebarTab);
  }
  CO.renderSidebar = renderSidebar;

  function initSidebarTabs() {
    var tabs = qsa(".co-panel-tab");
    tabs.forEach(function (tab) {
      tab.addEventListener("click", function () {
        state.sidebarTab = tab.getAttribute("data-tab");
        tabs.forEach(function (t) {
          var active = t === tab;
          t.classList.toggle("active", active);
          t.setAttribute("aria-selected", active ? "true" : "false");
        });
        renderSidebar();
      });
    });
  }

  // ---- Right timelines panel: collapse/expand (browser-local, like Studio) ----

  var SIDEBAR_OPEN_KEY = "clipgen-composer-sidebar-open";

  function applySidebarOpen() {
    var panel = qs("#coCutPanel");
    if (panel) panel.setAttribute("data-open", state.sidebarOpen ? "true" : "false");
    var btn = qs("#coCutPanelToggle");
    if (btn) btn.setAttribute("aria-expanded", state.sidebarOpen ? "true" : "false");
  }

  function readPersistedSidebarOpen() {
    try {
      var stored = localStorage.getItem(SIDEBAR_OPEN_KEY);
      if (stored !== null) state.sidebarOpen = stored !== "false";
    } catch (_) {}
    // Apply before first paint so a persisted-collapsed panel shows collapsed
    // immediately; the flex-basis transition is gated on `.tx-ready` (added
    // after first paint below), so the initial open→collapsed flip never
    // animates. Mirrors studio.js readPersistedSidebarOpen.
    applySidebarOpen();
    var panel = qs("#coCutPanel");
    if (!panel) return;
    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        panel.classList.add("tx-ready");
      });
    });
  }

  function toggleSidebar() {
    state.sidebarOpen = !state.sidebarOpen;
    try {
      localStorage.setItem(SIDEBAR_OPEN_KEY, state.sidebarOpen ? "true" : "false");
    } catch (_) {}
    applySidebarOpen();
  }

  // Run synchronously at script load (composer.js is deferred, so #coCutPanel
  // already exists) rather than in boot(): sets data-open before first paint.
  readPersistedSidebarOpen();

  // ---- Generate (Studio intake endpoint; NDJSON streaming) ----
  // apiPostNDJSON is an ambient utils.js global.

  var _generateAbort = null;

  function updateGenerateButton() {
    var btn = qs("#coGenerateBtn");
    btn.disabled = state.generating || participantCuts().length === 0;
  }
  CO.updateGenerateButton = updateGenerateButton;

  function onGenerate() {
    var cuts = participantCuts();
    if (!cuts.length || state.generating) return;
    state.generating = true;
    cuts.forEach(function (c) { delete c._genStatus; });
    var btn = qs("#coGenerateBtn");
    var cancelBtn = qs("#coCancelBtn");
    btn.textContent = "Generating… 0/" + cuts.length;
    cancelBtn.classList.remove("hidden");
    updateGenerateButton();

    var items = cuts.map(function (c) {
      return {
        participant: c.participant,
        start: c.start,
        end: c.end,
        event_type: c.label || "composer",
        event_ids: [c.id],
        source: "composer",
      };
    });
    var done = 0;
    var failed = 0;

    function handleLine(line) {
      var data;
      try { data = JSON.parse(line); } catch (_) { return; }
      if (!data || typeof data.index !== "number") return;
      var cut = cuts[data.index];
      if (cut) cut._genStatus = data.ok ? "ok" : "fail";
      if (!data.ok) failed++;
      done++;
      logArtifactResult(data, cut);
      btn.textContent = "Generating… " + done + "/" + cuts.length;
      renderCutList();
    }

    function finish(message) {
      state.generating = false;
      _generateAbort = null;
      btn.textContent = "Generate clips";
      cancelBtn.classList.add("hidden");
      updateGenerateButton();
      renderCutList();
      if (message) showToast(message);
      else if (failed) showToast((done - failed) + " clip(s) generated, " + failed + " failed");
      else showToast(done + " clip(s) generated");
    }

    _generateAbort = new AbortController();
    apiPostNDJSON(
      "../studio/api/generate-intake",
      { items: items, format: "clip" },
      { signal: _generateAbort.signal, onLine: handleLine }
    )
      .then(function () { finish(); })
      .catch(function (err) {
        var aborted = err && (err.name === "AbortError" || err.code === 20);
        finish(aborted ? "Generation cancelled" : "Generation failed: " + (err && err.message));
      });
  }

  function onCancelGenerate() {
    apiPost("../studio/api/generate-intake/cancel", {}).catch(function () {});
    if (_generateAbort) _generateAbort.abort();
  }

  // ---- Artifact log (TopNav #logBtn) ----
  //
  // Shares Studio's markup, styling AND open/close path: utils.js's
  // popModalIn/popModalOut drive the card pop and ramp the shared .cg-modal-veil
  // backdrop. A separate copy with its own veil constants and exit timer is what
  // made spam-toggling pop.

  function logArtifactResult(data, cut) {
    var artifact = data.artifact || {};
    state.artifactLog.push({
      ok: !!data.ok,
      type: artifact.type || "clip",
      file: artifact.file ? String(artifact.file).split(/[\\/]/).pop() : "",
      participant: artifact.participant || (cut && cut.participant) || "",
      start: artifact.start !== undefined ? artifact.start : (cut && cut.start),
      end: artifact.end !== undefined ? artifact.end : (cut && cut.end),
      error: data.error || "",
      at: new Date(),
    });
    if (!qs("#logOverlay").classList.contains("hidden")) renderLog();
  }

  function renderLog() {
    var content = qs("#logContent");
    var countEl = qs("#logCount");
    content.innerHTML = "";
    if (!state.artifactLog.length) {
      content.appendChild(el("div", "log-empty", "No clips generated yet."));
      countEl.textContent = "";
      return;
    }
    var frag = document.createDocumentFragment();
    for (var i = state.artifactLog.length - 1; i >= 0; i--) {
      var entry = state.artifactLog[i];
      var row = el("div", "log-entry");
      var typeLabel = entry.ok ? (entry.type || "clip") : "fail";
      var badge = el("span", "log-type-badge", typeLabel);
      badge.setAttribute("data-type", typeLabel);
      row.appendChild(badge);
      row.appendChild(el("span", "log-entry-file",
        entry.ok ? (entry.file || "clip") : (entry.error || "unknown error")));
      // A screenshot is a single instant (start === end): show one timestamp, not
      // a zero-length "t – t" range; drop the time entirely if we have no numbers.
      var meta = entry.participant;
      var hasStart = typeof entry.start === "number" && isFinite(entry.start);
      var hasEnd = typeof entry.end === "number" && isFinite(entry.end);
      if (hasStart && hasEnd && entry.start !== entry.end) {
        meta += " · " + formatTime(entry.start, { decimals: 1 }) +
          " – " + formatTime(entry.end, { decimals: 1 });
      } else if (hasStart) {
        meta += " · " + formatTime(entry.start, { decimals: 1 });
      }
      meta += " · " + entry.at.toLocaleTimeString();
      row.appendChild(el("span", "log-entry-meta", meta));
      frag.appendChild(row);
    }
    content.appendChild(frag);
    var okCount = state.artifactLog.filter(function (e) { return e.ok; }).length;
    countEl.textContent = okCount + " clip(s) generated this session";
  }

  function openLog() {
    var overlay = qs("#logOverlay");
    popModalIn(overlay, qs(".log-panel"));
    // Gates the topnav's own backdrop-filter so this modal's veil composites
    // over the bar instead of reading through it (see topnav.css).
    document.body.classList.add("modal-open");
    // Escape, Tab-trapping and focus restore, same as Studio's log. The trap
    // owns Escape while it is active, so the page's back-out cascade does not
    // need a branch for this overlay.
    openBlockingModal(overlay, {
      onEscape: closeLog,
      trapFocus: true,
      restoreFocus: true,
    });
    renderLog();
  }

  function closeLog() {
    var overlay = qs("#logOverlay");
    // Trap and topnav gate are released with the visual hide, not before it:
    // focus jumping back to the trigger while the veil is still up reads as the
    // panel already being gone. popModalOut's generation guard makes deferring
    // them safe against a re-open mid-exit.
    popModalOut(overlay, qs(".log-panel"), function () {
      closeBlockingModal(overlay);
      overlay.classList.add("hidden");
      document.body.classList.remove("modal-open");
    });
  }

  // ---- Keyboard (shared hotkeys.js registry) ----

  // j/k list-nav: select the next/previous cut (by start time) and move the
  // playhead to its in point.
  function selectAdjacentCut(delta) {
    var cuts = sortedCuts();
    if (!cuts.length) return;
    var idx = -1;
    for (var i = 0; i < cuts.length; i++) {
      if (cuts[i].id === state.selectedCutId) { idx = i; break; }
    }
    var next = idx < 0 ? (delta > 0 ? 0 : cuts.length - 1) : idx + delta;
    if (next < 0) next = 0;
    if (next > cuts.length - 1) next = cuts.length - 1;
    selectCut(cuts[next].id);
    seekVideo(cuts[next].start);
  }

  // One video frame in seconds, from the participant's probed frame rate.
  function frameStep() {
    var p = findParticipant(state.participant);
    return p && p.fps ? 1 / p.fps : 1 / 30;
  }

  // B: tap toggles annotation hiding, hold peeks (inverts while held).
  var _annHold = null; // {prev, at} while B is down

  function initKeyboard() {
    window.ClipgenHotkeys.register([
      { id: "transport.playPause", handler: function () { togglePlay(); } },
      { id: "transport.seekBack", handler: function () { seekVideo(state.playhead - 5); } },
      { id: "transport.seekFwd", handler: function () { seekVideo(state.playhead + 5); } },
      { id: "transport.stepBack", handler: function () { seekVideo(state.playhead - 1); } },
      { id: "transport.stepFwd", handler: function () { seekVideo(state.playhead + 1); } },
      { id: "composer.seekBackMid", handler: function () { seekVideo(state.playhead - 2.5); } },
      { id: "composer.seekFwdMid", handler: function () { seekVideo(state.playhead + 2.5); } },
      { id: "composer.stepBackFrame", handler: function () { seekVideo(state.playhead - frameStep()); } },
      { id: "composer.stepFwdFrame", handler: function () { seekVideo(state.playhead + frameStep()); } },
      {
        id: "composer.holdHideAnnotations",
        repeat: false,
        handler: function () {
          if (_annHold) return; // blur can swallow a keyup; don't restack
          _annHold = { prev: state.annHidden, at: Date.now() };
          setAnnotationsHidden(!_annHold.prev);
        },
        onRelease: function () {
          if (!_annHold) return;
          var hold = _annHold;
          _annHold = null;
          // A quick tap keeps the flip (toggle); a hold reverts it (peek).
          if (Date.now() - hold.at >= 250) setAnnotationsHidden(hold.prev);
        },
      },
      { id: "nav.next", handler: function () { selectAdjacentCut(1); } },
      { id: "nav.prev", handler: function () { selectAdjacentCut(-1); } },
      { id: "composer.setIn", handler: function () { setInPoint(); } },
      { id: "composer.setOut", handler: function () { setOutPoint(); } },
      { id: "composer.nudgeLeft", handler: function () { nudgeSelectedCut(-0.1); } },
      { id: "composer.nudgeRight", handler: function () { nudgeSelectedCut(0.1); } },
      { id: "composer.nudgeLeftBig", handler: function () { nudgeSelectedCut(-1); } },
      { id: "composer.nudgeRightBig", handler: function () { nudgeSelectedCut(1); } },
      {
        id: "composer.deleteSelection",
        when: function () { return !!(state.selectedAnnotationIds.length || state.selectedCutId); },
        handler: function () {
          if (state.selectedAnnotationIds.length) deleteSelectedAnnotations();
          else deleteCut(state.selectedCutId);
        },
      },
      { id: "composer.toolSelect", handler: function () { setAnnotateTool("select"); } },
      { id: "composer.toolText", handler: function () { setAnnotateTool("text"); } },
      { id: "composer.toolDraw", handler: function () { setAnnotateTool("draw"); } },
      { id: "composer.toolErase", handler: function () { setAnnotateTool("erase"); } },
      { id: "composer.toolRect", handler: function () { setAnnotateTool("rect"); } },
      { id: "composer.toolEllipse", handler: function () { setAnnotateTool("ellipse"); } },
      {
        id: "composer.toggleSource",
        handler: function (e, combo) {
          var src = ["sheet", "screenspace", "transcript"][parseInt(combo, 10) - 1];
          if (!src || !CO.toggleSource) return false;
          CO.toggleSource(src);
        },
      },
      {
        id: "composer.toggleAllSources",
        handler: function () { CO.toggleAllSources && CO.toggleAllSources(); },
      },
      {
        id: "composer.toggleThumbs",
        handler: function () { qs("#coThumbsToggle").click(); },
      },
      {
        id: "composer.toggleScrubAudio",
        handler: function () { qs("#coScrubAudioToggle").click(); },
      },
      { id: "composer.toggleSidebar", handler: function () { toggleSidebar(); } },
      { id: "global.primary", handler: function () { onGenerate(); } },
      { id: "edit.undo", handler: function () { undo(); } },
      { id: "edit.redo", handler: function () { redo(); } },
      { id: "composer.note.zoomTimeline" },
    ]);

    // Back-out cascade, one level per press (order matters: overlay first,
    // then pending in-point, then tool, then selections).
    // The artifact log is absent here on purpose: it now runs a blocking-modal
    // trap that owns Escape while open, so hotkeys.js never reaches this cascade.
    window.ClipgenHotkeys.registerEscape(function () {
      if (state.pendingIn !== null) {
        state.pendingIn = null;
        updatePendingInfo();
        renderPlayhead();
        return true;
      }
      if (state.annTool !== "select") { setAnnotateTool("select"); return true; }
      if (state.selectedAnnotationIds.length) { selectAnnotation(null); return true; }
      if (state.selectedCutId) { selectCut(null); return true; }
      return false;
    });
  }

  // Command palette (command-palette.js): Composer registers no TopNav quick
  // actions, so besides the built-in nav/global entries this is the page's
  // whole palette — toolbar actions, lane toggles, participant jumps.
  function initCommandPalette() {
    if (!window.ClipgenCommandPalette) return;
    function buttonCommand(id, title, icon, keywords, elId) {
      return {
        id: id,
        title: title,
        icon: icon,
        keywords: keywords,
        section: "Composer",
        enabled: function () {
          var btn = qs("#" + elId);
          return !!btn && !btn.disabled;
        },
        run: function () { qs("#" + elId).click(); },
      };
    }
    // Left timeline-list tabs (Cuts / Sheet / Screen / Script) carry data-tab;
    // click the matching tab so initSidebarTabs sets state.sidebarTab.
    function listTabCommand(dataTab, title, icon) {
      return {
        id: "composer:list-" + dataTab,
        title: title,
        icon: icon,
        keywords: "list panel sidebar timeline show " + dataTab,
        section: "Composer",
        visible: function () { return !!qs('.co-panel-tab[data-tab="' + dataTab + '"]'); },
        run: function () {
          var t = qs('.co-panel-tab[data-tab="' + dataTab + '"]');
          if (t) t.click();
        },
      };
    }
    window.ClipgenCommandPalette.setParticipants(function () {
      return (state.participants || []).map(function (p) { return p.id; });
    });
    window.ClipgenCommandPalette.register("composer", function () {
      var cmds = [
        buttonCommand("composer:generate", "Generate clips", "play",
          "cuts render build", "coGenerateBtn"),
        buttonCommand("composer:export-shot", "Export screenshot", "arrow-down-tray",
          "annotated frame image", "coExportShotBtn"),
        buttonCommand("composer:export-gif", "Export GIF", "arrow-down-tray",
          "annotated animation", "coExportGifBtn"),
        buttonCommand("composer:export-burn", "Export burned clip", "arrow-down-tray",
          "annotations video render", "coExportBurnBtn"),
        buttonCommand("composer:undo", "Undo", "arrow-uturn-left",
          "revert history", "coUndoBtn"),
        buttonCommand("composer:redo", "Redo", "arrow-uturn-right",
          "repeat history", "coRedoBtn"),
        buttonCommand("composer:lane-sheet", "Toggle Sheet marker lane", "queue-list",
          "timestamps markers", "coLaneSheet"),
        buttonCommand("composer:lane-screenspace", "Toggle Screenspace marker lane", "queue-list",
          "events markers", "coLaneScreenspace"),
        buttonCommand("composer:lane-transcript", "Toggle Transcript marker lane", "queue-list",
          "marks markers", "coLaneTranscript"),
        buttonCommand("composer:thumbs", "Toggle marker thumbnails", "photo",
          "filmstrip sprites frames strips", "coThumbsToggle"),
        buttonCommand("composer:scrub-audio", "Toggle marker audio scrub", "speaker-wave",
          "waveform sound hover", "coScrubAudioToggle"),
        buttonCommand("composer:toggle-sidebar", "Toggle timelines sidebar", "adjustments-horizontal",
          "show hide panel drawer collapse", "coCutPanelToggle"),
        buttonCommand("composer:shortcuts", "Keyboard shortcuts", "command-line",
          "cheatsheet keys help", "coShortcutsBtn"),
        listTabCommand("cuts", "Show Cuts list", "list-bullet"),
        listTabCommand("sheet", "Show Sheet list", "table-cells"),
        listTabCommand("screenspace", "Show Screenspace list", "queue-list"),
        listTabCommand("transcript", "Show Transcript list", "queue-list"),
        {
          id: "composer:log",
          title: "Toggle artifact log",
          icon: "list-bullet",
          keywords: "history builds panel drawer",
          section: "Composer",
          visible: function () { return !!document.getElementById("logBtn"); },
          run: function () { document.getElementById("logBtn").click(); },
        },
      ];
      // "Jump to … in Composer" = stays here and selects in place; the
      // palette's built-in provider adds the cross-page "Open … in <Page>".
      (state.participants || []).forEach(function (p) {
        cmds.push({
          id: "composer:p:" + p.id,
          title: "Jump to " + p.id + " in Composer",
          icon: "user",
          keywords: "participant select source",
          section: "Participants",
          run: function () {
            qs("#coParticipantSelect").value = p.id;
            selectParticipant(p.id);
          },
        });
      });
      return cmds;
    });
  }

  // ---- Boot ----

  function boot() {
    // TopNav renders the theme toggle (#themeToggle) and Settings (#settingsBtn)
    // buttons synchronously before this hub loads, so wire them here as the
    // other surfaces do. Theme flips repaint the canvases (their colors are
    // sampled from CSS variables at draw time).
    if (typeof initThemeToggle === "function") {
      initThemeToggle(function () {
        if (CO.invalidateLaneColors) CO.invalidateLaneColors();
        renderTimeline();
      });
    }
    var settingsBtn = qs("#settingsBtn");
    if (settingsBtn && typeof window.openSettingsModal === "function") {
      settingsBtn.addEventListener("click", function () {
        // Saved/reset settings apply live on this page: re-sync the mirrored
        // config flag and the footer hint that advertises it.
        function syncComposerSettings(settings) {
          (settings || []).forEach(function (s) {
            if (s.name === "COMPOSER_DOUBLE_CLICK_CUTS") {
              CLIPGEN_CONFIG.composerDoubleClickCuts = !!s.value;
            }
          });
          updateTimelineHint();
        }
        window.openSettingsModal({
          initialTab: "Composer",
          onSave: function (_applied, settings) { syncComposerSettings(settings); },
          onReset: function (_scope, settings) { syncComposerSettings(settings); },
        });
      });
    }

    state.annColor = CLIPGEN_CONFIG.composerAnnotationColor;
    state.annStrokeWidth = CLIPGEN_CONFIG.composerAnnotationStrokeWidth;
    state.annStrokeStyle = CLIPGEN_CONFIG.composerAnnotationStrokeStyle;
    updateTimelineHint();
    initCommandPalette();
    initParticipantSelect();
    initVideo();
    initKeyboard();
    initTimeline();
    initMarkerToggles();
    initMarkerScrub();
    initAnnotate();
    qs("#coToolHide").addEventListener("click", function () {
      setAnnotationsHidden(!state.annHidden);
    });
    qs("#coExportShotBtn").addEventListener("click", exportScreenshot);
    qs("#coExportGifBtn").addEventListener("click", function () { exportBurn(true); });
    qs("#coExportBurnBtn").addEventListener("click", function () { exportBurn(false); });

    qs("#coSetInBtn").addEventListener("click", setInPoint);
    qs("#coSetOutBtn").addEventListener("click", setOutPoint);
    qs("#coGenerateBtn").addEventListener("click", onGenerate);
    qs("#coCancelBtn").addEventListener("click", onCancelGenerate);
    qs("#coUndoBtn").addEventListener("click", undo);
    qs("#coRedoBtn").addEventListener("click", redo);
    // Open-only, matching Studio: dismiss is the X, the backdrop, or Escape. A
    // toggle here made the button a spam surface that fought its own animation.
    var logBtn = qs("#logBtn");
    if (logBtn) logBtn.addEventListener("click", openLog);
    qs("#logClose").addEventListener("click", closeLog);
    qs("#logOverlay").addEventListener("click", function (e) {
      if (e.target === qs("#logOverlay")) closeLog();
    });
    qs("#coCutPanelToggle").addEventListener("click", toggleSidebar);
    initSidebarTabs();

    // The two boot fetches run in parallel, but participant auto-select MUST
    // wait for the manifest: selectParticipant → loadMarkers → commitLane
    // overlays trims from state.trims, which is empty until the manifest
    // lands — racing them showed saved trims at their source spans.
    var manifestLoaded = apiGet("api/manifest").then(function (data) {
      if (!data.ok || !data.manifest) return;
      state.cuts = data.manifest.cuts || [];
      state.trims = data.manifest.trims || {};
      state.annotations = data.manifest.annotations || [];
      var ui = data.manifest.ui || {};
      if (ui.markerSources) {
        Object.keys(state.sourceToggles).forEach(function (src) {
          if (ui.markerSources[src] !== undefined) {
            state.sourceToggles[src] = !!ui.markerSources[src];
          }
        });
        CO.syncSourcePills && CO.syncSourcePills();
      }
      if (ui.laneFolds) {
        Object.keys(state.laneFolds).forEach(function (src) {
          if (ui.laneFolds[src] !== undefined) {
            state.laneFolds[src] = !!ui.laneFolds[src];
          }
        });
      }
      if (typeof ui.markerThumbnails === "boolean") {
        state.markerThumbnails = ui.markerThumbnails;
      }
      if (typeof ui.markerAudioScrub === "boolean") {
        state.markerAudioScrub = ui.markerAudioScrub;
      }
      if (typeof ui.followPlayhead === "boolean") {
        state.followPlayhead = ui.followPlayhead;
      }
      if (CO.syncScrubToggles) CO.syncScrubToggles();
      updateGenerateButton();
      renderCutList();
      if (CO.updateTimelineHeight) CO.updateTimelineHeight();
      renderTimeline();
      renderAnnotations();
    }).catch(function () {});

    apiGet("api/participants").then(function (data) {
      if (!data.ok) return;
      if (data.config) clipgenApplyConfig(data.config);
      updateTimelineHint(); // the double-click hint follows the fetched config
      state.hasSheet = !!data.has_sheet;
      state.participants = data.participants || [];
      populateParticipantSelect();
      return manifestLoaded.then(function () {
        // A /composer/#P07 hash (command palette / cross-page links) wins;
        // otherwise restore the last-worked-on participant, falling back to
        // auto-select when there is only one.
        var hashPid = clipgenHashParticipant();
        var stored = getStoredUIState("composer").participant;
        var initial = hashPid && findParticipant(hashPid)
          ? hashPid
          : stored && findParticipant(stored)
            ? stored
            : state.participants.length === 1 ? state.participants[0].id : null;
        if (initial) {
          qs("#coParticipantSelect").value = initial;
          selectParticipant(initial);
        }
      });
    }).catch(function () { showToast("Could not load participants"); });
  }

  document.addEventListener("DOMContentLoaded", boot);
})();
