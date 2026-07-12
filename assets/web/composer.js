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

  var state = {
    participants: [],       // [{id, parts:[{name,duration,offset}], total_duration}]
    participant: null,      // active participant id
    parts: [],              // active participant's parts
    duration: 0,            // stitched total (seconds)
    activePart: 0,
    playhead: 0,            // global seconds (kept current by timeupdate)
    playing: false,
    cuts: [],               // all cuts from the composer manifest (all participants)
    selectedCutId: null,
    pendingIn: null,        // in-point awaiting its out-point (global seconds)
    markers: { sheet: [], screenspace: [], transcript: [] },
    sourceToggles: { sheet: true, screenspace: true, transcript: true },
    zoom: 1,
    offset: 0,              // timeline pan offset (global seconds)
    dragging: false,        // timeline satellite sets during pan/edge drags
    generating: false,
  };

  var CO = { state: state };
  window.ClipgenComposer = CO;

  // ---- Satellite delegators (guarded; satellites load after the hub) ----

  function renderTimeline() { return CO.renderTimeline && CO.renderTimeline.apply(null, arguments); }
  function renderPlayhead() { return CO.renderPlayhead && CO.renderPlayhead.apply(null, arguments); }
  function initTimeline() { return CO.initTimeline && CO.initTimeline.apply(null, arguments); }
  function initMarkerToggles() { return CO.initMarkerToggles && CO.initMarkerToggles.apply(null, arguments); }
  function loadMarkers() { return CO.loadMarkers && CO.loadMarkers.apply(null, arguments); }

  // ---- API client ----

  function apiGet(path) {
    return fetch(path).then(function (r) { return r.json(); });
  }

  function apiSend(method, path, body) {
    return fetch(path, {
      method: method,
      headers: { "Content-Type": "application/json" },
      body: body === undefined ? undefined : JSON.stringify(body),
    }).then(function (r) { return r.json(); });
  }

  CO.apiGet = apiGet;
  CO.apiSend = apiSend;

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

  var _pendingSeekTime = null;
  var _seekRaf = 0;
  var _pendingSeekListener = null;

  function cancelPendingSeek() {
    var video = qs("#coVideo");
    _pendingSeekTime = null;
    cancelAnimationFrame(_seekRaf);
    _seekRaf = 0;
    if (_pendingSeekListener) {
      if (video) video.removeEventListener("loadedmetadata", _pendingSeekListener);
      _pendingSeekListener = null;
    }
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
    renderPlayhead();
    updateTimeLabel();
  }
  CO.seekVideo = seekVideo;

  function seekLocal(time) {
    var video = qs("#coVideo");
    if (!video || !video.src) return;

    if (_pendingSeekListener) {
      video.removeEventListener("loadedmetadata", _pendingSeekListener);
      _pendingSeekListener = null;
    }
    // Metadata not loaded yet: defer the seek until it is.
    if (video.readyState < 1) {
      _pendingSeekTime = time;
      _pendingSeekListener = function () {
        video.removeEventListener("loadedmetadata", _pendingSeekListener);
        _pendingSeekListener = null;
        var t = _pendingSeekTime;
        _pendingSeekTime = null;
        if (t !== null) seekLocal(t);
      };
      video.addEventListener("loadedmetadata", _pendingSeekListener);
      return;
    }
    // Coalesce rapid seeks into one per animation frame.
    _pendingSeekTime = time;
    cancelAnimationFrame(_seekRaf);
    _seekRaf = requestAnimationFrame(function () {
      var t = _pendingSeekTime;
      _pendingSeekTime = null;
      _seekRaf = 0;
      if (t === null) return;
      video.currentTime = t;
    });
  }

  function switchToPart(i, localTime, resume) {
    var video = qs("#coVideo");
    if (!video) return;
    state.activePart = i;
    video.src = "media/" + encodeURIComponent(state.parts[i].name);
    video.load();
    var onMeta = function () {
      video.removeEventListener("loadedmetadata", onMeta);
      video.currentTime = localTime;
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

  function updateTimeLabel() {
    var label = qs("#coTimeLabel");
    if (!label) return;
    label.textContent = formatTime(state.playhead, { decimals: 0 }) +
      " / " + formatTime(state.duration, { decimals: 0 });
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
        renderPlayhead();
      });
    });
    // Single-part participants may have no probed duration (server couldn't
    // ffprobe) — fall back to the element's own metadata.
    video.addEventListener("loadedmetadata", function () {
      if (!state.duration && state.parts.length === 1 && isFinite(video.duration)) {
        state.duration = video.duration;
        state.parts[0].duration = video.duration;
        updateTimeLabel();
        renderTimeline();
      }
    });
    qs("#coPlayBtn").addEventListener("click", togglePlay);
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
    cancelPendingSeek();
    state.participant = pid;
    state.parts = p.parts || [];
    state.duration = p.total_duration || 0;
    state.activePart = 0;
    state.playhead = 0;
    state.playing = false;
    state.pendingIn = null;
    state.selectedCutId = null;
    state.zoom = 1;
    state.offset = 0;
    state.markers = { sheet: [], screenspace: [], transcript: [] };

    var video = qs("#coVideo");
    video.pause();
    if (state.parts.length) {
      video.src = "media/" + encodeURIComponent(state.parts[0].name);
      video.load();
      qs("#coVideoFrame").classList.add("has-video");
    } else {
      video.removeAttribute("src");
      qs("#coVideoFrame").classList.remove("has-video");
    }

    ["#coPlayBtn", "#coSetInBtn", "#coSetOutBtn"].forEach(function (sel) {
      qs(sel).disabled = false;
    });
    updatePendingInfo();
    updatePlayButton();
    updateTimeLabel();
    updateGenerateButton();
    renderCutList();
    renderTimeline();
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
      var opt = el("option", "", p.id);
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
      : "In: " + formatTime(state.pendingIn, { decimals: 1 }) + " — press O";
  }

  function setInPoint() {
    if (!state.participant) return;
    state.pendingIn = state.playhead;
    updatePendingInfo();
    renderPlayhead();
  }

  function setOutPoint() {
    if (!state.participant || state.pendingIn === null) {
      showToast("Set an in point first (I)");
      return;
    }
    var start = state.pendingIn;
    var end = state.playhead;
    if (end < start) { var tmp = start; start = end; end = tmp; }
    if (end - start < 0.2) {
      showToast("Cut too short — move the playhead past the in point");
      return;
    }
    apiSend("POST", "api/cuts", {
      participant: state.participant,
      start: start,
      end: end,
    }).then(function (data) {
      if (!data.ok) { showToast(data.error || "Could not save cut"); return; }
      state.cuts.push(data.cut);
      state.pendingIn = null;
      state.selectedCutId = data.cut.id;
      updatePendingInfo();
      updateGenerateButton();
      renderCutList();
      renderTimeline();
    }).catch(function () { showToast("Could not save cut"); });
  }

  function deleteCut(id) {
    apiSend("DELETE", "api/cuts/" + encodeURIComponent(id)).then(function (data) {
      if (!data.ok) { showToast(data.error || "Could not delete cut"); return; }
      state.cuts = state.cuts.filter(function (c) { return c.id !== id; });
      if (state.selectedCutId === id) state.selectedCutId = null;
      updateGenerateButton();
      renderCutList();
      renderTimeline();
    }).catch(function () { showToast("Could not delete cut"); });
  }
  CO.deleteCut = deleteCut;

  // Persist a cut's edited span (timeline edge/body drags land here on drag end).
  function commitCutTimes(cut) {
    apiSend("PATCH", "api/cuts/" + encodeURIComponent(cut.id), {
      start: cut.start,
      end: cut.end,
    }).then(function (data) {
      if (!data.ok) { showToast(data.error || "Could not save cut"); return; }
      cut.start = data.cut.start;
      cut.end = data.cut.end;
      renderCutList();
      renderTimeline();
    }).catch(function () { showToast("Could not save cut"); });
  }
  CO.commitCutTimes = commitCutTimes;

  function selectCut(id) {
    state.selectedCutId = id;
    renderCutList();
    renderTimeline();
  }
  CO.selectCut = selectCut;

  function nudgeSelectedCut(deltaSeconds) {
    var cut = state.selectedCutId && findCut(state.selectedCutId);
    if (!cut) return;
    // Nudge the out edge; hold the in edge fixed (the common trim gesture).
    cut.end = clamp(cut.end + deltaSeconds, cut.start + 0.2, state.duration || cut.end + deltaSeconds);
    renderTimeline();
    commitCutTimes(cut);
  }

  function renderCutList() {
    var list = qs("#coCutList");
    list.innerHTML = "";
    var cuts = participantCuts();
    if (!cuts.length) {
      var hint = el("p", "co-empty-hint");
      hint.innerHTML = "Press <kbd>I</kbd> then <kbd>O</kbd> to add a cut.";
      list.appendChild(hint);
      return;
    }
    var frag = document.createDocumentFragment();
    cuts.forEach(function (cut) {
      var item = el("div", "co-cut-item" + (cut.id === state.selectedCutId ? " selected" : ""));
      item.appendChild(el("span", "co-cut-times",
        formatTime(cut.start, { decimals: 1 }) + " – " + formatTime(cut.end, { decimals: 1 })));
      item.appendChild(el("span", "co-cut-dur", formatDuration(cut.end - cut.start)));
      if (cut._genStatus) {
        item.appendChild(el("span", "co-cut-status " + cut._genStatus,
          cut._genStatus === "ok" ? "✓" : "✗"));
      }
      var del = el("button", "co-cut-delete");
      del.type = "button";
      del.title = "Delete cut";
      del.setAttribute("aria-label", "Delete cut");
      del.appendChild(el("span", "co-btn-icon co-icon-trash"));
      del.addEventListener("click", function (e) {
        e.stopPropagation();
        deleteCut(cut.id);
      });
      item.appendChild(del);
      item.addEventListener("click", function () {
        selectCut(cut.id);
        seekVideo(cut.start);
      });
      frag.appendChild(item);
    });
    list.appendChild(frag);
  }
  CO.renderCutList = renderCutList;

  // ---- Generate (Studio intake endpoint; NDJSON streaming) ----

  function readNDJSONStream(response, onLine) {
    if (!response.body || typeof response.body.getReader !== "function") {
      return Promise.reject(new Error("Streaming response not supported"));
    }
    var reader = response.body.getReader();
    var decoder = new TextDecoder();
    var buffer = "";
    function pump() {
      return reader.read().then(function (result) {
        if (result.done) {
          if (buffer.trim()) onLine(buffer.trim());
          return;
        }
        buffer += decoder.decode(result.value, { stream: true });
        var lines = buffer.split("\n");
        buffer = lines.pop();
        for (var i = 0; i < lines.length; i++) {
          if (lines[i].trim()) onLine(lines[i].trim());
        }
        return pump();
      });
    }
    return pump();
  }

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
    fetch("../studio/api/generate-intake", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ items: items, format: "clip" }),
      signal: _generateAbort.signal,
    })
      .then(function (response) {
        if (!response.ok) throw new Error("Server error " + response.status);
        return readNDJSONStream(response, handleLine).then(function () { finish(); });
      })
      .catch(function (err) {
        var aborted = err && (err.name === "AbortError" || err.code === 20);
        finish(aborted ? "Generation cancelled" : "Generation failed: " + (err && err.message));
      });
  }

  function onCancelGenerate() {
    apiSend("POST", "../studio/api/generate-intake/cancel").catch(function () {});
    if (_generateAbort) _generateAbort.abort();
  }

  // ---- Keyboard (transcripts input-guard pattern) ----

  function shortcutsMenu() { return qs("#coShortcutsMenu"); }

  function toggleShortcuts(force) {
    var menu = shortcutsMenu();
    var show = force !== undefined ? force : menu.classList.contains("hidden");
    menu.classList.toggle("hidden", !show);
    qs("#coShortcutsBtn").setAttribute("aria-expanded", show ? "true" : "false");
  }

  function onKeyDown(e) {
    if (e.target.matches("input, textarea, select, [contenteditable=true]")) return;
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    var handled = true;
    switch (e.key) {
      case " ":
      case "k":
        togglePlay();
        break;
      case "j":
        seekVideo(state.playhead - 5);
        break;
      case "l":
        seekVideo(state.playhead + 5);
        break;
      case ",":
        seekVideo(state.playhead - 1);
        break;
      case ".":
        seekVideo(state.playhead + 1);
        break;
      case "i":
        setInPoint();
        break;
      case "o":
        setOutPoint();
        break;
      case "[":
        nudgeSelectedCut(e.shiftKey ? -1 : -0.1);
        break;
      case "]":
        nudgeSelectedCut(e.shiftKey ? 1 : 0.1);
        break;
      case "x":
      case "Backspace":
        if (state.selectedCutId) deleteCut(state.selectedCutId);
        break;
      case "1":
      case "2":
      case "3":
        CO.toggleSource && CO.toggleSource(["sheet", "screenspace", "transcript"][Number(e.key) - 1]);
        break;
      case "`":
        CO.toggleAllSources && CO.toggleAllSources();
        break;
      case "g":
        onGenerate();
        break;
      case "?":
        toggleShortcuts();
        break;
      case "Escape":
        if (!shortcutsMenu().classList.contains("hidden")) toggleShortcuts(false);
        else if (state.pendingIn !== null) {
          state.pendingIn = null;
          updatePendingInfo();
          renderPlayhead();
        } else if (state.selectedCutId) {
          selectCut(null);
        }
        break;
      default:
        handled = false;
    }
    if (handled) e.preventDefault();
  }

  // Shift+[ / ] arrive as "{" / "}" on most layouts; route them to the nudge too.
  function normalizeBracketKeys(e) {
    if (e.key === "{") nudgeSelectedCut(-1);
    else if (e.key === "}") nudgeSelectedCut(1);
    else return false;
    e.preventDefault();
    return true;
  }

  function initKeyboard() {
    document.addEventListener("keydown", function (e) {
      if (e.target.matches("input, textarea, select, [contenteditable=true]")) return;
      if (e.metaKey || e.ctrlKey || e.altKey) return;
      if (normalizeBracketKeys(e)) return;
      onKeyDown(e);
    });
  }

  // ---- Boot ----

  function boot() {
    initParticipantSelect();
    initVideo();
    initKeyboard();
    initTimeline();
    initMarkerToggles();

    qs("#coSetInBtn").addEventListener("click", setInPoint);
    qs("#coSetOutBtn").addEventListener("click", setOutPoint);
    qs("#coGenerateBtn").addEventListener("click", onGenerate);
    qs("#coCancelBtn").addEventListener("click", onCancelGenerate);
    qs("#coShortcutsBtn").addEventListener("click", function () { toggleShortcuts(); });
    document.addEventListener("click", function (e) {
      var menu = shortcutsMenu();
      if (menu.classList.contains("hidden")) return;
      if (!menu.contains(e.target) && e.target !== qs("#coShortcutsBtn") &&
          !qs("#coShortcutsBtn").contains(e.target)) {
        toggleShortcuts(false);
      }
    });

    apiGet("api/participants").then(function (data) {
      if (!data.ok) return;
      if (data.config) clipgenApplyConfig(data.config);
      state.participants = data.participants || [];
      populateParticipantSelect();
      if (state.participants.length === 1) {
        qs("#coParticipantSelect").value = state.participants[0].id;
        selectParticipant(state.participants[0].id);
      }
    }).catch(function () { showToast("Could not load participants"); });

    apiGet("api/manifest").then(function (data) {
      if (!data.ok || !data.manifest) return;
      state.cuts = data.manifest.cuts || [];
      var ui = data.manifest.ui || {};
      if (ui.markerSources) {
        Object.keys(state.sourceToggles).forEach(function (src) {
          if (ui.markerSources[src] !== undefined) {
            state.sourceToggles[src] = !!ui.markerSources[src];
          }
        });
        CO.syncSourcePills && CO.syncSourcePills();
      }
      updateGenerateButton();
      renderCutList();
      renderTimeline();
    }).catch(function () {});
  }

  document.addEventListener("DOMContentLoaded", boot);
})();
