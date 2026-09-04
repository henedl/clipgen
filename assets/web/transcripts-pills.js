/* clipgen Transcripts participant-pills satellite — transcripts-pills.js
 *
 * The participant pill row: per-pill transcribe control (start/cancel/re-run),
 * agent-state dots, and the body-mounted options popover (model/language
 * overrides + per-agent run/stop rows with dependency gating), plus the
 * transcribe POST flow. Loaded after transcripts.js; reads the hub's shared
 * state + helpers through window.ClipgenTranscripts (TS) and publishes
 * renderPills / initPillOutsideClick / initPillWheelScroll back (loadParticipants,
 * selectParticipant, and the poller render pills; boot wires the listeners).
 *
 * loadFriction lives in the agents satellite, which loads AFTER this one, so the
 * friction run/stop rows reach it late-bound via TS.loadFriction rather than a
 * captured local. _confirmUncachedWhisperModels stays in the hub (model-install
 * state). updateTranscribeFill comes from the video satellite, which loads
 * before this one, so it can be destructured like the hub's own exports. Plain
 * utils.js globals (qs/el/escapeHtml/apiPost/apiDelete/clipgenPluralUnit/
 * applyMaskIcon/attachHoverTooltip/positionPopoverAnchored) are reached via the
 * scope chain.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var showToast = TS.showToast,
    selectParticipant = TS.selectParticipant,
    maybeWarmOnPillHover = TS.maybeWarmOnPillHover,
    tryPostTranscriptionWarmup = TS.tryPostTranscriptionWarmup,
    pollTaskStatus = TS.pollTaskStatus,
    refreshTranscribeWording = TS.refreshTranscribeWording,
    updateTranscribeFill = TS.updateTranscribeFill, // video satellite (loads before this one)
    startPolling = TS.startPolling,
    _refreshAgentStateNow = TS._refreshAgentStateNow,
    _trFetchModels = TS._trFetchModels,
    ensureAgentModelInstalled = TS.ensureAgentModelInstalled,
    _confirmUncachedWhisperModels = TS._confirmUncachedWhisperModels,
    _isSpeakerTask = TS._isSpeakerTask,
    speakersEnabledFor = TS.speakersEnabledFor, // speakers satellite (loads before this one)
    setSpeakersEnabled = TS.setSpeakersEnabled,
    regenerateSpeakers = TS.regenerateSpeakers,
    stopSpeakers = TS.stopSpeakers;

  // Default icon left, hover icon right; the click always matches the status.
  var PILL_TRIGGER = {
    idle: { rest: "icons/microphone.svg", hover: "icons/microphone.svg", label: "Transcribe", action: "transcribe" },
    failed: { rest: "icons/exclamation-triangle.svg", hover: "icons/microphone.svg", label: "Retry transcription", action: "transcribe" },
    queued: { rest: "icons/clock.svg", hover: "icons/stop-circle.svg", label: "Cancel", action: "cancel" },
    running: { rest: "icons/arrow-path.svg", hover: "icons/stop-circle.svg", label: "Cancel transcription", action: "cancel" },
    // Nothing left to click: same glyph both faces, action "none" guards re-clicks.
    cancelling: { rest: "icons/stop-circle.svg", hover: "icons/stop-circle.svg", label: "Cancelling…", action: "none" },
    completed: { rest: "icons/check-circle.svg", hover: "icons/arrow-path.svg", label: "Re-transcribe", action: "retranscribe" }
  };

  var PILL_LANGUAGES = [
    { code: "", label: "Auto-detect" },
    { code: "en", label: "English" },
    { code: "sv", label: "Swedish" },
    { code: "es", label: "Spanish" },
    { code: "fr", label: "French" },
    { code: "de", label: "German" },
    { code: "it", label: "Italian" },
    { code: "pt", label: "Portuguese" },
    { code: "nl", label: "Dutch" },
    { code: "no", label: "Norwegian" },
    { code: "da", label: "Danish" },
    { code: "fi", label: "Finnish" }
  ];

  // UI-only state for pills (reset on reload)
  state.pillOverrides = {};    // pid → { model, language }
  state.pillOptionsOpen = null; // pid of the pill whose options pane is open

  function _dotStateTranscription(p, task) {
    if (task && task.status === "queued") return "queued";
    if (task && task.status === "running") return "running";
    if (task && task.status === "failed") return "failed";
    if (p.has_transcript) return "done";
    return "idle";
  }

  // Latest task per participant, bucketed: tx (transcription) and spk (speakers).
  function _indexTasks() {
    var idx = { tx: {}, spk: {} };
    state.tasks.forEach(function (t) {
      var bucket = _isSpeakerTask(t) ? idx.spk : idx.tx;
      if (!bucket[t.participant] || t.created_at > bucket[t.participant].created_at) {
        bucket[t.participant] = t;
      }
    });
    return idx;
  }

  function _agentsAttr(s) {
    return s.agents.transcription + "," + s.agents.summary + "," + s.agents.citations + "," +
      s.agents.friction + "," + s.agents.speakers;
  }

  function pillState(p, idx) {
    var task = idx.tx[p.id];
    var spkTask = idx.spk[p.id];
    var status = "idle";
    var progress = 0;
    var taskId = null;

    if (task && (task.status === "running" || task.status === "queued")) {
      taskId = task.id;
      var pending = state.cancellingTasks[task.id];
      if (pending) {
        status = "cancelling";
        // Frozen at the click so the fill cannot creep under "Cancelling…".
        progress = pending.progress;
      } else {
        status = task.status;
        if (task.status === "running") progress = Math.round((task.progress || 0) * 100);
      }
    } else if (task && task.status === "failed") {
      status = "failed";
      taskId = task.id;
    } else if (p.has_transcript) {
      status = "completed";
      progress = 100;
    }
    var spkStatus = "idle";
    var spkProgress = 0;
    if (spkTask && (spkTask.status === "running" || spkTask.status === "queued")) {
      spkStatus = spkTask.status;
      if (spkTask.status === "running") spkProgress = Math.round((spkTask.progress || 0) * 100);
    } else if (spkTask && spkTask.status === "failed") {
      spkStatus = "failed";
    } else if (p.speakers && p.speakers.count > 0) {
      spkStatus = "done";
    }
    var agents = {
      transcription: _dotStateTranscription(p, task),
      summary: (p.agents && p.agents.summary) || "idle",
      citations: (p.agents && p.agents.citations) || "idle",
      friction: (p.agents && p.agents.friction) || "idle",
      speakers: spkStatus,
    };
    return {
      status: status,
      progress: progress,
      taskId: taskId,
      agents: agents,
      // Running sub-state ("loading_model" / "transcribing") for the dot tooltip.
      phase: task ? task.phase : null,
      speakers: { status: spkStatus, taskId: spkTask ? spkTask.id : null, progress: spkProgress },
    };
  }

  // "1" only with a sheet loaded; a string so the patch guard compares it.
  function offSheetFlag(p) {
    return state.hasSheet && p.in_sheet === false ? "1" : "0";
  }

  function renderPills() {
    var container = qs("#participantPills");
    if (!container) return;
    // Real answer now, so the boot skeletons and aria-busy are done.
    container.removeAttribute("aria-busy");

    if (state.participants.length === 0) {
      container.innerHTML = '<span class="pill-row-empty">No participants</span>';
      return;
    }

    var idx = _indexTasks();

    // In-place patch when structure is unchanged; preserves the open pane.
    var existing = container.querySelectorAll(".pill-wrap[data-pid]");
    if (existing.length === state.participants.length) {
      var canPatch = true;
      for (var k = 0; k < state.participants.length; k++) {
        var p0 = state.participants[k];
        var s0 = pillState(p0, idx);
        if (existing[k].getAttribute("data-pid") !== p0.id ||
            existing[k].getAttribute("data-status") !== s0.status ||
            existing[k].getAttribute("data-active") !== (state.selectedParticipant === p0.id ? "1" : "0") ||
            existing[k].getAttribute("data-agents") !== _agentsAttr(s0) ||
            // A speaker pass starting or finishing must rebuild the badge and fill.
            existing[k].getAttribute("data-speakers") !== s0.speakers.status ||
            // Phase feeds the dot tooltip closure; a flip must rebuild.
            existing[k].getAttribute("data-phase") !== (s0.phase || "") ||
            existing[k].getAttribute("data-offsheet") !== offSheetFlag(p0) ||
            // Studio regeneration flips this with identical status, so the badge would never clear.
            existing[k].getAttribute("data-stale") !== (p0.has_stale_artifacts ? "1" : "0")) {
          canPatch = false; break;
        }
      }
      if (canPatch) {
        for (k = 0; k < state.participants.length; k++) {
          var wrap = existing[k];
          p0 = state.participants[k];
          s0 = pillState(p0, idx);
          var bar = wrap.querySelector(".pill-progress");
          if (bar) bar.style.width = _pillFillPercent(s0) + "%";
        }
        // The open pane tracks state the guard cannot see, so refresh here too.
        if (state.pillOptionsOpen !== null) {
          _refreshPillOptionsContent(state.pillOptionsOpen, idx);
        }
        return;
      }
    }

    // Full rebuild
    var openPid = state.pillOptionsOpen;
    var frag = document.createDocumentFragment();
    state.participants.forEach(function (p) {
      frag.appendChild(buildPillWrap(p, idx));
    });
    container.innerHTML = "";
    container.appendChild(frag);
    // Pane is body-mounted: reposition to the new wrap and re-render.
    if (openPid !== null) {
      var newWrap = _findPillWrap(openPid);
      var floating = document.querySelector("body > .pill-options");
      if (newWrap && floating) {
        // Repositions against the rebuilt wrap on its own.
        _refreshPillOptionsContent(openPid, idx);
      } else {
        closePillOptions();
      }
    }
  }

  // A running speaker pass drives the fill; otherwise transcription does.
  function _pillFillPercent(s) {
    return s.speakers.status === "running" ? s.speakers.progress : s.progress;
  }

  function _refreshPillOptionsContent(pid, idx) {
    var floating = document.querySelector("body > .pill-options[data-pid='" + pid + "']");
    if (!floating) return;
    var p = null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) { p = state.participants[i]; break; }
    }
    if (!p) return;
    var s = pillState(p, idx);
    // Captured before the swap: a rebuild may add the audio-track row.
    var navId = _pillNavCursorId();
    var fresh = buildPillOptions(p, s);
    fresh.setAttribute("data-pid", pid);
    var wrap = _findPillWrap(pid);
    if (_paneShape(fresh) === _paneShape(floating)) {
      // Same rows: swap only the agent section; selects keep their options.
      var oldAgents = floating.querySelector(".pill-options-agents");
      var newAgents = fresh.querySelector(".pill-options-agents");
      if (oldAgents && newAgents) oldAgents.parentNode.replaceChild(newAgents, oldAgents);
      _syncPaneRows(floating, fresh);
      if (wrap) _positionPillOptions(floating, wrap);
      _pillNavRestoreCursor(navId);
      return;
    }
    floating.parentNode.replaceChild(fresh, floating);
    // Fixed-position pane: the fresh node lacks the inline left/top.
    if (wrap) _positionPillOptions(fresh, wrap);
    // The swap discards the painted cursor while pillOptionsCursor stays set.
    _pillNavRestoreCursor(navId);
  }

  // Row identity of a pane: nav ids in order plus the optional rows present.
  function _paneShape(pane) {
    var ids = [];
    var nodes = pane.querySelectorAll("[data-nav-id]");
    for (var i = 0; i < nodes.length; i++) ids.push(nodes[i].getAttribute("data-nav-id"));
    if (pane.querySelector(".pill-options-group")) ids.push("audio-group");
    if (pane.querySelector(".pill-options-range")) ids.push("range");
    return ids.join("|");
  }

  // Copy the poll-driven bits of the top rows without rebuilding them.
  function _syncPaneRows(live, fresh) {
    var liveBox = live.querySelector(".pill-options-check");
    var freshBox = fresh.querySelector(".pill-options-check");
    if (liveBox && freshBox) {
      liveBox.checked = freshBox.checked;
      liveBox.disabled = freshBox.disabled;
    }
    var pairs = [".pill-options-range", ".pill-options-hint"];
    for (var i = 0; i < pairs.length; i++) {
      var a = live.querySelector(pairs[i]);
      var b = fresh.querySelector(pairs[i]);
      if (a && b && a.textContent !== b.textContent) a.textContent = b.textContent;
    }
  }

  function buildPillWrap(p, idx) {
    var s = pillState(p, idx);
    var wrap = document.createElement("div");
    var isActive = state.selectedParticipant === p.id;
    wrap.className = "pill-wrap";
    wrap.setAttribute("data-pid", p.id);
    wrap.setAttribute("data-status", s.status);
    wrap.setAttribute("data-active", isActive ? "1" : "0");
    wrap.setAttribute("data-agents", _agentsAttr(s));
    wrap.setAttribute("data-speakers", s.speakers.status);
    wrap.setAttribute("data-phase", s.phase || "");
    wrap.setAttribute("data-offsheet", offSheetFlag(p));
    wrap.setAttribute("data-stale", p.has_stale_artifacts ? "1" : "0");
    wrap.appendChild(buildPill(p, s, isActive));
    wrap.appendChild(buildPillDots(p, s));
    if (state.pillOptionsOpen === p.id) {
      wrap.classList.add("pill-wrap--options-open");
    }
    wrap.addEventListener("mouseenter", function () {
      maybeWarmOnPillHover(p, s);
    });
    return wrap;
  }

  function _dotStateLabel(st) {
    if (st === "done") return "done";
    if (st === "running") return "in progress…";
    if (st === "queued") return "queued";
    if (st === "failed") return "failed";
    return "not started";
  }

  function buildPillDots(p, s) {
    var ag = s.agents;
    var labels = ["Transcription", "Summary", "Citations", "Friction"];
    var keys = ["transcription", "summary", "citations", "friction"];
    var anyActive = false;
    for (var i = 0; i < keys.length; i++) {
      if (ag[keys[i]] !== "idle") { anyActive = true; break; }
    }
    if (!anyActive) {
      var empty = document.createElement("div");
      empty.className = "pill-dots pill-dots--empty";
      return empty;
    }
    var row = document.createElement("div");
    row.className = "pill-dots";
    for (var j = 0; j < keys.length; j++) {
      var dot = document.createElement("span");
      dot.className = "pill-dot pill-dot--" + ag[keys[j]];
      row.appendChild(dot);
    }
    attachHoverTooltip(row, function () {
      var lines = [];
      for (var k = 0; k < keys.length; k++) {
        var label = _dotStateLabel(ag[keys[k]]);
        // Dot stays --running, but the wording must agree with the trigger.
        if (keys[k] === "transcription" && s.status === "cancelling") {
          label = "stopping…";
        } else if (keys[k] === "transcription" && ag[keys[k]] === "running" &&
            s.phase === "loading_model") {
          label = "loading model…";
        }
        lines.push(labels[k] + ": " + label);
      }
      return lines.join("\n");
    }, { multiline: true, align: "center" });
    return row;
  }

  function buildPill(p, s, isActive) {
    var pill = document.createElement("div");
    var classes = ["pill", "pill--" + s.status];
    if (isActive) classes.push("pill--active");
    pill.className = classes.join(" ");
    pill.setAttribute("role", "tab");
    pill.setAttribute("aria-selected", isActive ? "true" : "false");

    // Progress fill (background layer)
    var prog = document.createElement("div");
    prog.className = "pill-progress";
    var spkLive = s.speakers.status === "running" || s.speakers.status === "queued";
    if (spkLive) {
      classes.push("pill--speakers-" + s.speakers.status);
      pill.className = classes.join(" ");
      prog.classList.add("pill-progress--speakers");
    }
    prog.style.width = _pillFillPercent(s) + "%";
    pill.appendChild(prog);

    // Trigger — status icon doubles as action button (hover swaps the glyph)
    pill.appendChild(buildPillTrigger(p, s));

    var idSpan = document.createElement("span");
    idSpan.className = "pill-id";
    idSpan.textContent = p.id;
    pill.appendChild(idSpan);

    // Off-sheet marker: a bare glyph, since it sits on every non-sheet pill.
    if (offSheetFlag(p) === "1") {
      var offSheet = document.createElement("span");
      offSheet.className = "pill-offsheet-badge";
      offSheet.setAttribute("aria-label", "Not in sheet");
      // data-tooltip matches the stale badge; one pill must not mix tooltip systems.
      offSheet.setAttribute("data-tooltip", "Source video found on disk; not a column in the loaded sheet");
      pill.appendChild(offSheet);
    }

    // Stale badge (inline)
    if (p.has_stale_artifacts) {
      var stale = document.createElement("span");
      stale.className = "pill-stale-badge";
      stale.textContent = "stale";
      stale.setAttribute("data-tooltip", "Artifacts built from an older transcript");
      pill.appendChild(stale);
    }

    // Speaker pass badge; the trigger stays on transcription.
    if (spkLive) {
      var spkBadge = document.createElement("span");
      spkBadge.className = "pill-speakers-badge cg-shimmer";
      spkBadge.textContent = "speakers";
      spkBadge.setAttribute("data-tooltip",
        s.speakers.status === "running" ? "Detecting speakers…" : "Speaker detection queued");
      pill.appendChild(spkBadge);
    }

    // Chevron for options pane
    var chevBtn = document.createElement("button");
    chevBtn.type = "button";
    chevBtn.className = "pill-chevron-btn";
    chevBtn.setAttribute("aria-label", "Transcription options");
    chevBtn.setAttribute("aria-expanded", state.pillOptionsOpen === p.id ? "true" : "false");
    // Alt-hold hint: O opens the selected participant, so only the active pill.
    if (isActive) chevBtn.setAttribute("data-hotkey", "transcripts.pillMenu");
    var chev = document.createElement("span");
    chev.className = "pill-chevron";
    chevBtn.appendChild(chev);
    chevBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      togglePillOptions(p.id);
    });
    pill.appendChild(chevBtn);

    // Clicking the pill body (not its buttons) selects the participant
    pill.addEventListener("click", function () {
      if (p.id !== state.selectedParticipant) selectParticipant(p.id);
    });

    return pill;
  }

  function buildPillTrigger(p, s) {
    var cfg = PILL_TRIGGER[s.status] || PILL_TRIGGER.idle;
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "pill-trigger pill-trigger--" + s.status;
    btn.setAttribute("aria-label", cfg.label);
    btn.setAttribute("data-tooltip", cfg.label);
    // Inert while the cancel is in flight; tokens.css keeps the tooltip reachable.
    if (s.status === "cancelling") btn.setAttribute("disabled", "disabled");

    // Two stacked icons; CSS hides one and shows the other on hover/focus.
    var rest = document.createElement("span");
    rest.className = "pill-trigger-icon pill-trigger-icon--rest";
    applyMaskIcon(rest, "url(" + cfg.rest + ")");
    btn.appendChild(rest);

    var hover = document.createElement("span");
    hover.className = "pill-trigger-icon pill-trigger-icon--hover";
    applyMaskIcon(hover, "url(" + cfg.hover + ")");
    btn.appendChild(hover);

    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      if (cfg.action === "cancel") {
        cancelTranscribeTask(s.taskId, s.progress);
      } else if (cfg.action === "retranscribe") {
        startTranscribe(p.id, true);
      } else if (cfg.action === "transcribe") {
        // Explicit, not else: falling through would turn a cancel click into a start.
        startTranscribe(p.id, false);
      }
    });
    return btn;
  }

  // ---- Participant pill popover (model/lang + agent dependency rows) ----

  function buildPillOptions(p, s) {
    var pane = document.createElement("div");
    pane.className = "pill-options";
    pane.addEventListener("click", function (e) { e.stopPropagation(); });

    var ov = state.pillOverrides[p.id] || {};

    // Model row
    var modelRow = document.createElement("div");
    modelRow.className = "pill-options-row";
    var modelLabel = document.createElement("label");
    modelLabel.textContent = "Model";
    var modelSelect = document.createElement("select");
    // data-nav-id keys the cursor to identity, not position.
    modelSelect.setAttribute("data-nav-id", "model");
    modelSelect.innerHTML = '<option value="">Loading…</option>';
    _trFetchModels().then(function (data) {
      var models = (data && data.whisper && data.whisper.models) || [];
      if (models.length === 0) {
        modelSelect.innerHTML = '<option value="">Default</option>';
        return;
      }
      var defaultName = "";
      var opts = '<option value="">Default</option>';
      models.forEach(function (m) {
        if (m.selected) defaultName = m.name;
        var label = m.name + (m.cached === false ? " (not downloaded)" : "");
        opts += '<option value="' + escapeHtml(m.name) + '">' + escapeHtml(label) + '</option>';
      });
      modelSelect.innerHTML = opts;
      modelSelect.options[0].textContent = "Default (" + (defaultName || "base") + ")";
      modelSelect.value = ov.model || "";
    });
    modelSelect.addEventListener("change", function () {
      _setOverride(p.id, "model", this.value);
    });
    modelRow.appendChild(modelLabel);
    modelRow.appendChild(modelSelect);
    pane.appendChild(modelRow);

    // Language row
    var langRow = document.createElement("div");
    langRow.className = "pill-options-row";
    var langLabel = document.createElement("label");
    langLabel.textContent = "Language";
    var langSelect = document.createElement("select");
    langSelect.setAttribute("data-nav-id", "language");
    var lh = "";
    for (var i = 0; i < PILL_LANGUAGES.length; i++) {
      var L = PILL_LANGUAGES[i];
      lh += '<option value="' + escapeHtml(L.code) + '">' + escapeHtml(L.label) + '</option>';
    }
    langSelect.innerHTML = lh;
    langSelect.value = ov.language || "";
    langSelect.addEventListener("change", function () {
      _setOverride(p.id, "language", this.value);
    });
    langRow.appendChild(langLabel);
    langRow.appendChild(langSelect);
    pane.appendChild(langRow);

    // Multitrack only, rendered off the warm cache so the poll cannot strobe it.
    var audioInfo = TS.audioInfoCached(p.id, p.video_version);
    if (audioInfo) {
      if (audioInfo.count > 1) pane.appendChild(buildAudioTrackRow(p, audioInfo, ov));
    } else if (p.has_video) {
      TS._trFetchAudioInfo(p.id, p.video_version).then(function (info) {
        // The captured pane may already be detached; appending would drop the row.
        var live = document.querySelector("body > .pill-options[data-pid='" + p.id + "']");
        if (live !== pane || !info || info.count <= 1) return;
        // The row lands above the agent buttons, shifting a painted cursor.
        var navId = _pillNavCursorId();
        pane.insertBefore(
          buildAudioTrackRow(p, info, state.pillOverrides[p.id] || {}),
          pane.querySelector(".pill-options-speakers") || pane.querySelector(".pill-options-agents")
        );
        _pillNavRestoreCursor(navId);
      });
    }

    // Range row only with in/out markers; the poll keeps it tracking edits.
    var mk = TS.getStoredMarkersFor ? TS.getStoredMarkersFor(p.id) : null;
    if (mk && (mk.in !== null || mk.out !== null)) {
      var rangeRow = document.createElement("div");
      rangeRow.className = "pill-options-row";
      var rangeLabel = document.createElement("label");
      rangeLabel.textContent = "Range";
      var rangeValue = document.createElement("span");
      rangeValue.className = "pill-options-range";
      rangeValue.textContent =
        (mk.in !== null ? formatTime(mk.in, { decimals: 1 }) : "start") +
        " – " +
        (mk.out !== null ? formatTime(mk.out, { decimals: 1 }) : "end");
      var rangeClear = document.createElement("button");
      rangeClear.className = "btn btn-small";
      rangeClear.setAttribute("data-nav-id", "range-clear");
      rangeClear.textContent = "Clear";
      rangeClear.addEventListener("click", function () {
        if (TS.clearMarkersFor) TS.clearMarkersFor(p.id);
        // Full re-render, not a pane refresh: it drops the Range row immediately.
        renderPills();
      });
      rangeRow.appendChild(rangeLabel);
      rangeRow.appendChild(rangeValue);
      rangeRow.appendChild(rangeClear);
      pane.appendChild(rangeRow);
    }

    pane.appendChild(buildSpeakersRow(p));

    // Agent rows with dependency gating; re-running summary cascades to citations server-side.
    pane.appendChild(buildPillAgentsSection(p, s));

    return pane;
  }

  // Per-participant switch; unset follows the global TRANSCRIBE_SPEAKERS.
  function buildSpeakersRow(p) {
    var row = document.createElement("div");
    row.className = "pill-options-row pill-options-speakers";
    var label = document.createElement("label");
    label.textContent = "Speakers";
    label.htmlFor = "pillSpeakers-" + p.id;
    var box = document.createElement("input");
    box.type = "checkbox";
    box.id = label.htmlFor;
    box.className = "pill-options-check";
    box.setAttribute("data-nav-id", "speakers");
    box.checked = speakersEnabledFor(p);
    if (state.speakerModel === false) {
      box.disabled = true;
      row.setAttribute("data-tooltip", "Speaker model is not installed");
    } else {
      row.setAttribute("data-tooltip", "Label lines by detected speaker");
    }
    box.addEventListener("change", function () {
      box.disabled = true; // the reload rebuilds the pane
      setSpeakersEnabled(p.id, box.checked);
    });
    row.appendChild(label);
    row.appendChild(box);
    return row;
  }

  // Track labels can be long and the popover is 220px, so clip here.
  function _trackOptionLabel(track, index) {
    var label = String((track && track.label) || "Track " + (index + 1));
    if (label.length > 20) label = label.slice(0, 19) + "…";
    return "Track " + (index + 1) + " · " + label;
  }

  function buildAudioTrackRow(p, info, ov) {
    // Wrapper, not a bare row: the hint must stack beneath the select.
    var group = document.createElement("div");
    group.className = "pill-options-group";
    var row = document.createElement("div");
    row.className = "pill-options-row";
    var label = document.createElement("label");
    label.textContent = "Audio track";
    var select = document.createElement("select");
    select.setAttribute("data-nav-id", "audio-track");
    var auto = info.tracks[info.auto] || null;
    // The server owns the auto-pick heuristic; this only labels it.
    var html = '<option value="">' +
      escapeHtml("Auto (" + _trackOptionLabel(auto, info.auto) + ")") + "</option>";
    for (var i = 0; i < info.tracks.length; i++) {
      html += '<option value="' + i + '">' +
        escapeHtml(_trackOptionLabel(info.tracks[i], i)) + "</option>";
    }
    select.innerHTML = html;
    select.value = ov.audioTrack || "";
    select.addEventListener("change", function () {
      _setOverride(p.id, "audioTrack", this.value);
    });
    row.appendChild(label);
    row.appendChild(select);
    group.appendChild(row);

    // What the last run used: the only way to explain a changed transcript.
    if (p.audio_track_label) {
      var hint = document.createElement("div");
      hint.className = "pill-options-hint";
      hint.textContent = "Last run: " +
        _trackOptionLabel({ label: p.audio_track_label }, p.audio_index || 0);
      group.appendChild(hint);
    }
    return group;
  }

  function buildPillAgentsSection(p, s) {
    var section = document.createElement("div");
    section.className = "pill-options-agents";

    // 1. Transcription
    section.appendChild(buildAgentRow({
      pid: p.id,
      label: "Transcription",
      agent: "transcription",
      depMet: true,
      agentState: s.agents.transcription,
      // Reads the trigger flag so an optimistic "Stopping…" survives the next poll.
      cancelPending: s.status === "cancelling",
      hasResult: !!p.has_transcript,
      cascadeWarning: !!(p.agents && (p.agents.summary === "done" || p.agents.citations === "done")),
      onStart: function () { startTranscribe(p.id, !!p.has_transcript); },
      onStop: function () { cancelTranscribeTask(s.taskId, s.progress); },
    }));

    // 2. Summary
    section.appendChild(buildAgentRow({
      pid: p.id,
      label: "Summary",
      agent: "summary",
      aiBadge: true,
      depLabel: "transcription",
      depMet: s.agents.transcription === "done",
      agentState: s.agents.summary,
      hasResult: !!(p.agents && p.agents.summary === "done"),
      cascadeWarning: !!(p.agents && p.agents.citations === "done"),
      onStart: function () {
        ensureAgentModelInstalled("summary").then(function (ok) {
          if (!ok) return;
          apiPost("api/agent/summary/" + p.id + "/regenerate", {}).then(function () {
            _refreshAgentStateNow();
          }).catch(function () {
            showToast("Failed to start summary");
          });
        });
      },
      onStop: function () {
        apiPost("api/agent/summary/" + p.id + "/stop", {}).then(function () {
          _refreshAgentStateNow();
        }).catch(function () {
          showToast("Failed to stop summary");
        });
      },
    }));

    // 3. Citations
    section.appendChild(buildAgentRow({
      pid: p.id,
      label: "Citations",
      agent: "citations",
      aiBadge: true,
      depLabel: "summary",
      depMet: s.agents.summary === "done",
      agentState: s.agents.citations,
      hasResult: !!(p.agents && p.agents.citations === "done"),
      cascadeWarning: false,
      onStart: function () {
        ensureAgentModelInstalled("citations").then(function (ok) {
          if (!ok) return;
          apiPost("api/agent/citations/" + p.id + "/regenerate", {}).then(function () {
            _refreshAgentStateNow();
          }).catch(function () {
            showToast("Failed to start citations");
          });
        });
      },
      onStop: function () {
        apiPost("api/agent/citations/" + p.id + "/stop", {}).then(function () {
          _refreshAgentStateNow();
        }).catch(function () {
          showToast("Failed to stop citations");
        });
      },
    }));

    // 4. Friction — depends on summary only (independent of citations).
    section.appendChild(buildAgentRow({
      pid: p.id,
      label: "Friction",
      agent: "friction",
      aiBadge: true,
      depLabel: "summary",
      depMet: s.agents.summary === "done",
      agentState: s.agents.friction,
      hasResult: !!(p.agents && p.agents.friction === "done"),
      cascadeWarning: false,
      onStart: function () {
        ensureAgentModelInstalled("friction").then(function (ok) {
          if (!ok) return;
          apiPost("api/agent/friction/" + p.id + "/regenerate", {}).then(function () {
            _refreshAgentStateNow();
            // loadFriction lives in the agents satellite (loads after this one).
            if (state.selectedParticipant === p.id && TS.loadFriction) TS.loadFriction(p.id);
          }).catch(function () {
            showToast("Failed to start friction");
          });
        });
      },
      onStop: function () {
        apiPost("api/agent/friction/" + p.id + "/stop", {}).then(function () {
          _refreshAgentStateNow();
          if (state.selectedParticipant === p.id && TS.loadFriction) TS.loadFriction(p.id);
        }).catch(function () {
          showToast("Failed to stop friction");
        });
      },
    }));

    // 5. Speakers — only while switched on; needs a finished transcript.
    if (speakersEnabledFor(p)) {
      section.appendChild(buildAgentRow({
        pid: p.id,
        label: "Speakers",
        agent: "speakers",
        depLabel: "transcription",
        depMet: s.agents.transcription === "done",
        agentState: s.agents.speakers,
        hasResult: !!(p.speakers && p.speakers.count > 0),
        cascadeWarning: false,
        onStart: function () { regenerateSpeakers(p.id); },
        onStop: function () { stopSpeakers(p.id); },
      }));
    }

    return section;
  }

  function buildAgentRow(opts) {
    var row = document.createElement("div");
    row.className = "pill-options-row pill-options-agent-row";
    row.setAttribute("data-agent", opts.agent);

    var label = document.createElement("label");
    label.textContent = opts.label;
    row.appendChild(label);

    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn-small pill-agent-btn";
    btn.setAttribute("data-nav-id", "agent:" + opts.agent);
    // Agent rows carry the badge; its own node so Run/Stop swaps leave it.
    if (opts.aiBadge) {
      var badge = document.createElement("span");
      badge.className = "ai-agent-badge";
      badge.setAttribute("aria-hidden", "true");
      btn.appendChild(badge);
    }
    var btnLabel = document.createElement("span");
    btnLabel.className = "agent-run-label";
    btn.appendChild(btnLabel);
    // Alt-hold hint: digits 1-5 route to these rows while the dropdown is open.
    var agentIdx = PILL_AGENT_ORDER.indexOf(opts.agent);
    if (agentIdx >= 0 && agentIdx < 9) {
      btn.setAttribute("data-hotkey", "transcripts.markCategory");
      btn.setAttribute("data-hotkey-combo", String(agentIdx));
    }

    var running = opts.agentState === "running";
    var mode = "start"; // start | stop | disabled
    var title = "";

    if (opts.cancelPending) {
      // Disabled rather than "stop": the label must come from the shared flag.
      btnLabel.textContent = "Stopping…";
      btn.classList.add("pill-agent-btn--stop");
      mode = "disabled";
    } else if (running) {
      btnLabel.textContent = "Stop";
      btn.classList.add("pill-agent-btn--stop");
      mode = "stop";
    } else if (!opts.depMet) {
      btnLabel.textContent = "Run";
      mode = "disabled";
      title = "Requires " + opts.depLabel + " to finish first";
    } else if (opts.hasResult) {
      btnLabel.textContent = "Re-run";
      if (opts.cascadeWarning) {
        title = opts.agent === "transcription"
          ? "Re-transcribing invalidates Summary and Citations"
          : "Re-running will also re-run Citations";
      }
    } else {
      btnLabel.textContent = "Run";
    }
    // Fallback wording; the dependency and cascade warnings above keep precedence.
    if (!title && opts.aiBadge) title = "Runs a local AI thinking agent";

    if (mode === "disabled") btn.setAttribute("disabled", "disabled");
    if (title) btn.setAttribute("data-tooltip", title);

    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      if (btn.hasAttribute("disabled")) return;
      // Optimistic UI swap; the next poll re-renders via _refreshPillOptionsContent.
      if (mode === "stop") {
        btnLabel.textContent = "Stopping…";
        btn.setAttribute("disabled", "disabled");
        opts.onStop();
      } else {
        btnLabel.textContent = "Starting…";
        btn.setAttribute("disabled", "disabled");
        opts.onStart();
      }
    });

    row.appendChild(btn);
    return row;
  }

  function _setOverride(pid, key, value) {
    if (!state.pillOverrides[pid]) state.pillOverrides[pid] = {};
    if (value) state.pillOverrides[pid][key] = value;
    else delete state.pillOverrides[pid][key];
  }

  function _findPillWrap(pid) {
    var container = qs("#participantPills");
    if (!container) return null;
    var wraps = container.querySelectorAll(".pill-wrap[data-pid]");
    for (var i = 0; i < wraps.length; i++) {
      if (wraps[i].getAttribute("data-pid") === pid) return wraps[i];
    }
    return null;
  }

  function _positionPillOptions(pane, wrap) {
    positionPopoverAnchored(pane, wrap.getBoundingClientRect());
  }

  function _closeOpenPaneInDom() {
    if (state.pillOptionsOpen === null) return;
    var wrap = _findPillWrap(state.pillOptionsOpen);
    if (wrap) {
      wrap.classList.remove("pill-wrap--options-open");
      var chev = wrap.querySelector(".pill-chevron-btn");
      if (chev) chev.setAttribute("aria-expanded", "false");
    }
    var floating = document.querySelector("body > .pill-options");
    if (floating) floating.remove();
    if (state.pillOptionsReposition) {
      window.removeEventListener("resize", state.pillOptionsReposition);
      window.removeEventListener("scroll", state.pillOptionsReposition, true);
      state.pillOptionsReposition = null;
    }
    state.pillOptionsOpen = null;
    state.pillOptionsCursor = -1;
  }

  function togglePillOptions(pid) {
    // Close whichever pane is currently open (possibly same pill → toggle off)
    var wasOpen = state.pillOptionsOpen === pid;
    _closeOpenPaneInDom();
    if (wasOpen) return;

    var wrap = _findPillWrap(pid);
    if (!wrap) return;
    var p = null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) { p = state.participants[i]; break; }
    }
    if (!p) return;
    var s = pillState(p, _indexTasks());

    // Body-mounted so it escapes the pill row's overflow clipping.
    var pane = buildPillOptions(p, s);
    pane.setAttribute("data-pid", pid);
    document.body.appendChild(pane);
    _positionPillOptions(pane, wrap);

    wrap.classList.add("pill-wrap--options-open");
    var chev = wrap.querySelector(".pill-chevron-btn");
    if (chev) chev.setAttribute("aria-expanded", "true");
    state.pillOptionsOpen = pid;
    state.pillOptionsCursor = -1; // no keyboard cursor until an arrow is pressed

    var reposition = function () {
      var w = _findPillWrap(pid);
      var floating = document.querySelector("body > .pill-options[data-pid='" + pid + "']");
      if (w && floating) _positionPillOptions(floating, w);
    };
    state.pillOptionsReposition = reposition;
    window.addEventListener("resize", reposition);
    window.addEventListener("scroll", reposition, true);

    if (state.transcribePrewarm === "queue_open") {
      tryPostTranscriptionWarmup();
    }
  }

  function closePillOptions() {
    _closeOpenPaneInDom();
  }

  function initPillOutsideClick() {
    document.addEventListener("click", function (e) {
      if (state.pillOptionsOpen === null) return;
      var wrap = _findPillWrap(state.pillOptionsOpen);
      if (wrap && wrap.contains(e.target)) return;
      var floating = document.querySelector("body > .pill-options");
      if (floating && floating.contains(e.target)) return;
      closePillOptions();
    });
    window.addEventListener("pagehide", closePillOptions);
  }

  function initPillWheelScroll() {
    var el = qs("#participantPills");
    if (!el) return;
    clipgenWheelToHorizontal(el);
  }

  function startTranscribe(pid, force) {
    transcribeParticipants([pid], force);
  }

  // Single cancel path: the flag drops before the request so everything stops together.
  function cancelTranscribeTask(taskId, progress) {
    if (!taskId || state.cancellingTasks[taskId]) return;
    state.cancellingTasks[taskId] = { at: Date.now(), progress: progress || 0 };
    renderPills();
    updateTranscribeFill();
    refreshTranscribeWording();
    apiDelete("api/transcribe/" + taskId).then(function () {
      pollTaskStatus();
    }).catch(function () {
      // Only toast when we are the ones putting the state back.
      if (!state.cancellingTasks[taskId]) return;
      delete state.cancellingTasks[taskId];
      renderPills();
      updateTranscribeFill();
      refreshTranscribeWording();
      showToast("Failed to cancel transcription");
    });
  }

  // The server has no in-flight guard, so two POSTs would make two tasks.
  var _transcribeInFlight = {};

  function _clearTranscribeInFlight(pids) {
    for (var i = 0; i < pids.length; i++) delete _transcribeInFlight[pids[i]];
  }

  // One POST for many participants, each carrying its own dropdown overrides.
  function transcribeParticipants(pids, force) {
    var fresh = [];
    var overrides = {};
    for (var i = 0; i < pids.length; i++) {
      var pid = pids[i];
      if (_transcribeInFlight[pid]) continue;
      _transcribeInFlight[pid] = true;
      fresh.push(pid);
      var ov = state.pillOverrides[pid];
      // A <select> gives track 0 as "0"; absent, not falsy, means auto-detect.
      if (ov && (ov.model || ov.language || ov.audioTrack)) {
        overrides[pid] = {};
        if (ov.model) overrides[pid].model = ov.model;
        if (ov.language) overrides[pid].language = ov.language;
        if (ov.audioTrack) overrides[pid].audio_index = parseInt(ov.audioTrack, 10);
      }
      // Per-participant range markers from sessionStorage; omitted when unset.
      var mk = TS.getStoredMarkersFor ? TS.getStoredMarkersFor(pid) : null;
      if (mk && (mk.in !== null || mk.out !== null)) {
        overrides[pid] = overrides[pid] || {};
        if (mk.in !== null) overrides[pid].start_seconds = mk.in;
        if (mk.out !== null) overrides[pid].end_seconds = mk.out;
      }
    }
    if (!fresh.length) return;
    _postTranscribe(fresh, force, overrides, false);
  }

  // The server decides if a model is cached; every exit releases the claim except retry.
  function _postTranscribe(pids, force, overrides, allowDownload) {
    var body = { participants: pids, force: force };
    if (overrides && Object.keys(overrides).length > 0) body.overrides = overrides;
    if (allowDownload) body.allow_download = true;
    apiPost("api/transcribe", body).then(function (data) {
      if (!data.ok) {
        if (data.reason === "model_not_cached" && data.uncached && data.uncached.length) {
          _confirmUncachedWhisperModels(data.uncached).then(function (ok) {
            if (ok) {
              _postTranscribe(pids, force, overrides, true);
              return;
            }
            _clearTranscribeInFlight(pids);
          });
          return;
        }
        _clearTranscribeInFlight(pids);
        showToast("Failed to enqueue transcription");
        return;
      }
      // Adopt the server records now: they paint pills and gate the next enqueue.
      state.tasks = state.tasks.concat(data.tasks);
      _clearTranscribeInFlight(pids);
      showToast("Enqueued " + clipgenPluralUnit(data.tasks.length, "transcription", "transcriptions"));
      renderPills();
      startPolling();
      pollTaskStatus();
    }).catch(function () {
      // Without this the pids stay claimed and can never be re-enqueued.
      _clearTranscribeInFlight(pids);
      showToast("Failed to enqueue transcription");
    });
  }

  function isPillMenuOpen() {
    return state.pillOptionsOpen !== null;
  }

  // Agent rows in display order; digits 1-5 map onto these.
  var PILL_AGENT_ORDER = ["transcription", "summary", "citations", "friction", "speakers"];

  function triggerPillOption(n) {
    if (state.pillOptionsOpen === null) return false;
    var agent = PILL_AGENT_ORDER[n - 1];
    if (!agent) return false;
    var pane = document.querySelector("body > .pill-options");
    if (!pane) return false;
    var row = pane.querySelector('.pill-options-agent-row[data-agent="' + agent + '"]');
    var btn = row && row.querySelector(".pill-agent-btn");
    if (!btn || btn.hasAttribute("disabled")) return false;
    btn.click();
  }

  // Keyboard navigation: a painted cursor roves the selects and agent buttons.

  function pillNavControls() {
    var pane = document.querySelector("body > .pill-options");
    if (!pane) return [];
    return Array.prototype.slice.call(
      pane.querySelectorAll(".pill-options-row select, .pill-options-check, .pill-agent-btn")
    );
  }

  // Cursor carried across rebuilds by control identity, since index arithmetic cannot work.
  function _pillNavCursorId() {
    if (state.pillOptionsCursor < 0) return null;
    var cur = pillNavControls()[state.pillOptionsCursor];
    return cur ? cur.getAttribute("data-nav-id") : null;
  }

  function _pillNavRestoreCursor(navId) {
    if (state.pillOptionsCursor < 0) return; // nothing painted; nothing to keep
    if (navId) {
      var controls = pillNavControls();
      for (var i = 0; i < controls.length; i++) {
        if (controls[i].getAttribute("data-nav-id") === navId) {
          state.pillOptionsCursor = i;
          break;
        }
      }
    }
    // Repaints, and clamps if that control is gone from the rebuilt pane.
    pillNavPaint();
  }

  function pillNavPaint() {
    var pane = document.querySelector("body > .pill-options");
    if (!pane) return;
    var prev = pane.querySelectorAll(".pill-nav-cursor");
    for (var i = 0; i < prev.length; i++) prev[i].classList.remove("pill-nav-cursor");
    if (state.pillOptionsCursor < 0) return;
    var controls = pillNavControls();
    if (!controls.length) return;
    state.pillOptionsCursor = Math.max(0, Math.min(state.pillOptionsCursor, controls.length - 1));
    var cur = controls[state.pillOptionsCursor];
    if (cur) {
      cur.classList.add("pill-nav-cursor");
      if (cur.scrollIntoView) cur.scrollIntoView({ block: "nearest" });
    }
  }

  function pillNavMove(delta) {
    var controls = pillNavControls();
    if (!controls.length) return;
    if (state.pillOptionsCursor < 0) {
      state.pillOptionsCursor = delta > 0 ? 0 : controls.length - 1;
    } else {
      state.pillOptionsCursor = Math.max(
        0, Math.min(state.pillOptionsCursor + delta, controls.length - 1)
      );
    }
    pillNavPaint();
  }

  function pillNavAdjust(dir) {
    var controls = pillNavControls();
    if (state.pillOptionsCursor < 0 || !controls.length) return;
    var cur = controls[state.pillOptionsCursor];
    if (cur && cur.tagName === "SELECT" && cur.options.length) {
      cur.selectedIndex = Math.max(0, Math.min(cur.selectedIndex + dir, cur.options.length - 1));
      cur.dispatchEvent(new Event("change", { bubbles: true }));
    } else if (cur && cur.type === "checkbox" && !cur.disabled) {
      // Right = on, Left = off; click() toggles and fires change.
      if (cur.checked !== (dir > 0)) cur.click();
    }
  }

  function pillNavActivate() {
    var controls = pillNavControls();
    if (state.pillOptionsCursor < 0 || !controls.length) return;
    var cur = controls[state.pillOptionsCursor];
    if (cur && cur.classList.contains("pill-agent-btn") && !cur.hasAttribute("disabled")) {
      cur.click();
    } else if (cur && cur.type === "checkbox" && !cur.disabled) {
      cur.click();
    }
  }

  // ---- Published back to the hub (loadParticipants/selectParticipant/poller
  // render pills; boot wires the listeners) ----
  TS.renderPills = renderPills;
  TS.transcribeParticipants = transcribeParticipants; // hub (Transcribe All quick action)
  TS.initPillOutsideClick = initPillOutsideClick;
  TS.initPillWheelScroll = initPillWheelScroll;
  TS.togglePillOptions = togglePillOptions; // video (O hotkey)
  TS.closePillOptions = closePillOptions; // video (Escape closes the dropdown)
  TS.isPillMenuOpen = isPillMenuOpen; // video (digit branch)
  TS.triggerPillOption = triggerPillOption; // video (1–4 while dropdown open)
  TS.pillNavMove = pillNavMove; // video (Up/Down while dropdown open)
  TS.pillNavAdjust = pillNavAdjust; // video (Left/Right while dropdown open)
  TS.pillNavActivate = pillNavActivate; // video (Enter while dropdown open)
  TS.trackOptionLabel = _trackOptionLabel; // hub (Normalize Audio track checkboxes)
})();
