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
 * state). Plain utils.js globals (qs/el/escapeHtml/apiPost/apiDelete/
 * clipgenPluralUnit/applyMaskIcon/attachHoverTooltip/positionPopoverAnchored)
 * are reached via the scope chain.
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
    startPolling = TS.startPolling,
    _refreshAgentStateNow = TS._refreshAgentStateNow,
    _trFetchModels = TS._trFetchModels,
    ensureAgentModelInstalled = TS.ensureAgentModelInstalled,
    _confirmUncachedWhisperModels = TS._confirmUncachedWhisperModels;

  // Icon shown on the trigger by default (left), and on hover (right).
  // The trigger click always invokes the appropriate action for the status.
  var PILL_TRIGGER = {
    idle: { rest: "icons/microphone.svg", hover: "icons/microphone.svg", label: "Transcribe", action: "transcribe" },
    failed: { rest: "icons/exclamation-triangle.svg", hover: "icons/microphone.svg", label: "Retry transcription", action: "transcribe" },
    queued: { rest: "icons/clock.svg", hover: "icons/stop-circle.svg", label: "Cancel", action: "cancel" },
    running: { rest: "icons/arrow-path.svg", hover: "icons/stop-circle.svg", label: "Cancel transcription", action: "cancel" },
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

  function pillState(p, taskByPid) {
    var task = taskByPid[p.id];
    var status = "idle";
    var progress = 0;
    var taskId = null;

    if (task && (task.status === "running" || task.status === "queued")) {
      status = task.status;
      taskId = task.id;
      if (task.status === "running") progress = Math.round((task.progress || 0) * 100);
    } else if (task && task.status === "failed") {
      status = "failed";
      taskId = task.id;
    } else if (p.has_transcript) {
      status = "completed";
      progress = 100;
    }
    var agents = {
      transcription: _dotStateTranscription(p, task),
      summary: (p.agents && p.agents.summary) || "idle",
      citations: (p.agents && p.agents.citations) || "idle",
      friction: (p.agents && p.agents.friction) || "idle",
    };
    return { status: status, progress: progress, taskId: taskId, agents: agents };
  }

  function renderPills() {
    var container = qs("#participantPills");
    if (!container) return;
    // Whatever we render now is the real answer, so the boot skeletons (and the
    // aria-busy that announces them) are done.
    container.removeAttribute("aria-busy");

    if (state.participants.length === 0) {
      container.innerHTML = '<span class="pill-row-empty">No participants</span>';
      return;
    }

    var taskByPid = {};
    state.tasks.forEach(function (t) {
      if (!taskByPid[t.participant] || t.created_at > taskByPid[t.participant].created_at) {
        taskByPid[t.participant] = t;
      }
    });

    // In-place patch when structure unchanged (avoids layout thrash + preserves
    // the open options pane across polling ticks)
    var existing = container.querySelectorAll(".pill-wrap[data-pid]");
    if (existing.length === state.participants.length) {
      var canPatch = true;
      for (var k = 0; k < state.participants.length; k++) {
        var p0 = state.participants[k];
        var s0 = pillState(p0, taskByPid);
        var agentsAttr = s0.agents.transcription + "," + s0.agents.summary + "," + s0.agents.citations + "," + s0.agents.friction;
        if (existing[k].getAttribute("data-pid") !== p0.id ||
            existing[k].getAttribute("data-status") !== s0.status ||
            existing[k].getAttribute("data-active") !== (state.selectedParticipant === p0.id ? "1" : "0") ||
            existing[k].getAttribute("data-agents") !== agentsAttr) {
          canPatch = false; break;
        }
      }
      if (canPatch) {
        for (k = 0; k < state.participants.length; k++) {
          var wrap = existing[k];
          p0 = state.participants[k];
          s0 = pillState(p0, taskByPid);
          var bar = wrap.querySelector(".pill-progress");
          if (bar) bar.style.width = s0.progress + "%";
        }
        return;
      }
    }

    // Full rebuild
    var openPid = state.pillOptionsOpen;
    var frag = document.createDocumentFragment();
    state.participants.forEach(function (p) {
      frag.appendChild(buildPillWrap(p, taskByPid));
    });
    container.innerHTML = "";
    container.appendChild(frag);
    // The pane is mounted on <body>, not inside the rebuilt wrap — reposition
    // it to the new wrap's rect and re-render its contents so the agent rows
    // reflect the latest state. Tear it down if the pill disappeared.
    if (openPid !== null) {
      var newWrap = _findPillWrap(openPid);
      var floating = document.querySelector("body > .pill-options");
      if (newWrap && floating) {
        _refreshPillOptionsContent(openPid, taskByPid);
        var refreshed = document.querySelector("body > .pill-options[data-pid='" + openPid + "']");
        if (refreshed) _positionPillOptions(refreshed, newWrap);
      } else {
        closePillOptions();
      }
    }
  }

  function _refreshPillOptionsContent(pid, taskByPid) {
    var floating = document.querySelector("body > .pill-options[data-pid='" + pid + "']");
    if (!floating) return;
    var p = null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) { p = state.participants[i]; break; }
    }
    if (!p) return;
    var s = pillState(p, taskByPid);
    // Captured before the swap: the rebuilt pane may carry an audio-track row
    // the old one lacked (the layout arrives asynchronously, and a rebuild can
    // race that fetch), which shifts every control below it down one.
    var navId = _pillNavCursorId();
    var fresh = buildPillOptions(p, s);
    fresh.setAttribute("data-pid", pid);
    floating.parentNode.replaceChild(fresh, floating);
    // The whole pane node is swapped, so a painted keyboard cursor goes with it
    // while state.pillOptionsCursor stays set — and the arrows are still claimed
    // by _pillMenuActive, so it would look like nothing has focus until the next
    // keypress. This poll fires most often while a transcription is running,
    // which is exactly when the dropdown is likely open.
    _pillNavRestoreCursor(navId);
  }

  function buildPillWrap(p, taskByPid) {
    var s = pillState(p, taskByPid);
    var wrap = document.createElement("div");
    var isActive = state.selectedParticipant === p.id;
    wrap.className = "pill-wrap";
    wrap.setAttribute("data-pid", p.id);
    wrap.setAttribute("data-status", s.status);
    wrap.setAttribute("data-active", isActive ? "1" : "0");
    wrap.setAttribute("data-agents", s.agents.transcription + "," + s.agents.summary + "," + s.agents.citations + "," + s.agents.friction);
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
        lines.push(labels[k] + ": " + _dotStateLabel(ag[keys[k]]));
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
    prog.style.width = s.progress + "%";
    pill.appendChild(prog);

    // Trigger — status icon doubles as action button (hover swaps the glyph)
    pill.appendChild(buildPillTrigger(p, s));

    var idSpan = document.createElement("span");
    idSpan.className = "pill-id";
    idSpan.textContent = p.id;
    pill.appendChild(idSpan);

    // Stale badge (inline)
    if (p.has_stale_artifacts) {
      var stale = document.createElement("span");
      stale.className = "pill-stale-badge";
      stale.textContent = "stale";
      stale.setAttribute("data-tooltip", "Artifacts built from an older transcript");
      pill.appendChild(stale);
    }

    // Chevron for options pane
    var chevBtn = document.createElement("button");
    chevBtn.type = "button";
    chevBtn.className = "pill-chevron-btn";
    chevBtn.setAttribute("aria-label", "Transcription options");
    chevBtn.setAttribute("aria-expanded", state.pillOptionsOpen === p.id ? "true" : "false");
    // Alt-hold hint: `O` (pillMenu) opens the options for the *selected*
    // participant, so only the active pill's chevron carries the hint. Pills
    // re-render on selection change, keeping this in sync.
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
        if (s.taskId) {
          apiDelete("api/transcribe/" + s.taskId).then(function () { pollTaskStatus(); });
        }
      } else if (cfg.action === "retranscribe") {
        startTranscribe(p.id, true);
      } else {
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
    // data-nav-id keys the keyboard cursor to a control's identity rather than
    // its position — see _pillNavRestoreCursor.
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

    // Audio track row — multitrack sources only. Rendered synchronously off the
    // hub's warm cache when possible: the 3 s poll rebuilds this whole pane via
    // _refreshPillOptionsContent, and an async insert on every rebuild would
    // strobe the row and shift the keyboard cursor under the user.
    var audioInfo = TS.audioInfoCached(p.id, p.video_version);
    if (audioInfo) {
      if (audioInfo.count > 1) pane.appendChild(buildAudioTrackRow(p, audioInfo, ov));
    } else if (p.has_video) {
      TS._trFetchAudioInfo(p.id, p.video_version).then(function (info) {
        // The pane we captured may already have been swapped out by a poll tick;
        // appending to the detached node would silently drop the row.
        var live = document.querySelector("body > .pill-options[data-pid='" + p.id + "']");
        if (live !== pane || !info || info.count <= 1) return;
        // The row lands *above* the agent buttons, so a cursor painted on one of
        // them would otherwise end up pointing one control too high.
        var navId = _pillNavCursorId();
        pane.insertBefore(
          buildAudioTrackRow(p, info, state.pillOverrides[p.id] || {}),
          pane.querySelector(".pill-options-agents")
        );
        _pillNavRestoreCursor(navId);
      });
    }

    // Agent rows — manual run / re-run / stop controls with dependency gating.
    // Order: Transcription → Summary → Citations. Summary requires
    // transcription; citations requires summary. Re-running summary cascades
    // to citations server-side (see transcripts_server.py).
    pane.appendChild(buildPillAgentsSection(p, s));

    return pane;
  }

  // Track labels come from the container's stream metadata and can be long
  // ("Blackhole 16ch (Participant Mic)"); the popover is 220px wide, so clip
  // them here rather than letting a <select> stretch the pane.
  function _trackOptionLabel(track, index) {
    var label = String((track && track.label) || "Track " + (index + 1));
    if (label.length > 20) label = label.slice(0, 19) + "…";
    return "Track " + (index + 1) + " · " + label;
  }

  function buildAudioTrackRow(p, info, ov) {
    // Wrapper, not a bare row: .pill-options-row is a horizontal flex line, so
    // the "Last run" hint has to stack beneath it rather than sit beside the
    // select. pillNavControls still finds the select via ".pill-options-row select".
    var group = document.createElement("div");
    group.className = "pill-options-group";
    var row = document.createElement("div");
    row.className = "pill-options-row";
    var label = document.createElement("label");
    label.textContent = "Audio track";
    var select = document.createElement("select");
    select.setAttribute("data-nav-id", "audio-track");
    var auto = info.tracks[info.auto] || null;
    // The server computes the auto pick (video.pick_speech_audio_track) so the
    // heuristic has exactly one implementation; this only labels it.
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

    // What the last completed run actually used — the only way to explain a
    // transcript that changed because auto-detect moved off track 1.
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
      hasResult: !!p.has_transcript,
      cascadeWarning: !!(p.agents && (p.agents.summary === "done" || p.agents.citations === "done")),
      onStart: function () { startTranscribe(p.id, !!p.has_transcript); },
      onStop: function () {
        if (s.taskId) {
          apiDelete("api/transcribe/" + s.taskId).then(function () { pollTaskStatus(); });
        }
      },
    }));

    // 2. Summary
    section.appendChild(buildAgentRow({
      pid: p.id,
      label: "Summary",
      agent: "summary",
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
    // Alt-hold hint: while the dropdown is open, digits 1-4 route to these rows
    // (triggerPillOption). Tag each with the markCategory combo index so the
    // "1".."4" chips show (mirrors how Screenspace reuses selectTool for menus).
    var agentIdx = PILL_AGENT_ORDER.indexOf(opts.agent);
    if (agentIdx >= 0 && agentIdx < 9) {
      btn.setAttribute("data-hotkey", "transcripts.markCategory");
      btn.setAttribute("data-hotkey-combo", String(agentIdx));
    }

    var running = opts.agentState === "running";
    var mode = "start"; // start | stop | disabled
    var title = "";

    if (running) {
      btn.textContent = "Stop";
      btn.classList.add("pill-agent-btn--stop");
      mode = "stop";
    } else if (!opts.depMet) {
      btn.textContent = "Run";
      mode = "disabled";
      title = "Requires " + opts.depLabel + " to finish first";
    } else if (opts.hasResult) {
      btn.textContent = "Re-run";
      if (opts.cascadeWarning) {
        title = opts.agent === "transcription"
          ? "Re-transcribing invalidates Summary and Citations"
          : "Re-running will also re-run Citations";
      }
    } else {
      btn.textContent = "Run";
    }

    if (mode === "disabled") btn.setAttribute("disabled", "disabled");
    if (title) btn.setAttribute("data-tooltip", title);

    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      if (btn.hasAttribute("disabled")) return;
      // Optimistic UI swap; the next poll re-renders via _refreshPillOptionsContent.
      if (mode === "stop") {
        btn.textContent = "Stopping…";
        btn.setAttribute("disabled", "disabled");
        opts.onStop();
      } else {
        btn.textContent = "Starting…";
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
    var taskByPid = {};
    state.tasks.forEach(function (t) {
      if (!taskByPid[t.participant] || t.created_at > taskByPid[t.participant].created_at) {
        taskByPid[t.participant] = t;
      }
    });
    var s = pillState(p, taskByPid);

    // Mount the pane on <body> (not inside #participantPills) so it escapes the
    // pill row's overflow clipping and renders above the search bar.
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
    el.addEventListener("wheel", function (e) {
      if (el.scrollWidth > el.clientWidth) {
        e.preventDefault();
        el.scrollLeft += e.deltaY;
      }
    }, { passive: false });
  }

  function startTranscribe(pid, force) {
    var overrides = {};
    var ov = state.pillOverrides[pid];
    // ov.audioTrack is a <select> value, so track 0 arrives as the string "0" —
    // truthy, unlike the number. The server parses it back to an int and treats
    // absent (not falsy) as "auto-detect".
    if (ov && (ov.model || ov.language || ov.audioTrack)) {
      overrides[pid] = {};
      if (ov.model) overrides[pid].model = ov.model;
      if (ov.language) overrides[pid].language = ov.language;
      if (ov.audioTrack) overrides[pid].audio_index = parseInt(ov.audioTrack, 10);
    }
    transcribeParticipants([pid], force, overrides);
  }

  function transcribeParticipants(pids, force, overrides) {
    _postTranscribe(pids, force, overrides, false);
  }

  // POST to the transcribe endpoint. The server is the authority on whether a
  // Whisper model is cached: when it rejects with reason "model_not_cached", we
  // confirm each uncached download with the user and retry with allow_download.
  function _postTranscribe(pids, force, overrides, allowDownload) {
    var body = { participants: pids, force: force };
    if (overrides && Object.keys(overrides).length > 0) body.overrides = overrides;
    if (allowDownload) body.allow_download = true;
    apiPost("api/transcribe", body).then(function (data) {
      if (!data.ok) {
        if (data.reason === "model_not_cached" && data.uncached && data.uncached.length) {
          _confirmUncachedWhisperModels(data.uncached).then(function (ok) {
            if (ok) _postTranscribe(pids, force, overrides, true);
          });
          return;
        }
        showToast("Failed to enqueue transcription");
        return;
      }
      showToast("Enqueued " + clipgenPluralUnit(data.tasks.length, "transcription", "transcriptions"));
      startPolling();
      pollTaskStatus();
    });
  }

  function isPillMenuOpen() {
    return state.pillOptionsOpen !== null;
  }

  // Agent rows in the open participant dropdown, in display order. The digit
  // hotkeys (1–4) map onto these while the dropdown is open (the branch lives in
  // transcripts-video.js's markCategory handler).
  var PILL_AGENT_ORDER = ["transcription", "summary", "citations", "friction"];

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

  // ---- Keyboard navigation inside the open dropdown ----
  // The pane is mounted on <body>; its interactive controls in DOM order are the
  // Model + Language selects, then the four agent buttons. A painted cursor
  // (.pill-nav-cursor) roves them — no real DOM focus, so the arrow hotkeys are
  // not swallowed by the dispatcher's typing gate. Left/Right steps a focused
  // select; Enter runs a focused agent button (digits 1–4 still shortcut them).

  function pillNavControls() {
    var pane = document.querySelector("body > .pill-options");
    if (!pane) return [];
    return Array.prototype.slice.call(
      pane.querySelectorAll(".pill-options-row select, .pill-agent-btn")
    );
  }

  // The pane is rebuilt wholesale every poll tick and can gain the audio-track
  // row between two builds (the layout arrives asynchronously), so the cursor is
  // carried across a rebuild by the control's identity, not its index. Index
  // arithmetic at each mutation site cannot work: _refreshPillOptionsContent
  // replaces the node, so by the time it could count controls the old pane —
  // and the number of rows the cursor was measured against — is already gone.
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
    }
  }

  function pillNavActivate() {
    var controls = pillNavControls();
    if (state.pillOptionsCursor < 0 || !controls.length) return;
    var cur = controls[state.pillOptionsCursor];
    if (cur && cur.classList.contains("pill-agent-btn") && !cur.hasAttribute("disabled")) {
      cur.click();
    }
  }

  // ---- Published back to the hub (loadParticipants/selectParticipant/poller
  // render pills; boot wires the listeners) ----
  TS.renderPills = renderPills;
  TS.initPillOutsideClick = initPillOutsideClick;
  TS.initPillWheelScroll = initPillWheelScroll;
  TS.togglePillOptions = togglePillOptions; // video (O hotkey)
  TS.closePillOptions = closePillOptions; // video (Escape closes the dropdown)
  TS.isPillMenuOpen = isPillMenuOpen; // video (digit branch)
  TS.triggerPillOption = triggerPillOption; // video (1–4 while dropdown open)
  TS.pillNavMove = pillNavMove; // video (Up/Down while dropdown open)
  TS.pillNavAdjust = pillNavAdjust; // video (Left/Right while dropdown open)
  TS.pillNavActivate = pillNavActivate; // video (Enter while dropdown open)
})();
