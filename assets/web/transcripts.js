/* clipgen Transcripts page.
 *
 * Editor for per-participant Whisper transcripts, plus cross-references into
 * Screenspace events and the source spreadsheet. Two long-running side flows
 * piggyback on the same `state`:
 *
 *   - Transcription warmup: a single `tryPostTranscriptionWarmup()` post that
 *     asks the backend to preload the Whisper model. `_transcriptionWarmupPosted`
 *     guards it so we never double-post per page load.
 *   - Summary / citations: Ollama-generated; `_summaryPoller` and
 *     `_citationsPoller` (createPoller handles) poll the backend until the
 *     result lands or the user navigates away.
 */

(function () {
  "use strict";

  var state = {
    participants: [],
    selectedParticipant: null,
    segments: [],
    corrections: [],
    tasks: [],
    // Task ids with a cancel DELETE in flight → { at, progress }. The server
    // keeps reporting a cancelled-but-still-running task as "running" until the
    // worker reaches its next checkpoint — up to a whole cold model load away —
    // so every surface that reads task.status has to ask this too, or it goes on
    // saying "transcribing" for ten seconds after the click.
    //
    // Keyed by task id, never by participant: a re-transcribe issued right after
    // a cancel is a *new* task, and a pid key that had not yet been swept would
    // paint the fresh run as "Cancelling…" and disable its own stop button.
    // Written by the pills satellite, read by pills + video — hence state, not a
    // module var (see agents/skills/carve-satellite/SKILL.md).
    cancellingTasks: {},
    // In/out transcribe-range markers (seconds on the global timeline, null =
    // unset). Owned by the video satellite (set/persist/restore/draw); the
    // pills satellite reads sessionStorage directly for the request bounds.
    inMarker: null,
    outMarker: null,
    searchQuery: "",
    searchResults: null,
    activeSegmentIndex: -1,
    editingTextEl: null,
    pollPoller: null,
    lastMarkCategory: "bookmark",
    streamingParticipant: null,
    ssEvents: [],
    ssEventsLoaded: false,
    sheetRows: [],
    sheetParticipants: [],
    sheetLoaded: false,
    // Whether a spreadsheet is loaded at all — gates the per-pill off-sheet badge.
    hasSheet: false,
    xrefPoller: null,
    xrefEligible: false,
    xrefIndex: { eventsByParticipant: {}, sheetByParticipant: {} },
    summaryEditing: false,
    summaryText: "",
    summaryCitations: null,
    citationsGenerating: false,
    activeTab: "summary",
    frictionData: null,
    // Which participant frictionData belongs to. loadFriction blanks the pane
    // only when this changes, so a same-participant refetch mid-run keeps the
    // deterministic scores rather than emptying the tab until the agent lands.
    frictionPid: null,
    frictionBySegId: {},
    frictionGenerating: false,
    // Server-recorded friction run start (epoch ms) so the elapsed clock
    // survives page navigation; null while idle or for a just-clicked run.
    frictionStartedAt: null,
    // "off" | "highlight" | "isolate" — what the friction filter does to
    // #segmentList. Persisted; owned by transcripts-agents.js, read here and by
    // the video satellite's timeline band.
    frictionMode: "off",
    // Score band the filter keeps, both ends draggable on the histogram. Opens
    // fully open (every segment the scorer flagged) — the histogram makes the
    // distribution visible, so narrowing from everything beats opening
    // pre-filtered to something the user can't see.
    frictionMin: 0,
    frictionMax: 1,
    // Two independent filters, one per evidence source — the keyword scorer
    // labels segments, the agent labels its own moments, and the two never
    // agree on a line. Sharing one dict meant hiding a category on one side
    // silently hid it on the other. Both persist; frictionMomentFilter also
    // carries the "other" bucket for categories the model invented.
    frictionCategoryFilter: null,
    frictionMomentFilter: null,
    // Derived from (threshold, categoryFilter, frictionData, segments) by
    // _recomputeFrictionMatches. The single source every friction consumer reads:
    // segment id -> score for matching segments, the visible moments resolved to
    // segment indices, and segment id -> 1-based moment number for cited rows.
    frictionMatchBySegId: {},
    // The union both evidence sources contribute to, keyed to the strongest
    // score on each line. What the timeline density band draws and hit-tests —
    // an AI-cited line scores 0 with the keyword scorer, so the match map alone
    // left the agent's moments off the band entirely.
    frictionBandBySegId: {},
    frictionVisibleMoments: [],
    frictionCitedBySegId: {},
    // Moments that passed the filter but cite segment ids absent from
    // state.segments (transcript edited since the run). Counted alongside the
    // three maps above so the jump strip can say *why* it came out empty.
    frictionUnsourcedMoments: 0,
    frictionMomentIndex: -1,
    transcribePrewarm: "queue_open",
    modelStatus: null,
    modelFailSince: 0,
    videoPlaying: false,
    videoMuted: false,
    // Audio-track layout for the current participant (from /api/audio-info);
    // feeds the volume popover's detected-track caption + per-track mixer.
    audioTracks: [],
    audioPanel: null, // ClipgenVideoControls audio-popover controller
    videoPlaybackRate: 1,
    ccEnabled: false,
    pipActive: false,
    pipEnabled: true,
    videoCollapsed: false,
    // Lazily-built cache of .segment-row elements (invalidated on re-render);
    // shared by the segments editor (hub) and the video satellite's playhead
    // highlight / marker-click. frictionTooltipShown coordinates the shared
    // #trTooltip between the timeline-canvas hover (video) and the hot-segment
    // friction tooltip (agents), so neither clobbers the other.
    cachedSegmentRows: null,
    frictionTooltipShown: false,
    // Bumped on every participant switch so per-participant fetches that resolve
    // late (loadTranscript, loadSummary, summary/citations/friction polls — now
    // in the agents satellite) can detect they're stale and bail before
    // clobbering the active participant's UI.
    participantReqVer: 0,
    // Keyboard cursor into the open participant-options dropdown (-1 = none);
    // see pillNav* in transcripts-pills.js.
    pillOptionsCursor: -1,
  };

  var _transcriptionWarmupPosted = false;
  // Prewarm never downloads silently: when the model isn't cached we confirm
  // the download with the user. _prewarmDownloadPrompting guards against
  // double-prompting; _prewarmDeclinedModel records the specific model the user
  // declined so we stop re-asking for it — but switching to a different model
  // (or changing TRANSCRIBE_MODEL in settings) still gets its own prompt. The
  // model still loads on demand, with confirmation, at transcribe time.
  var _prewarmDownloadPrompting = false;
  var _prewarmDeclinedModel = null;
  // Last-known TRANSCRIBE_MODEL, so a settings change can reset the prewarm
  // guards for the new model. Seeded once from the model-status hint.
  var _lastTranscribeModel = null;
  var _modelHintPoller = null;
  var _hadActiveTranscriptionLastPoll = false;
  var MODEL_FAIL_GRACE_MS = 10000;

  // ---- Helpers (showToast: 2500ms hide; shared default in utils.js is 3000ms) ----

  var _utilsShowToast = window.showToast;
  function showToast(msg) {
    _utilsShowToast(msg, { durationMs: 2500 });
  }

  // ---- Nav links ----

  function checkNavLinks() {
    apiGet("../api/status").then(function (data) {
      if (data.screenspace || data.studio) {
        state.xrefEligible = true;
        startXrefPolling();
      }
    }).catch(function () {});
  }

  // ---- Cross-reference data ----

  function startXrefPolling() {
    if (!state.xrefEligible || state.xrefPoller) return;
    // Re-arm the sheet leg: this also runs on tab focus, and a spreadsheet may
    // have been opened from another tab since we last looked.
    _sheetXrefIdle = false;
    // createPoller runs loadCrossRefData once immediately (runImmediately
    // default), then every 30s.
    state.xrefPoller = createPoller(loadCrossRefData, 30000, { label: "transcripts.xref" });
    state.xrefPoller.start();
  }

  function stopXrefPolling() {
    if (state.xrefPoller) {
      state.xrefPoller.stop();
      state.xrefPoller = null;
    }
  }

  // A genuinely unreachable sibling tool (network error or non-2xx) flips a
  // per-source flag so the status tooltip can say cross-references are missing,
  // rather than the badges silently never appearing. A successful poll clears it.
  function _markXrefSource(source, failed) {
    var prev = !!(state.xrefErrors && (state.xrefErrors.screenspace || state.xrefErrors.studio));
    if (!state.xrefErrors) state.xrefErrors = { screenspace: false, studio: false };
    state.xrefErrors[source] = failed;
    var now = state.xrefErrors.screenspace || state.xrefErrors.studio;
    if (now && !prev && !state.xrefErrorToastShown) {
      state.xrefErrorToastShown = true;
      showToast("Cross-references unavailable. Is Screenspace/Studio running?");
    }
    if (!now) state.xrefErrorToastShown = false;
    updateStatusIndicator();
  }

  // With no spreadsheet open, /studio/api/sheet answers {ok, sheet_loaded: false}
  // with no rows — nothing this page can index — so stop asking after the first
  // such answer. Cleared on tab focus (see startXrefPolling's resume): a sheet
  // opened from another tab reloads only that document, not this one. The
  // clipgenApplyConfig call below is a refresh only; loadParticipants applies
  // the same config from this page's own api/participants at boot.
  var _sheetXrefIdle = false;

  function loadCrossRefData() {
    fetch("../screenspace/api/events?excluded=false")
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      })
      .then(function (data) {
        _markXrefSource("screenspace", false);
        if (data.ok) {
          state.ssEvents = data.events || [];
          state.ssEventsLoaded = true;
          _buildEventsIndex();
        }
      })
      .catch(function () { _markXrefSource("screenspace", true); });

    if (_sheetXrefIdle) return;
    fetch("../studio/api/sheet")
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      })
      .then(function (data) {
        _markXrefSource("studio", false);
        if (data.ok && data.sheet_loaded === false) _sheetXrefIdle = true;
        if (data.ok) {
          clipgenApplyConfig(data.config);
          state.sheetRows = data.rows || [];
          state.sheetParticipants = data.participants || [];
          state.sheetLoaded = true;
          _buildSheetIndex();
        }
      })
      .catch(function () { _markXrefSource("studio", true); });
  }

  function parseSheetTimestamps(raw) {
    // Cross-reference search; baselines are not applied here so timestamps
    // are interpreted in their relative form (MM:SS for 2-part).
    var segs = parseClipSegmentsForCell(raw, 0, CLIPGEN_CONFIG.defaultDuration);
    return segs.map(function (s) {
      return { start: s.startSeconds, duration: s.duration };
    });
  }

  // Per-participant indexes keyed on start time. Rebuilt when loadCrossRefData
  // receives fresh data, so findOverlapsForSearch can binary-search instead of
  // linearly scanning ssEvents / sheetRows (and re-parsing sheet timestamps)
  // on every rendered segment.

  function _buildEventsIndex() {
    var byP = {};
    for (var i = 0; i < state.ssEvents.length; i++) {
      var ev = state.ssEvents[i];
      var p = ev.participant;
      if (!byP[p]) byP[p] = [];
      byP[p].push({ in: ev.time_in, out: ev.time_out, ev: ev });
    }
    for (var k in byP) {
      byP[k].sort(function (a, b) { return a.in - b.in; });
    }
    state.xrefIndex.eventsByParticipant = byP;
  }

  function _buildSheetIndex() {
    var byP = {};
    for (var j = 0; j < state.sheetRows.length; j++) {
      var row = state.sheetRows[j];
      if (!row.cells) continue;
      for (var p in row.cells) {
        var cell = row.cells[p];
        if (!cell || !cell.valid) continue;
        var segs = parseSheetTimestamps(cell.value);
        if (segs.length === 0) continue;
        if (!byP[p]) byP[p] = [];
        for (var k = 0; k < segs.length; k++) {
          byP[p].push({
            start: segs[k].start,
            end: segs[k].start + segs[k].duration,
            observation: row.observation,
            category: row.category,
            rowIdx: j,
          });
        }
      }
    }
    for (var pp in byP) {
      byP[pp].sort(function (a, b) { return a.start - b.start; });
    }
    state.xrefIndex.sheetByParticipant = byP;
  }

  // Binary search: return the first index where arr[i].key >= value, using the
  // provided key function.
  function _lowerBound(arr, value, keyFn) {
    var lo = 0;
    var hi = arr.length;
    while (lo < hi) {
      var mid = (lo + hi) >>> 1;
      if (keyFn(arr[mid]) < value) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  function findOverlapsForSearch(participant, start, end) {
    var result = { screenspaceEvents: [], sheetObservations: [] };

    var events = state.xrefIndex.eventsByParticipant[participant];
    if (events && events.length > 0) {
      // Candidates are entries where `in < end`. Since events are sorted by
      // `in`, lower-bound on `end` gives us the exclusive upper cursor.
      var upper = _lowerBound(events, end, function (e) { return e.in; });
      for (var i = 0; i < upper; i++) {
        if (events[i].out > start) result.screenspaceEvents.push(events[i].ev);
      }
    }

    var segs = state.xrefIndex.sheetByParticipant[participant];
    if (segs && segs.length > 0) {
      var upper2 = _lowerBound(segs, end, function (s) { return s.start; });
      // Preserve the original "first observation per row" behavior by de-duping
      // on rowIdx. Row-iteration order matched sheet row order before; replicate
      // that by collecting matches and sorting by rowIdx.
      var seenRow = {};
      var matches = [];
      for (var m = 0; m < upper2; m++) {
        if (segs[m].end <= start) continue;
        var rid = segs[m].rowIdx;
        if (seenRow[rid]) continue;
        seenRow[rid] = true;
        matches.push(segs[m]);
      }
      matches.sort(function (a, b) { return a.rowIdx - b.rowIdx; });
      for (var n = 0; n < matches.length; n++) {
        result.sheetObservations.push({
          observation: matches[n].observation,
          category: matches[n].category,
        });
      }
    }

    return result;
  }

  // ---- Participants ----

  function needsTranscription() {
    for (var i = 0; i < state.participants.length; i++) {
      var p = state.participants[i];
      if (p.has_video && !p.has_transcript) return true;
    }
    return false;
  }

  function applyTranscriptionModelHint(data) {
    if (!data || !data.ok) return;
    state.modelStatus = data;
    // Seed the change-detector once; thereafter only settings saves update it,
    // so a poll landing mid-save can't mask a model change.
    if (data.model && _lastTranscribeModel === null) _lastTranscribeModel = data.model;
    // Track how long we've been in an apparent "failed to load" state, so
    // the indicator doesn't flash red on cold load before the first response.
    var looksFailed = !data.loaded && !data.warming && data.prewarm !== "off";
    if (looksFailed) {
      if (state.modelFailSince === 0) state.modelFailSince = Date.now();
    } else {
      state.modelFailSince = 0;
    }
    updateStatusIndicator();
  }

  // ---- Status indicator ----

  function _taskForSelectedParticipant() {
    var pid = state.selectedParticipant;
    if (!pid) return null;
    var latest = null;
    for (var i = 0; i < state.tasks.length; i++) {
      var t = state.tasks[i];
      if (t.participant !== pid) continue;
      // Priority: running > queued > failed > completed/cancelled > stale
      if (!latest) { latest = t; continue; }
      var order = { running: 5, queued: 4, failed: 3, completed: 2, cancelled: 1 };
      if ((order[t.status] || 0) > (order[latest.status] || 0)) latest = t;
    }
    return latest;
  }

  // A cancel DELETE is in flight for this task. Every surface that words a
  // task's progress asks this; without a shared predicate each one re-derives
  // "running means working" and they drift apart.
  function _cancelPending(task) {
    return !!(task && state.cancellingTasks[task.id]);
  }

  function _selectedParticipantRow() {
    var pid = state.selectedParticipant;
    if (!pid) return null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) return state.participants[i];
    }
    return null;
  }

  function _modelLine() {
    var ms = state.modelStatus;
    if (!ms) return "Model: loading\u2026";
    if (ms.loaded) return "Model: ready" + (ms.model ? " \u00B7 " + ms.model : "");
    if (ms.warming) return "Model: loading\u2026";
    if (ms.prewarm === "off") return "Model: idle (loads on demand)";
    if (state.modelFailSince && Date.now() - state.modelFailSince >= MODEL_FAIL_GRACE_MS) {
      return "Model: failed to load";
    }
    return "Model: not loaded";
  }

  function _modelHasFailed() {
    var ms = state.modelStatus;
    if (!ms) return false;
    if (ms.loaded || ms.warming) return false;
    if (ms.prewarm === "off") return false;
    return state.modelFailSince > 0 && Date.now() - state.modelFailSince >= MODEL_FAIL_GRACE_MS;
  }

  function computeIndicatorState() {
    var pid = state.selectedParticipant;
    var task = _taskForSelectedParticipant();
    var row = _selectedParticipantRow();
    var cls = "status-indicator--ready";
    var taskLine;

    if (!pid) {
      taskLine = "No participant selected";
    } else if (_cancelPending(task)) {
      // Stays --working, not --ready: the worker really is still winding down,
      // and a ready dot beside a pill that says "Cancelling…" is a second
      // contradiction on top of the one this whole change removes.
      cls = "status-indicator--working";
      taskLine = pid + ": cancelling…";
    } else if (task && task.status === "running" && task.phase === "loading_model") {
      cls = "status-indicator--working";
      taskLine = pid + ": loading transcription model\u2026";
    } else if (task && task.status === "running") {
      cls = "status-indicator--working";
      var pct = Math.round((task.progress || 0) * 100);
      taskLine = pid + ": transcribing\u2026 " + pct + "%" + _txEtaSuffix(pid, task);
    } else if (task && task.status === "queued") {
      cls = "status-indicator--working";
      taskLine = pid + ": queued";
    } else if (task && task.status === "failed") {
      cls = "status-indicator--error";
      taskLine = pid + ": transcription failed" + (task.error ? " (" + task.error + ")" : "");
    } else if (row && row.has_transcript) {
      taskLine = pid + ": " + (row.segment_count || 0) + " segments";
      if (row.has_stale_artifacts) taskLine += " \u00B7 artifacts outdated";
    } else if (row && row.has_video) {
      taskLine = pid + ": not transcribed";
    } else {
      taskLine = pid + ": no source video";
    }

    // Model failure overrides non-task-active states (keep working class if a
    // task is currently running — the user can see the task is progressing).
    if (cls !== "status-indicator--working" && _modelHasFailed()) {
      cls = "status-indicator--error";
    }

    var lines = [_modelLine(), taskLine];
    if (state.xrefErrors && (state.xrefErrors.screenspace || state.xrefErrors.studio)) {
      var down = [];
      if (state.xrefErrors.screenspace) down.push("Screenspace");
      if (state.xrefErrors.studio) down.push("Studio");
      lines.push("Cross-references unavailable (" + down.join(", ") + ")");
    }
    var ariaLabel;
    if (cls === "status-indicator--working") ariaLabel = taskLine;
    else if (cls === "status-indicator--error") ariaLabel = lines.join(" \u2014 ");
    else ariaLabel = "Ready \u2014 " + taskLine;

    return { cls: cls, ariaLabel: ariaLabel, lines: lines };
  }

  function updateStatusIndicator() {
    var indicator = qs("#trStatusIndicator");
    if (!indicator) return;
    var s = computeIndicatorState();
    indicator.classList.remove(
      "status-indicator--ready",
      "status-indicator--error",
      "status-indicator--working"
    );
    indicator.classList.add(s.cls);
    indicator.setAttribute("aria-label", s.ariaLabel);
    var sr = qs("#trStatusSr");
    if (sr) sr.textContent = s.lines.join(" \u2014 ");
  }

  function initStatusIndicatorTooltip() {
    var indicator = qs("#trStatusIndicator");
    if (!indicator) return;
    attachHoverTooltip(indicator, function () {
      return computeIndicatorState().lines.join("\n");
    }, { multiline: true, align: "center" });
    updateStatusIndicator();
  }

  // Repaint the three hub-owned surfaces that word a task's progress — the
  // status dot, the empty-pane copy and the streaming footer — without waiting
  // for the next poll. Only the pills satellite needs this: it is the one place
  // that changes task wording outside the poll loop. Deliberately NOT folded
  // into _txEtaTicker's tick body, which is a per-second hot path.
  function refreshTranscribeWording() {
    updateStatusIndicator();
    var task = _taskForSelectedParticipant();
    _setTranscriptEmptyText(task);
    var txt = document.querySelector("#segmentList .streaming-text");
    if (txt && task) txt.textContent = _streamingTextStr(task.progress || 0);
  }

  function refreshTranscriptionModelHintOnce() {
    apiGet("api/transcribe/model-status")
      .then(function (data) {
        applyTranscriptionModelHint(data);
      })
      .catch(function () {});
  }

  function stopModelHintPoll() {
    if (_modelHintPoller) {
      _modelHintPoller.stop();
      _modelHintPoller = null;
    }
  }

  // Forget which Whisper downloads the user has agreed to and refresh the
  // models cache. Call when a download attempt has concluded (transcription
  // finished, or a warmup load loaded/failed) so the next gate re-reads real
  // cache state — re-prompting for a model whose download failed, staying
  // quiet for one that succeeded. Never call mid-download (it would re-prompt).
  function _forgetWhisperDownloadAgreements() {
    _whisperDownloadConfirmed = {};
    _trModelsCache = null;
    _trModelsCachePromise = null;
  }

  function startModelHintPoll() {
    stopModelHintPoll();
    var ticks = 0;
    var poll = function () {
      ticks++;
      if (ticks > 120) {
        stopModelHintPoll();
        return;
      }
      apiGet("api/transcribe/model-status")
        .then(function (data) {
          if (!data.ok) return;
          applyTranscriptionModelHint(data);
          if (data.loaded) {
            stopModelHintPoll();
            _forgetWhisperDownloadAgreements();
            return;
          }
          if (!data.warming) {
            // Warmup ended without loading (failed/idle) — let the next attempt
            // re-confirm rather than silently re-downloading.
            stopModelHintPoll();
            _forgetWhisperDownloadAgreements();
          }
        })
        .catch(function () {});
    };
    // createPoller runs poll() once immediately (runImmediately default), then
    // every 1.5s.
    _modelHintPoller = createPoller(poll, 1500, { label: "transcripts.modelHint" });
    _modelHintPoller.start();
  }

  function maybeWarmOnPillHover(p, s) {
    if (state.transcribePrewarm !== "queue_open") return;
    if (!s || s.status === "completed") return;
    tryPostTranscriptionWarmup();
  }

  // Ask the backend to preload the Whisper model. Idempotent per page load via
  // `_transcriptionWarmupPosted`, whose flag resets whenever the post didn't lead
  // to a load (skipped / error / no-op) so a later trigger can retry.
  //   skipped           — backend declined; flag stays down to allow a re-prompt.
  //   already_loaded    — in memory, nothing to poll.
  //   started / warming — run the model-hint poller until it finishes.
  //   !ok / catch       — transient; clear the flag for retry.
  function tryPostTranscriptionWarmup() {
    if (_transcriptionWarmupPosted) return;
    if (state.transcribePrewarm === "off") return;
    if (!needsTranscription()) return;
    _transcriptionWarmupPosted = true;
    apiPost("api/transcribe/warmup", {})
      .then(function (data) {
        if (!data.ok) {
          _transcriptionWarmupPosted = false;
          return;
        }
        if (data.skipped) {
          if (data.reason === "model_not_cached") {
            _confirmPrewarmDownload(data);
            return;
          }
          _transcriptionWarmupPosted = false;
          refreshTranscriptionModelHintOnce();
          return;
        }
        if (data.already_loaded) {
          refreshTranscriptionModelHintOnce();
          return;
        }
        if (data.started || data.already_warming) {
          startModelHintPoll();
          return;
        }
        _transcriptionWarmupPosted = false;
      })
      .catch(function () {
        _transcriptionWarmupPosted = false;
      });
  }

  // Prewarm wanted to load a model that isn't downloaded yet. We never
  // download silently — confirm with the user first (once per session). On
  // confirm, re-post warmup with force=true so the backend proceeds; on
  // decline, leave the warmup flag set so we stop re-posting/re-prompting
  // (the model still loads, with confirmation, at transcribe time).
  function _confirmPrewarmDownload(data) {
    if (_prewarmDownloadPrompting) return;
    if (_prewarmDeclinedModel === data.model) {
      _transcriptionWarmupPosted = true;
      refreshTranscriptionModelHintOnce();
      return;
    }
    _prewarmDownloadPrompting = true;
    confirmModelInstall({
      kind: "whisper",
      prewarm: true,
      model: data.model,
      sizeMb: data.size_mb,
    }).then(function (ok) {
      _prewarmDownloadPrompting = false;
      if (!ok) {
        _prewarmDeclinedModel = data.model;
        _transcriptionWarmupPosted = true;
        refreshTranscriptionModelHintOnce();
        return;
      }
      // Agreed once — don't re-prompt at transcribe time before the download
      // completes (the models cache still reports it as not cached).
      if (data.model) _whisperDownloadConfirmed[data.model] = true;
      apiPost("api/transcribe/warmup", { force: true })
        .then(function (d) {
          if (!d.ok) {
            _transcriptionWarmupPosted = false;
            return;
          }
          if (d.started || d.already_warming) {
            startModelHintPoll();
            return;
          }
          if (d.already_loaded) {
            refreshTranscriptionModelHintOnce();
            return;
          }
          _transcriptionWarmupPosted = false;
        })
        .catch(function () {
          _transcriptionWarmupPosted = false;
        });
    });
  }

  // One-shot #P07 deep-link guard: the hash must not re-hijack selection on
  // later participant-list refreshes after the user has moved on.
  var _hashPidApplied = false;

  // The pill row and the transcript pane both ship placeholders (skeleton pills,
  // a hidden #transcriptEmpty) that only the first successful render replaces.
  // If that render never happens the page shimmers forever and the transcript
  // pane stays blank, which reads as "still loading" rather than "unreachable" —
  // so a failed *boot* fetch has to fall back to the real empty states. A later
  // refresh failing is harmless: the already-rendered list stays put.
  function _clearBootPlaceholders() {
    if (state.participants.length) return;
    renderPills();
    renderEmptyState();
  }

  function loadParticipants() {
    return apiGet("api/participants").then(function (data) {
      if (!data.ok) {
        _clearBootPlaceholders();
        return;
      }
      // Primary config channel for this page: the only other caller of
      // clipgenApplyConfig here is the xref poller's /studio/api/sheet leg,
      // which is status-gated and lands (if ever) after first render.
      if (data.config) clipgenApplyConfig(data.config);
      state.participants = data.participants;
      state.hasSheet = !!data.has_sheet;
      state.transcribePrewarm = data.transcribe_prewarm || "queue_open";
      renderPills();
      refreshTopNavActions();

      if (!needsTranscription()) {
        _transcriptionWarmupPosted = false;
      } else if (state.transcribePrewarm === "page_load") {
        tryPostTranscriptionWarmup();
      }
      refreshTranscriptionModelHintOnce();

      // Precedence: deep link (#P07, from the Overview Map's explain panel;
      // wins once, on the first load that has the participant list) > current
      // in-memory selection (soft refresh) > localStorage (fresh page load).
      var hashPid = _hashPidApplied ? "" : clipgenHashParticipant();
      if (hashPid && state.participants.length) _hashPidApplied = true;
      var pick = clipgenPickParticipant(state.participants, {
        hashPid: hashPid,
        currentId: state.selectedParticipant,
        storedId: getStoredUIState("transcripts").selectedParticipant,
      });
      if (pick) {
        if (pick !== state.selectedParticipant) selectParticipant(pick);
        return;
      }

      // Auto-select first participant with a transcript, or just the first
      var first = null;
      for (var k = 0; k < state.participants.length; k++) {
        if (state.participants[k].has_transcript) { first = state.participants[k]; break; }
      }
      if (!first && state.participants.length > 0) first = state.participants[0];
      if (first) selectParticipant(first.id);
      else renderEmptyState();
    }).catch(function () {
      _clearBootPlaceholders();
    });
  }


  // Move the selection to the previous/next participant in the sidebar order,
  // wrapping around. Bound to Z / X (see transcripts-video.js).
  function cycleParticipant(delta) {
    var list = state.participants;
    if (!list || list.length < 2) return;
    var idx = -1;
    for (var i = 0; i < list.length; i++) {
      if (list[i].id === state.selectedParticipant) { idx = i; break; }
    }
    if (idx === -1) idx = 0;
    var next = list[(idx + delta + list.length) % list.length];
    if (next && next.id !== state.selectedParticipant) selectParticipant(next.id);
  }

  function selectParticipant(pid) {
    state.participantReqVer++;
    _stopSummaryPoll();
    _stopCitationsPoll();
    _stopFrictionPoll();
    hideMarkPopover();
    state.selectedParticipant = pid;
    setStoredUIStateField("transcripts", "selectedParticipant", pid);
    // Restore this participant's transcribe-range markers (video satellite;
    // late-bound — loadedmetadata clamps them once the duration is known).
    if (TS.restoreMarkers) {
      TS.restoreMarkers(pid);
      TS.updateMarkerInfo();
      renderTimeline();
    }
    renderPills();
    refreshTopNavActions();

    // Find participant info
    var p = null;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) { p = state.participants[i]; break; }
    }
    if (!p) return;

    // Fragmented-MP4 warning + remux action, above the player.
    if (window.clipgenMediaBanner) {
      window.clipgenMediaBanner.show(qs("#videoSection"), p);
    }

    // Clean up previous video state
    var video = qs("#videoPlayer");
    var videoEmpty = qs("#videoEmpty");
    video.pause();
    cancelPendingSeek();

    // ?v=<mtime_ns> mirrors the screenspace cache-bust, so a replaced source file
    // invalidates the browser cache rather than relying on Last-Modified
    // revalidation. The audio-layout reset sits outside the has_video branch on
    // purpose: selecting a participant *without* video must also drop the mix, or
    // orphaned <audio> elements keep playing, the <video> stays force-muted, and
    // the volume popover shows the previous participant's sliders.
    state.audioTracks = [];
    if (state.audioPanel) state.audioPanel.refresh();

    if (p.has_video) {
      // Lazily probe the audio-track layout for the volume popover's caption +
      // per-track mixer, then reconfigure once the layout arrives. Shares the
      // pill picker's cache (_trFetchAudioInfo). Guarded by participantReqVer so
      // a stale response can't overwrite the current tracks.
      (function (ver) {
        _trFetchAudioInfo(pid, p.video_version)
          .then(function (info) {
            if (ver !== state.participantReqVer) return;
            if (info) state.audioTracks = info.tracks;
            if (state.audioPanel) state.audioPanel.refresh();
          })
          .catch(function () {});
      })(state.participantReqVer);

      // Multi-video participants carry a per-part timeline; play part-by-part
      // with client-side source switching (see the timeline helpers).
      state.videoTimeline = p.timeline && p.timeline.length > 1 ? p.timeline : null;
      state.videoVersion = p.video_version != null ? p.video_version : null;
      state.videoActivePart = 0;
      state.videoOffset = 0;
      video.classList.remove("hidden");
      videoEmpty.classList.add("hidden");

      // Set VTT track. Browsers reset textTracks[0].mode when the track src
      // changes, so re-apply the user's preference now and again on the
      // track's load event in initVideoPlayer. The native overlay track carries
      // GLOBAL cue times, which only align with a single continuous file — for
      // multi-video participants it is disabled (the in-app transcript list still
      // highlights the active segment via global time).
      var track = qs("#subtitleTrack");
      if (state.videoTimeline) {
        track.removeAttribute("src");
      } else {
        track.src = "api/vtt/" + pid;
      }
      applyCaptionMode();

      // Restore the saved playback offset (GLOBAL) for this participant if we
      // have one; otherwise seek to 0.001s so the first frame renders without
      // waiting for play/scrub. `preload="metadata"` decodes duration but not
      // pixels, so without this nudge the viewer shows a blank gray box on load.
      var storedMap = getStoredUIState("transcripts").videoTimeByParticipant;
      var savedTime =
        storedMap && typeof storedMap[pid] === "number" ? storedMap[pid] : 0.001;
      if (state.videoTimeline) {
        var pi = _partForGlobal(state.videoTimeline, savedTime);
        state.videoActivePart = pi;
        state.videoOffset = state.videoTimeline[pi].cumulativeStart;
        var localStart = savedTime - state.videoOffset;
        video.src = _partMediaUrl(pi);
      } else {
        var mediaUrl = "media/" + p.video_filename;
        if (p.video_version != null) {
          mediaUrl += "?v=" + encodeURIComponent(p.video_version);
        }
        video.src = mediaUrl;
      }
      var restoreTime = function () {
        video.removeEventListener("loadedmetadata", restoreTime);
        if (state.selectedParticipant !== pid) return;
        video.currentTime = state.videoTimeline ? localStart : savedTime;
      };
      video.addEventListener("loadedmetadata", restoreTime);
    } else {
      state.videoTimeline = null;
      state.videoOffset = 0;
      state.videoActivePart = 0;
      video.removeAttribute("src");
      video.load();
      video.classList.add("hidden");
      videoEmpty.classList.remove("hidden");
    }

    var taskForPid = null;
    state.tasks.forEach(function (t) {
      if (t.participant === pid && (t.status === "running" || t.status === "queued")) {
        taskForPid = t;
      }
    });

    updateStatusIndicator();

    // Load transcript
    if (p.has_transcript) {
      state.streamingParticipant = null;
      _setAnalysisPanelVisible(true);
      _restoreActiveTab(pid);
      loadTranscript(pid);
      loadSummary(pid);
      loadFriction(pid);
    } else if (taskForPid && taskForPid.status === "running" && taskForPid.partial_count > 0) {
      state.streamingParticipant = pid;
      clearAnalysisPanel();
      _syncStreamSegs(taskForPid, function (segs) {
        if (segs.length > 0) renderPartialSegments(segs, taskForPid.progress);
      });
    } else {
      state.segments = [];
      state.streamingParticipant = null;
      renderSegments();
      // _taskForSelectedParticipant, not taskForPid: the pane must agree with
      // the indicator's running-over-queued pick when duplicate tasks exist.
      _setTranscriptEmptyText(_taskForSelectedParticipant());
      renderTimeline();
      clearAnalysisPanel();
    }

    // Reflect the newly-selected participant's transcription progress on the
    // timeline immediately (draws the band if it's mid-transcription, clears it
    // otherwise) rather than waiting up to a full poll interval.
    updateTranscribeFill();
  }

  // Contextual text for the #transcriptEmpty pane: while the newest task for
  // the selected participant is queued / loading the model / decoding with no
  // segments streamed yet, "No transcript available" is misleading — say what
  // the wait actually is. Called with null to restore the defaults.
  function _setTranscriptEmptyText(task) {
    var empty = qs("#transcriptEmpty");
    if (!empty) return;
    var main = empty.querySelector("p");
    var hint = empty.querySelector(".empty-hint");
    // The three waiting lines shimmer; "No transcript available" is a resting
    // state, not an in-flight one, so it stays flat.
    var waiting = !!(task && (task.status === "running" || task.status === "queued"));
    main.classList.toggle("cg-shimmer", waiting);
    if (_cancelPending(task)) {
      main.textContent = "Cancelling…";
      // Deliberately not "finishing the current segment": a cancel during
      // loading_model is waiting on the model load, not on a segment.
      hint.textContent = "Waiting for the transcription worker to stop";
    } else if (task && task.status === "running" && task.phase === "loading_model") {
      main.textContent = "Loading transcription model…";
      hint.textContent = "The first transcription after a restart takes a few extra seconds";
    } else if (task && task.status === "running") {
      main.textContent = "Starting transcription…";
      hint.textContent = "Lines appear here as they are transcribed";
    } else if (task && task.status === "queued") {
      main.textContent = "Queued for transcription…";
      hint.textContent = "Waiting for the current task to finish";
    } else {
      main.textContent = "No transcript available";
      hint.textContent = "Use the Queue panel to transcribe this participant's video";
    }
  }

  function renderEmptyState() {
    qs("#videoPlayer").classList.add("hidden");
    qs("#videoEmpty").classList.remove("hidden");
    qs("#segmentList").innerHTML = "";
    _setTranscriptEmptyText(null);
    qs("#transcriptEmpty").classList.remove("hidden");
    clearAnalysisPanel();
    clearTimelineMarkers();
    updateTranscribeFill(); // clear any lingering transcribe band
    renderTimeline();
  }

  // ---- Transcript loading ----

  function loadTranscript(pid) {
    var ver = state.participantReqVer;
    return apiGet("api/transcript/" + pid).then(function (data) {
      if (ver !== state.participantReqVer) return;
      if (!data.ok) {
        state.segments = [];
        renderSegments();
        renderTimeline();
        return;
      }
      state.segments = data.segments;
      state.activeSegmentIndex = -1;
      renderSegments();
      renderTimeline();
    });
  }

  // ---- Analysis panel: summary + citations + friction (impl in transcripts-agents.js) ----
  // Thin hub delegators forward to the satellite. selectParticipant + the poller +
  // the visibility/focus handlers drive load*/clear*/stop*/_restoreActiveTab; boot
  // wires the init*; the segment-list hover calls _show/_hideFrictionTooltip; the
  // poller's re-arm asks _currentParticipant/_frictionDepMet.
  function loadSummary() { return TS.loadSummary && TS.loadSummary.apply(null, arguments); }
  function loadFriction() { return TS.loadFriction && TS.loadFriction.apply(null, arguments); }
  function clearAnalysisPanel() { return TS.clearAnalysisPanel && TS.clearAnalysisPanel(); }
  function _setAnalysisPanelVisible() { return TS._setAnalysisPanelVisible && TS._setAnalysisPanelVisible.apply(null, arguments); }
  function _restoreActiveTab() { return TS._restoreActiveTab && TS._restoreActiveTab.apply(null, arguments); }
  function initPanelTabs() { return TS.initPanelTabs && TS.initPanelTabs(); }
  function initSummaryActions() { return TS.initSummaryActions && TS.initSummaryActions(); }
  function initFriction() { return TS.initFriction && TS.initFriction(); }
  function initFrictionMode() { return TS.initFrictionMode && TS.initFrictionMode(); }
  function applyFrictionDecorations() { return TS.applyFrictionDecorations && TS.applyFrictionDecorations(); }
  function cycleFrictionMode() { return TS.cycleFrictionMode && TS.cycleFrictionMode(); }
  function _stopSummaryPoll() { return TS._stopSummaryPoll && TS._stopSummaryPoll(); }
  function _stopCitationsPoll() { return TS._stopCitationsPoll && TS._stopCitationsPoll(); }
  function _stopFrictionPoll() { return TS._stopFrictionPoll && TS._stopFrictionPoll(); }
  function _currentParticipant() { return TS._currentParticipant && TS._currentParticipant(); }
  function _frictionDepMet() { return TS._frictionDepMet && TS._frictionDepMet(); }
  function _showFrictionTooltip() { return TS._showFrictionTooltip && TS._showFrictionTooltip.apply(null, arguments); }
  function _hideFrictionTooltip() { return TS._hideFrictionTooltip && TS._hideFrictionTooltip(); }

  // ---- Segment rendering ----

  // Spanned to close PERFORMANCE-PLAN-3 §8c, which gates virtualizing this list
  // on "profiling shows >2000-segment sessions hurting". Same shape as
  // studio.renderGrid; the CSS content-visibility and scroll-preservation
  // halves already landed.
  function renderSegments() {
    return clipgenPerf.span("transcripts.renderSegments", renderSegmentsImpl);
  }

  function renderSegmentsImpl() {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    state.editingTextEl = null;
    state.cachedSegmentRows = null;
    // We're rendering the finalized transcript — drop any queued streaming
    // indicator so a paused-tab RAF can't re-insert it over the real segments.
    _cancelStreamingIndicator();

    // Scroll lives on #trMain, not on #segmentList (the floating nav scrolls
    // under it) — same probe renderPartialSegments uses. A full rebuild of a
    // same-participant list (heatmap toggle, tooltip toggle, streaming→final
    // swap) must not drop the reader to the top; a participant change must.
    var scrollHost = qs("#trMain") || container;
    var samePid = _renderedSegmentsPid === state.selectedParticipant;
    var restoreTop = samePid ? scrollHost.scrollTop : 0;
    _renderedSegmentsPid = state.selectedParticipant;

    if (state.segments.length === 0) {
      container.innerHTML = "";
      empty.classList.remove("hidden");
      return;
    }
    empty.classList.add("hidden");

    var html = "";
    for (var i = 0; i < state.segments.length; i++) {
      var seg = state.segments[i];
      var activeClass = i === state.activeSegmentIndex ? " active" : "";
      var correctedClass = seg.corrected ? " segment-corrected" : "";
      var markObj = seg.marks && seg.marks.length > 0 ? seg.marks[0] : null;
      var markCat = markObj ? (MARK_CATEGORIES[markObj.category] || MARK_CATEGORIES.bookmark) : null;
      var markColor = markCat ? markCat.color : null;
      var markClass = markObj ? "segment-mark marked" : "segment-mark";
      var markStyle = markColor ? ' style="background:' + markColor + '"' : "";
      var markLabel = markObj && markObj.label ? ' data-tooltip="' + escapeHtml(markObj.label) + '"' : "";
      var annoBadgeHtml = "";
      if (markObj && markObj.label && markColor) {
        var bgMix = "color-mix(in oklch, " + markColor + " 18%, transparent)";
        var borderMix = "color-mix(in oklch, " + markColor + " 50%, transparent)";
        var badgeStyle = "--anno-badge-fg:" + markColor + ";--anno-badge-bg:" + bgMix + ";--anno-badge-border:" + borderMix;
        annoBadgeHtml = '<span class="segment-anno-badge" style="' + badgeStyle + '">' + escapeHtml(markObj.label) + '</span>';
      }
      var sevDotHtml = "";
      if (markObj && markObj.severity) {
        sevDotHtml = '<span class="segment-sev-dot ' + severityClass(markObj.severity) + '" data-tooltip="' + escapeHtml(markObj.severity) + '"></span>';
      }

      // No friction markup here on purpose: applyFrictionDecorations() below owns
      // every friction class/inline var, so the rebuild path and the live filter
      // path can't drift apart.
      html += '<div class="segment-row' + activeClass + correctedClass + '" data-index="' + i + '" data-start="' + seg.start + '">';
      html += '<span class="' + markClass + '" data-segment-id="' + escapeHtml(seg.id) + '"' + markStyle + markLabel + '></span>';
      html += sevDotHtml;
      html += '<span class="segment-timestamp">' + formatTime(seg.start);
      // Cross-reference badges in gutter (inside timestamp, positioned at right edge)
      if (CLIPGEN_CONFIG.crossReferences) {
        var xref = findOverlapsForSearch(state.selectedParticipant, seg.start, seg.end);
        if (xref.screenspaceEvents.length > 0 || xref.sheetObservations.length > 0) {
          html += '<span class="segment-xref-badges">';
          if (xref.screenspaceEvents.length > 0) {
            var evTypes = [];
            var evSeen = {};
            for (var ei = 0; ei < xref.screenspaceEvents.length; ei++) {
              var et = xref.screenspaceEvents[ei].event_type || xref.screenspaceEvents[ei].detector;
              if (!evSeen[et]) { evSeen[et] = true; evTypes.push(et); }
            }
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.screenspace.color + '" data-tooltip="' + escapeHtml(evTypes.join(", ")) + '"><span class="xref-badge-icon" style="' + iconMaskStyle(XREF_BADGES.screenspace.icon) + '"></span></span>';
          }
          if (xref.sheetObservations.length > 0) {
            var obsTitle = xref.sheetObservations[0].observation;
            html += '<span class="segment-xref-badge" style="background:' + XREF_BADGES.sheet.color + '" data-tooltip="' + escapeHtml(obsTitle) + '"><span class="xref-badge-icon" style="' + iconMaskStyle(XREF_BADGES.sheet.icon) + '"></span></span>';
          }
          html += '</span>';
        }
      }
      html += '</span>';
      // Split text into word spans. When per-word timing exists and the row is
      // uncorrected, the Nth token gets the Nth word's data-ws/data-we for the
      // karaoke sweep (transcripts-video.js). Display text stays exactly
      // seg.text; a token/word count mismatch (corrections, tokenization drift)
      // just leaves the spans untimed and the row degrades to row-level highlight.
      var tokens = seg.text.split(/(\s+)/);
      var wordCount = 0;
      for (var w = 0; w < tokens.length; w++) {
        if (tokens[w] && !/^\s+$/.test(tokens[w])) wordCount++;
      }
      var segWords = (!seg.corrected && seg.words && seg.words.length === wordCount) ? seg.words : null;
      var wordHtml = "";
      var wi = 0;
      for (var w = 0; w < tokens.length; w++) {
        if (/^\s+$/.test(tokens[w])) {
          wordHtml += tokens[w];
        } else if (tokens[w]) {
          var timing = segWords ? ' data-ws="' + segWords[wi].start + '" data-we="' + segWords[wi].end + '"' : "";
          wordHtml += '<span class="segment-word"' + timing + '>' + escapeHtml(tokens[w]) + '</span>';
          wi++;
        }
      }
      html += '<span class="segment-text" data-id="' + escapeHtml(seg.id) + '">' + annoBadgeHtml + wordHtml + '</span>';
      html += '<span class="segment-copy" data-tooltip="Copy text"><span class="segment-copy-icon"></span></span>';
      html += '</div>';
    }
    container.innerHTML = html;
    // Before the scroll restore, not after: isolate mode hides rows and the
    // callouts add height, so decorating afterwards would let the browser clamp
    // the restored offset against a stale scrollHeight.
    applyFrictionDecorations();
    // Same task as the wipe, so the new scrollHeight is already laid out and the
    // browser can't clamp us to 0 — and initPipScroll's rAF-coalesced listener
    // reads the restored value rather than the transient top. The write itself
    // looks like a reader scroll to the auto-follow pause, hence the marker.
    // Written unconditionally: on a participant switch restoreTop is 0, and
    // without the write the browser keeps the outgoing transcript's offset and
    // drops the reader into the middle of the new one.
    ignoreNextScroll();
    scrollHost.scrollTop = restoreTop;

    _ensureSegmentListDelegation();
    _partialRender.count = 0;
    _partialRender.pid = null;
    _partialRender.segments = null;
    _partialRender.marksVersion = _streamingMarksVersion;
  }

  // Which participant #segmentList currently shows, so a rebuild can tell a
  // same-transcript re-render (restore scroll) from a participant switch (top).
  // renderPartialSegments keeps it current too — the streaming→final swap is
  // exactly when the reader is deepest in the list.
  var _renderedSegmentsPid = null;

  // Append-only state for renderPartialSegments. Each streaming poll appends new
  // trailing segments to #segmentList instead of rebuilding the entire list. A
  // full rebuild is only performed when the participant changes, when the
  // segment count drops (restart), or when the in-memory marks cache changes.

  var _partialRender = {
    pid: null,
    count: 0,
    segments: null,
    marksVersion: 0,
  };

  // Client-side accumulator for streaming partial segments. The status poll only
  // carries partial_count now (not the full array), so we fetch the tail beyond
  // our cursor and append. partial_segments is append-only server-side, so the
  // count is a safe cursor and the request payload stays flat as the transcript
  // grows. Resets when the streamed task changes.
  var _streamSegs = { taskId: null, segments: [], fetching: false };

  function _syncStreamSegs(task, cb) {
    if (_streamSegs.taskId !== task.id) {
      _streamSegs.taskId = task.id;
      _streamSegs.segments = [];
      _streamSegs.fetching = false;
    }
    var count = task.partial_count || 0;
    var have = _streamSegs.segments.length;
    if (count <= have || _streamSegs.fetching) {
      cb(_streamSegs.segments);
      return;
    }
    _streamSegs.fetching = true;
    var reqTaskId = task.id;
    var since = have;
    apiGet(
      "api/transcribe/" + encodeURIComponent(task.id) + "/segments?since=" + since
    ).then(
      function (data) {
        _streamSegs.fetching = false;
        // Reject if the streamed task changed or our cursor moved under us.
        if (!data.ok || _streamSegs.taskId !== reqTaskId) return;
        if (since === _streamSegs.segments.length && data.segments && data.segments.length) {
          _streamSegs.segments = _streamSegs.segments.concat(data.segments);
        }
        cb(_streamSegs.segments);
      },
      function () {
        _streamSegs.fetching = false;
      }
    );
  }

  function _renderPartialSegmentRow(seg, i, pid) {
    var segId = pid + ":" + i;
    var cachedMark = _streamingMarks[segId];
    var cachedColor = cachedMark ? cachedMark.color : null;
    var markClass = "segment-mark" + (cachedColor ? " marked" : "");
    var markStyle = cachedColor ? ' style="background:' + cachedColor + '"' : "";
    var sevDotHtml = "";
    if (cachedMark && cachedMark.severity) {
      sevDotHtml = '<span class="segment-sev-dot ' + severityClass(cachedMark.severity) + '" data-tooltip="' + escapeHtml(cachedMark.severity) + '"></span>';
    }
    var html = '<div class="segment-row segment-streaming" data-index="' + i + '" data-start="' + seg.start + '">';
    html += '<span class="' + markClass + '" data-segment-id="' + escapeHtml(segId) + '"' + markStyle + '></span>';
    html += sevDotHtml;
    html += '<span class="segment-timestamp">' + formatTime(seg.start) + '</span>';
    html += '<span class="segment-text">' + escapeHtml(seg.text) + '</span>';
    html += '<span class="segment-copy" data-tooltip="Copy text"><span class="segment-copy-icon"></span></span>';
    html += '</div>';
    return html;
  }

  // ---- Elapsed / ETA tracking ----
  // Transcription progress is a linear fraction of media duration, so its ETA
  // extrapolation is meaningful (per-participant trackers). The thinking agents
  // expose no progress fraction, so they show elapsed only (single trackers).
  var _txEtaTrackers = {};
  var _summaryEtaTracker = createEtaTracker();
  var _citationsEtaTracker = createEtaTracker();
  var _frictionEtaTracker = createEtaTracker();
  var _txEtaTicker = createIntervalTicker(_tickTxEta, {
    isActive: _anyTxEtaActive,
  });

  // " \u00b7 0:42 \u00b7 ~1:20 left" suffix for a participant's running transcription, or
  // "" when not running (or still loading the model \u2014 no decode progress to
  // time yet). Each entry is keyed by the task's created_at so a re-run of the
  // same participant seeds a fresh tracker from the new task rather than
  // continuing the prior run's elapsed. Seeded from transcribe_started_at so
  // elapsed/ETA exclude the queue wait and the model load. Stale entries are
  // pruned in _tickTxEta.
  function _txEtaSuffix(pid, task) {
    if (!pid || !task || task.status !== "running") return "";
    if (task.phase === "loading_model") return "";
    var entry = _txEtaTrackers[pid];
    if (!entry || entry.createdAt !== task.created_at) {
      var t = createEtaTracker();
      var seedIso = task.transcribe_started_at || task.created_at;
      var seed = seedIso ? Date.parse(seedIso) : NaN;
      t.start(isNaN(seed) ? undefined : seed);
      entry = { tracker: t, createdAt: task.created_at };
      _txEtaTrackers[pid] = entry;
    }
    var e = entry.tracker.update(task.progress);
    var s = " \u00b7 " + formatDuration(e.elapsedSec);
    var eta = formatEtaLabel(e.remainingSec);
    if (eta) s += " \u00b7 " + eta;
    return s;
  }

  function _streamingTextStr(progress) {
    var pid = state.streamingParticipant || state.selectedParticipant;
    var task = _taskForSelectedParticipant();
    // Single choke point for the footer \u2014 the append path, the coalesced RAF
    // update and the per-second ETA ticker all render through here, so one
    // branch keeps all three honest while a cancel is in flight.
    if (_cancelPending(task)) return "Cancelling\u2026";
    return "Transcribing\u2026 " + Math.round(progress * 100) + "%" + _txEtaSuffix(pid, task);
  }

  // Paint a thinking-agent's elapsed clock (elapsed only \u2014 no progress signal).
  function _updateAgentElapsed(spanId, tracker) {
    var sp = document.getElementById(spanId);
    if (!sp) return;
    var e = tracker.update();
    sp.textContent = formatDuration(e.elapsedSec);
  }

  function _anyTxEtaActive() {
    if (_anyAgentActive()) return true;
    for (var i = 0; i < state.tasks.length; i++) {
      if (state.tasks[i].status === "running") return true;
    }
    return false;
  }

  function _tickTxEta() {
    // The ticker's isActive guard (_anyTxEtaActive) self-stops it; this body
    // only runs while transcription or a thinking agent is active.
    // Drop trackers for participants with no running transcription so memory
    // stays bounded and a later re-run starts fresh.
    var runningPids = {};
    for (var r = 0; r < state.tasks.length; r++) {
      if (state.tasks[r].status === "running") runningPids[state.tasks[r].participant] = true;
    }
    Object.keys(_txEtaTrackers).forEach(function (p) {
      if (!runningPids[p]) delete _txEtaTrackers[p];
    });
    // Transcription: status-indicator tooltip + the streaming "Transcribing\u2026" line.
    updateStatusIndicator();
    var txt = document.querySelector("#segmentList .streaming-text");
    if (txt) {
      var task = _taskForSelectedParticipant();
      if (task && task.status === "running") txt.textContent = _streamingTextStr(task.progress || 0);
    }
    // Thinking agents: elapsed-only clocks (spans present only while generating).
    _updateAgentElapsed("summaryElapsed", _summaryEtaTracker);
    _updateAgentElapsed("citationsElapsed", _citationsEtaTracker);
    _updateAgentElapsed("frictionElapsed", _frictionEtaTracker);
  }

  function _streamingIndicatorHtml(progress) {
    return '<div class="streaming-indicator">' +
      '<span class="streaming-dot"></span>' +
      '<span class="streaming-text cg-shimmer">' + _streamingTextStr(progress) + '</span>' +
      '</div>';
  }

  var _streamIndicatorRaf = null;
  var _streamIndicatorPending = null;

  // Cancel a queued streaming-indicator insert. requestAnimationFrame callbacks
  // are paused while the tab is backgrounded, so a RAF scheduled during the last
  // streaming poll can outlive the transcript being finalized and re-insert a
  // stale "Transcribing… X%" row when the user returns. Call this whenever the
  // finalized transcript replaces the streaming view.
  function _cancelStreamingIndicator() {
    if (_streamIndicatorRaf) {
      cancelAnimationFrame(_streamIndicatorRaf);
      _streamIndicatorRaf = null;
    }
    _streamIndicatorPending = null;
  }

  function _updateStreamingIndicator(container, progress) {
    _streamIndicatorPending = { container: container, progress: progress };
    if (_streamIndicatorRaf) return;
    _streamIndicatorRaf = requestAnimationFrame(function () {
      _streamIndicatorRaf = null;
      var pending = _streamIndicatorPending;
      _streamIndicatorPending = null;
      if (!pending || !pending.container || !pending.container.isConnected) return;
      var ind = pending.container.querySelector(".streaming-indicator");
      if (ind) ind.parentNode.removeChild(ind);
      pending.container.insertAdjacentHTML(
        "beforeend",
        _streamingIndicatorHtml(pending.progress)
      );
    });
  }

  // The streaming counterpart: fires per poll tick during transcription, so it
  // is the one that would show up as jank while a long session decodes.
  function renderPartialSegments(segments, progress) {
    return clipgenPerf.span("transcripts.renderPartialSegments", function () {
      return renderPartialSegmentsImpl(segments, progress);
    });
  }

  function renderPartialSegmentsImpl(segments, progress) {
    var container = qs("#segmentList");
    var empty = qs("#transcriptEmpty");
    var pid = state.streamingParticipant || state.selectedParticipant;
    empty.classList.add("hidden");

    // If user is actively editing a segment, skip DOM mutation to preserve edit state
    if (state.editingTextEl && state.editingTextEl.isConnected) return;

    // Row list changes shape on both append and rebuild paths.
    state.cachedSegmentRows = null;

    // Pass 6 floating-nav scroll-under: scroll lives on #trMain, not on
    // #segmentList. Probe the actual scroll container so the
    // auto-follow-streaming-tail behaviour keeps working.
    var trMain = qs("#trMain");
    var scrollHost = trMain || container;
    var nearBottom = scrollHost.scrollHeight - scrollHost.scrollTop - scrollHost.clientHeight < 100;

    var canAppend =
      _partialRender.pid === pid &&
      _partialRender.marksVersion === _streamingMarksVersion &&
      segments.length >= _partialRender.count &&
      container.querySelector(".segment-streaming") !== null;
    var indicatorOnly =
      canAppend && segments.length === _partialRender.count;

    if (canAppend && segments.length > _partialRender.count) {
      // Append only the new trailing rows, then move/refresh the indicator.
      var indicator = container.querySelector(".streaming-indicator");
      if (indicator) indicator.parentNode.removeChild(indicator);
      var tmp = document.createElement("div");
      var fragHtml = "";
      for (var a = _partialRender.count; a < segments.length; a++) {
        fragHtml += _renderPartialSegmentRow(segments[a], a, pid);
      }
      tmp.innerHTML = fragHtml;
      var frag = document.createDocumentFragment();
      while (tmp.firstChild) frag.appendChild(tmp.firstChild);
      container.appendChild(frag);
      container.insertAdjacentHTML("beforeend", _streamingIndicatorHtml(progress));
    } else if (canAppend && segments.length === _partialRender.count) {
      // Same segment count - only refresh the progress indicator (coalesced).
      _updateStreamingIndicator(container, progress);
    } else {
      // Full rebuild: pid change, restart, or marks cache changed.
      var html = "";
      for (var i = 0; i < segments.length; i++) {
        html += _renderPartialSegmentRow(segments[i], i, pid);
      }
      html += _streamingIndicatorHtml(progress);
      container.innerHTML = html;
      _ensureSegmentListDelegation();
    }

    _partialRender.pid = pid;
    _partialRender.count = segments.length;
    _partialRender.segments = segments;
    _partialRender.marksVersion = _streamingMarksVersion;
    // The list now shows this participant, so the finalized render that replaces
    // it counts as a same-participant rebuild and keeps the reader's position.
    _renderedSegmentsPid = pid;

    if (nearBottom) {
      scrollHost.scrollTop = scrollHost.scrollHeight;
    }

    // Mirror partial segments into state.segments so the marker timeline can
    // resolve marker positions during streaming. Each row's id is the
    // composed "<pid>:<index>" string, matching what _streamingMarks is keyed
    // on via _renderPartialSegmentRow above.
    var mirrored = [];
    for (var mi = 0; mi < segments.length; mi++) {
      var s = segments[mi];
      mirrored.push({
        id: pid + ":" + mi,
        start: s.start,
        end: s.end,
        text: s.text,
        marks: [],
      });
    }
    state.segments = mirrored;
    if (!indicatorOnly) {
      renderTimeline();
    }
  }

  // ---- Segment list event delegation ----

  var _segmentListDelegated = false;
  // Coalesces the segment-list mousemove (hot-segment friction tooltip) the same
  // way the timeline canvas does. Lives here with its only user — the delegation
  // handler below — even though the tooltip render itself is in the agents
  // satellite (reached via TS._show/_hideFrictionTooltip).
  var _segTooltipRaf = 0;

  function _ensureSegmentListDelegation() {
    if (_segmentListDelegated) return;
    var container = qs("#segmentList");
    if (!container) return;
    _segmentListDelegated = true;

    container.addEventListener("click", function (e) {
      var row = e.target.closest(".segment-row");
      if (!row) return;
      var isStreaming = row.classList.contains("segment-streaming");
      var idx = parseInt(row.getAttribute("data-index"), 10);
      var start = parseFloat(row.getAttribute("data-start"));

      var markEl = e.target.closest(".segment-mark");
      if (markEl && row.contains(markEl)) {
        e.stopPropagation();
        var segId = markEl.getAttribute("data-segment-id");
        if (isStreaming) {
          var existing = _streamingMarks[segId];
          if (existing) showMarkPopover(markEl, segId, existing);
          else toggleMarkStreaming(segId, markEl);
        } else {
          var seg = state.segments[idx];
          var mark = seg && seg.marks && seg.marks.length > 0 ? seg.marks[0] : null;
          if (mark) showMarkPopover(markEl, segId, mark);
          else toggleMark(segId);
        }
        return;
      }

      var copyEl = e.target.closest(".segment-copy");
      if (copyEl && row.contains(copyEl)) {
        e.stopPropagation();
        var src = isStreaming ? (_partialRender.segments || []) : state.segments;
        var segCopy = src[idx];
        if (!segCopy) return;
        navigator.clipboard.writeText(segCopy.text).then(function () {
          showToast("Copied to clipboard");
        });
        return;
      }

      var tsEl = e.target.closest(".segment-timestamp");
      if (tsEl && row.contains(tsEl)) {
        e.stopPropagation();
        seekVideo(start);
        return;
      }

      var textEl = e.target.closest(".segment-text");
      if (textEl && row.contains(textEl)) {
        e.stopPropagation();
        if (state.editingTextEl === textEl) return;
        seekVideo(start);
        return;
      }
    });

    container.addEventListener("dblclick", function (e) {
      var row = e.target.closest(".segment-row");
      if (!row) return;
      var textEl = e.target.closest(".segment-text");
      if (!textEl || !row.contains(textEl)) return;
      e.stopPropagation();
      startSegmentEditing(textEl);
    });

    // Friction tooltip on hot segments (only while friction mode is on).
    // RAF-coalesced like the timeline canvas so getBoundingClientRect isn't
    // called on every mousemove event.
    container.addEventListener("mousemove", function (e) {
      if (_segTooltipRaf) return;
      var cx = e.clientX, cy = e.clientY, tgt = e.target;
      _segTooltipRaf = requestAnimationFrame(function () {
        _segTooltipRaf = 0;
        if (state.frictionMode === "off") { _hideFrictionTooltip(); return; }
        var row = tgt.closest && tgt.closest(".segment-row");
        if (!row) { _hideFrictionTooltip(); return; }
        var idx = parseInt(row.getAttribute("data-index"), 10);
        var seg = state.segments[idx];
        var frow = seg ? state.frictionBySegId[seg.id] : null;
        if (!frow || !(frow.score > 0)) { _hideFrictionTooltip(); return; }
        _showFrictionTooltip(frow, seg, cx, cy);
      });
    });
    container.addEventListener("mouseleave", function () { _hideFrictionTooltip(); });
  }

  // Cache marks made during streaming so they survive DOM rebuilds.
  // Each entry: { color, id, category, label, severity }. `version` is bumped on
  // any write to invalidate renderPartialSegments' append-only fast path.
  var _streamingMarks = {};
  var _streamingMarksVersion = 0;
  // Keyed by participant: a single boolean would make the first streaming
  // participant's load swallow every later one's, leaving a second live
  // stream's persisted marks unrendered until some task completed.
  var _streamingMarksLoadedByPid = {};

  function _bumpStreamingMarksVersion() {
    _streamingMarksVersion++;
    renderTimeline();
  }

  function _loadStreamingMarks(pid) {
    if (_streamingMarksLoadedByPid[pid]) return;
    _streamingMarksLoadedByPid[pid] = true;
    apiGet("api/marks").then(function (data) {
      if (!data.ok) return;
      if (data.categories) setMarkCategories(data.categories);
      var added = false;
      data.marks.forEach(function (m) {
        if (!m.valid || m.participant !== pid) return;
        if (_streamingMarks[m.segment_id]) return; // don't overwrite fresh marks
        var cat = MARK_CATEGORIES[m.category] || MARK_CATEGORIES.bookmark;
        _streamingMarks[m.segment_id] = {
          color: cat.color,
          id: m.id,
          category: m.category,
          label: m.label || "",
          severity: m.severity || "",
        };
        added = true;
      });
      if (added) _bumpStreamingMarksVersion();
    });
  }

  // getMarkForSegment resolves a segment's active mark — the persisted
  // seg.marks[0], falling back to the streaming-marks cache. Kept in the hub
  // (reads _streamingMarks); the video satellite reads it via TS.getMarkForSegment
  // for timeline marker rendering + tooltips.
  function getMarkForSegment(seg) {
    if (seg.marks && seg.marks.length > 0) return seg.marks[0];
    var streaming = _streamingMarks[seg.id];
    return streaming || null;
  }

  // ---- Video player + timeline (impl in transcripts-video.js) ----
  // Thin hub delegators forward to the satellite; selectParticipant, the segment
  // list, renderEmptyState/loadTranscript, the agents panel, search, and boot keep
  // calling these bare names. cancelPendingSeek / clearTimelineMarkers /
  // hasTimelineHover encapsulate the video-internal state the hub used to poke.
  function initVideoPlayer() { return TS.initVideoPlayer && TS.initVideoPlayer(); }
  function initTimelineCanvas() { return TS.initTimelineCanvas && TS.initTimelineCanvas(); }
  function initPipScroll() { return TS.initPipScroll && TS.initPipScroll(); }
  function initVideoSync() { return TS.initVideoSync && TS.initVideoSync(); }
  function initPlayerKeyboard() { return TS.initPlayerKeyboard && TS.initPlayerKeyboard(); }
  function renderTimeline() { return TS.renderTimeline && TS.renderTimeline(); }
  function updateTranscribeFill() { return TS.updateTranscribeFill && TS.updateTranscribeFill(); }
  function seekVideo() { return TS.seekVideo && TS.seekVideo.apply(null, arguments); }
  function scrollToSegment() { return TS.scrollToSegment && TS.scrollToSegment.apply(null, arguments); }
  function ignoreNextScroll() { return TS.ignoreNextScroll && TS.ignoreNextScroll(); }
  function applyCaptionMode() { return TS.applyCaptionMode && TS.applyCaptionMode(); }
  function _partForGlobal() { return TS._partForGlobal && TS._partForGlobal.apply(null, arguments); }
  function _partMediaUrl() { return TS._partMediaUrl && TS._partMediaUrl.apply(null, arguments); }
  function cancelPendingSeek() { return TS.cancelPendingSeek && TS.cancelPendingSeek(); }
  function clearTimelineMarkers() { return TS.clearTimelineMarkers && TS.clearTimelineMarkers(); }

  // ---- Inline segment editing ----

  function startSegmentEditing(textEl) {
    if (state.editingTextEl === textEl) return;
    if (state.editingTextEl) finishSegmentEditing(state.editingTextEl, false);

    state.editingTextEl = textEl;
    textEl.setAttribute("data-original-text", textEl.textContent);
    textEl.setAttribute("contenteditable", "true");
    textEl.setAttribute("spellcheck", "false");
    textEl.classList.add("segment-text-editing");

    function onBlur() {
      textEl.removeEventListener("blur", onBlur);
      textEl.removeEventListener("keydown", onKeydown);
      textEl.removeEventListener("paste", onPaste);
      finishSegmentEditing(textEl, false);
    }

    function onKeydown(e) {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        textEl.blur();
      }
      if (e.key === "Escape") {
        e.preventDefault();
        textEl.removeEventListener("blur", onBlur);
        textEl.removeEventListener("keydown", onKeydown);
        textEl.removeEventListener("paste", onPaste);
        finishSegmentEditing(textEl, true);
      }
    }

    function onPaste(e) {
      e.preventDefault();
      var text = (e.clipboardData || window.clipboardData).getData("text/plain");
      text = text.replace(/[\r\n]+/g, " ").trim();
      document.execCommand("insertText", false, text);
    }

    textEl.addEventListener("blur", onBlur);
    textEl.addEventListener("keydown", onKeydown);
    textEl.addEventListener("paste", onPaste);
  }

  function finishSegmentEditing(textEl, cancel) {
    var originalText = textEl.getAttribute("data-original-text") || "";
    var newText = textEl.textContent.trim();

    textEl.removeAttribute("contenteditable");
    textEl.removeAttribute("data-original-text");
    textEl.classList.remove("segment-text-editing");
    if (state.editingTextEl === textEl) state.editingTextEl = null;

    if (cancel || !newText || newText === originalText) {
      // During streaming, skip reload — next poll will re-render
      if (state.streamingParticipant) return;
      // Reload to restore clean word spans
      var pid = state.selectedParticipant;
      if (pid) loadTranscript(pid);
      return;
    }

    var corrections = extractCorrections(originalText, newText);
    if (corrections.length === 0) return;

    saveCorrections(corrections);
  }

  function extractCorrections(oldText, newText) {
    var oldWords = oldText.trim().split(/\s+/).filter(Boolean);
    var newWords = newText.trim().split(/\s+/).filter(Boolean);
    if (oldWords.join(" ") === newWords.join(" ")) return [];

    // LCS table for word-level alignment
    var m = oldWords.length, n = newWords.length;
    var dp = [];
    for (var i = 0; i <= m; i++) {
      dp[i] = [];
      for (var j = 0; j <= n; j++) {
        if (i === 0 || j === 0) dp[i][j] = 0;
        else if (oldWords[i - 1] === newWords[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1;
        else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }

    // Backtrack to get edit operations
    var ops = [];
    i = m; j = n;
    while (i > 0 || j > 0) {
      if (i > 0 && j > 0 && oldWords[i - 1] === newWords[j - 1]) {
        ops.unshift({ type: "eq" });
        i--; j--;
      } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
        ops.unshift({ type: "ins", word: newWords[j - 1] });
        j--;
      } else {
        ops.unshift({ type: "del", word: oldWords[i - 1] });
        i--;
      }
    }

    // Group consecutive non-equal ops into from→to correction pairs
    var corrections = [];
    var k = 0;
    while (k < ops.length) {
      if (ops[k].type !== "eq") {
        var fromParts = [];
        var toParts = [];
        while (k < ops.length && ops[k].type !== "eq") {
          if (ops[k].type === "del") fromParts.push(ops[k].word);
          else toParts.push(ops[k].word);
          k++;
        }
        if (fromParts.length > 0 && toParts.length > 0) {
          corrections.push({ from: fromParts.join(" "), to: toParts.join(" ") });
        }
      } else {
        k++;
      }
    }
    return corrections;
  }

  function saveCorrections(corrections) {
    var updated = 0, removed = 0;
    var chain = Promise.resolve();
    corrections.forEach(function (c) {
      chain = chain.then(function () {
        return apiPost("api/corrections", { from: c.from, to: c.to }).then(function (data) {
          if (data.ok) {
            if (data.removed) removed++;
            else if (data.correction) updated++;  // covers both new and updated
          }
        });
      });
    });
    chain.then(function () {
      var parts = [];
      if (updated) parts.push(updated === 1 ? "1 correction saved" : updated + " corrections saved");
      if (removed) parts.push(removed === 1 ? "1 reverted" : removed + " reverted");
      showToast(parts.join(", ") || "No changes");
      // During streaming, skip reload — corrections are persisted and will apply on completion
      if (state.streamingParticipant) return;
      var pid = state.selectedParticipant;
      if (pid) {
        loadTranscript(pid);
        loadCorrections();
      }
    }).catch(function () {
      showToast("Failed to save correction");
    });
  }

  // ---- Marks ----

  // Find the loaded (non-streaming) segment index for an id; -1 if absent.
  function _segmentIndexById(segmentId) {
    for (var i = 0; i < state.segments.length; i++) {
      if (state.segments[i].id === segmentId) return i;
    }
    return -1;
  }

  // Find the loaded segment carrying a mark id; null if absent.
  function _findSegmentByMarkId(markId) {
    for (var i = 0; i < state.segments.length; i++) {
      var marks = state.segments[i].marks || [];
      for (var j = 0; j < marks.length; j++) {
        if (marks[j].id === markId) return { idx: i, seg: state.segments[i], mark: marks[j] };
      }
    }
    return null;
  }

  // Repaint a single segment's mark dot + annotation badge in place (mirrors the
  // template in renderSegments) so a mark add/remove/recolor shows instantly
  // without a full loadTranscript round-trip. markObj null clears the mark.
  function _paintSegmentMark(idx, markObj) {
    var list = qs("#segmentList");
    if (!list) return;
    var row = list.querySelector('.segment-row[data-index="' + idx + '"]');
    if (!row) return;
    var dot = row.querySelector(".segment-mark");
    var textEl = row.querySelector(".segment-text");
    var oldBadge = textEl ? textEl.querySelector(".segment-anno-badge") : null;
    if (oldBadge) oldBadge.remove();
    var oldSevDot = row.querySelector(".segment-sev-dot");
    if (oldSevDot) oldSevDot.remove();
    if (!dot) return;
    if (markObj) {
      var cat = MARK_CATEGORIES[markObj.category] || MARK_CATEGORIES.bookmark;
      dot.classList.add("marked");
      dot.style.background = cat.color;
      if (markObj.severity) {
        var sevDot = document.createElement("span");
        sevDot.className = "segment-sev-dot " + severityClass(markObj.severity);
        sevDot.setAttribute("data-tooltip", markObj.severity);
        dot.insertAdjacentElement("afterend", sevDot);
      }
      if (markObj.label) {
        dot.setAttribute("data-tooltip", markObj.label);
        if (textEl) {
          var bgMix = "color-mix(in oklch, " + cat.color + " 18%, transparent)";
          var borderMix = "color-mix(in oklch, " + cat.color + " 50%, transparent)";
          var badge = document.createElement("span");
          badge.className = "segment-anno-badge";
          badge.style.cssText =
            "--anno-badge-fg:" + cat.color + ";--anno-badge-bg:" + bgMix + ";--anno-badge-border:" + borderMix;
          badge.textContent = markObj.label;
          textEl.insertBefore(badge, textEl.firstChild);
        }
      } else {
        dot.removeAttribute("data-tooltip");
      }
    } else {
      dot.classList.remove("marked");
      dot.style.background = "";
      dot.removeAttribute("data-tooltip");
    }
  }

  function toggleMark(segmentId) {
    var idx = _segmentIndexById(segmentId);
    var seg = idx >= 0 ? state.segments[idx] : null;
    if (!seg) {
      // No loaded row (shouldn't happen for the persisted path) — fall back to a
      // reload so the new mark still appears.
      apiPost("api/marks", { segment_ids: [segmentId], category: state.lastMarkCategory }).then(function (data) {
        if (data.ok && state.selectedParticipant) loadTranscript(state.selectedParticipant);
      });
      return;
    }
    // Optimistic: fill the dot now, reconcile the real id on success, revert on failure.
    var prevMarks = seg.marks || [];
    var provisional = { id: null, category: state.lastMarkCategory, label: "" };
    seg.marks = [provisional];
    _paintSegmentMark(idx, provisional);
    function revert() {
      seg.marks = prevMarks;
      _paintSegmentMark(idx, prevMarks[0] || null);
      showToast("Failed to mark");
    }
    apiPost("api/marks", { segment_ids: [segmentId], category: state.lastMarkCategory })
      .then(function (data) {
        if (data.ok && data.marks && data.marks.length > 0) {
          seg.marks = [data.marks[0]];
          showToast("Marked");
        } else {
          revert();
        }
      })
      .catch(revert);
  }

  function toggleMarkStreaming(segmentId, markEl) {
    var cat = MARK_CATEGORIES[state.lastMarkCategory] || MARK_CATEGORIES.bookmark;
    apiPost("api/marks", {
      segment_ids: [segmentId],
      category: state.lastMarkCategory,
    }).then(function (data) {
      if (data.ok && data.marks && data.marks.length > 0) {
        var m = data.marks[0];
        showToast("Marked");
        markEl.classList.add("marked");
        markEl.style.background = cat.color;
        _streamingMarks[segmentId] = {
          color: cat.color,
          id: m.id,
          category: m.category,
          label: m.label || "",
          severity: m.severity || "",
        };
        _bumpStreamingMarksVersion();
      }
    });
  }

  function removeMark(markId) {
    hideMarkPopover();
    // Streaming participant: keep the existing reload-on-success path.
    if (state.streamingParticipant) {
      apiDelete("api/marks/" + markId).then(function (data) {
        if (data.ok) {
          showToast("Mark removed");
          for (var key in _streamingMarks) {
            if (_streamingMarks[key].id === markId) {
              delete _streamingMarks[key];
              _bumpStreamingMarksVersion();
              break;
            }
          }
          pollTaskStatus();
        }
      });
      return;
    }
    // Optimistic: clear the dot now, restore on failure.
    var found = _findSegmentByMarkId(markId);
    if (!found) {
      apiDelete("api/marks/" + markId).then(function (data) {
        if (data.ok && state.selectedParticipant) loadTranscript(state.selectedParticipant);
      });
      return;
    }
    var prevMarks = found.seg.marks;
    found.seg.marks = [];
    _paintSegmentMark(found.idx, null);
    apiDelete("api/marks/" + markId)
      .then(function (data) {
        if (data.ok) {
          showToast("Mark removed");
        } else {
          found.seg.marks = prevMarks;
          _paintSegmentMark(found.idx, prevMarks[0] || null);
          showToast("Failed to remove mark");
        }
      })
      .catch(function () {
        found.seg.marks = prevMarks;
        _paintSegmentMark(found.idx, prevMarks[0] || null);
        showToast("Failed to remove mark");
      });
  }

  function updateMarkCategory(markId, category) {
    state.lastMarkCategory = category;
    hideMarkPopover();
    // Streaming participant: keep the existing reload-on-success path.
    if (state.streamingParticipant) {
      apiPut("api/marks/" + markId, { category: category }).then(function (data) {
        if (data.ok) {
          var cat = MARK_CATEGORIES[category] || MARK_CATEGORIES.bookmark;
          for (var key in _streamingMarks) {
            if (_streamingMarks[key].id === markId) {
              _streamingMarks[key].category = category;
              _streamingMarks[key].color = cat.color;
              _bumpStreamingMarksVersion();
              break;
            }
          }
          pollTaskStatus();
        }
      });
      return;
    }
    // Optimistic: recolor the dot now, restore on failure.
    var found = _findSegmentByMarkId(markId);
    if (!found) {
      apiPut("api/marks/" + markId, { category: category }).then(function (data) {
        if (data.ok && state.selectedParticipant) loadTranscript(state.selectedParticipant);
      });
      return;
    }
    var prevCategory = found.mark.category;
    found.mark.category = category;
    _paintSegmentMark(found.idx, found.mark);
    apiPut("api/marks/" + markId, { category: category }).catch(function () {
      found.mark.category = prevCategory;
      _paintSegmentMark(found.idx, found.mark);
      showToast("Failed to update mark");
    });
  }

  function updateMarkLabel(markId, label) {
    if (state.streamingParticipant) {
      apiPut("api/marks/" + markId, { label: label || null }).catch(function () {
        showToast("Failed to update mark");
      });
      for (var key in _streamingMarks) {
        if (_streamingMarks[key].id === markId) {
          _streamingMarks[key].label = label || "";
          break;
        }
      }
      return;
    }
    // Mirror updateMarkCategory/Severity: write the state and repaint
    // optimistically, restore on failure. Without the state write the badge
    // never appears, reopening the popover shows the old label, and a later
    // category/severity repaint resurrects it.
    var found = _findSegmentByMarkId(markId);
    if (!found) {
      apiPut("api/marks/" + markId, { label: label || null }).catch(function () {
        showToast("Failed to update mark");
      });
      return;
    }
    var prevLabel = found.mark.label;
    found.mark.label = label || "";
    _paintSegmentMark(found.idx, found.mark);
    apiPut("api/marks/" + markId, { label: label || null }).catch(function () {
      found.mark.label = prevLabel;
      _paintSegmentMark(found.idx, found.mark);
      showToast("Failed to update mark");
    });
  }

  function updateMarkSeverity(markId, severity) {
    var sev = severity || null;
    // Streaming participant: mirror updateMarkCategory — update the cache and
    // re-render on confirmed success. Severity changes the visible segment dot,
    // so a version bump alone isn't enough; pollTaskStatus re-renders the list
    // now instead of leaving a stale dot until the next scheduled poll.
    if (state.streamingParticipant) {
      apiPut("api/marks/" + markId, { severity: sev }).then(function (data) {
        if (data.ok) {
          for (var key in _streamingMarks) {
            if (_streamingMarks[key].id === markId) {
              _streamingMarks[key].severity = severity || "";
              _bumpStreamingMarksVersion();
              break;
            }
          }
          pollTaskStatus();
        }
      });
      return;
    }
    // Loaded participant: optimistically repaint the dot, restore on failure
    // (mirrors updateMarkCategory).
    var found = _findSegmentByMarkId(markId);
    if (!found) {
      apiPut("api/marks/" + markId, { severity: sev }).then(function (data) {
        if (data.ok && state.selectedParticipant) loadTranscript(state.selectedParticipant);
      });
      return;
    }
    var prevSeverity = found.mark.severity;
    found.mark.severity = sev;
    _paintSegmentMark(found.idx, found.mark);
    apiPut("api/marks/" + markId, { severity: sev }).catch(function () {
      found.mark.severity = prevSeverity;
      _paintSegmentMark(found.idx, found.mark);
      showToast("Failed to update mark");
    });
  }

  function showMarkPopover(anchorEl, segmentId, markObj) {
    // A provisional mark (optimistic paint, POST still in flight) has no id
    // yet — every popover action would hit api/marks/null. The POST resolves
    // in well under a click-reopen, so just don't open for it.
    if (markObj && markObj.id == null) return;
    var popover = qs("#markPopover");
    hideMarkPopover();

    // Build category pills
    var catContainer = popover.querySelector(".mark-popover-categories");
    catContainer.innerHTML = "";
    var cats = Object.keys(MARK_CATEGORIES);
    var pills = [];
    var activeIdx = 0;
    for (var i = 0; i < cats.length; i++) {
      (function (key, idx) {
        var cat = MARK_CATEGORIES[key];
        var pill = document.createElement("button");
        pill.type = "button";
        pill.className = "mark-cat-pill" + (markObj.category === key ? " active" : "");
        if (markObj.category === key) activeIdx = idx;
        pill.style.background = cat.color;
        pill.setAttribute("data-tooltip", cat.label);
        pill.setAttribute("aria-label", cat.label);
        pill.addEventListener("click", function (e) {
          e.stopPropagation();
          updateMarkCategory(markObj.id, key);
        });
        pills.push(pill);
        catContainer.appendChild(pill);
      })(cats[i], i);
    }

    // Label input
    var labelInput = popover.querySelector(".mark-popover-label");
    labelInput.value = markObj.label || "";
    labelInput._markId = markObj.id;
    labelInput.onblur = function () {
      var val = labelInput.value.trim();
      if (val !== (markObj.label || "")) {
        updateMarkLabel(markObj.id, val);
      }
    };
    labelInput.onkeydown = function (e) {
      if (e.key === "Enter") { e.preventDefault(); labelInput.blur(); hideMarkPopover(); }
      if (e.key === "Escape") { e.preventDefault(); hideMarkPopover(); }
    };

    // Severity dropdown (optional). Options come from CLIPGEN_CONFIG.severity
    // (mirrored from config.py SEVERITY_NUMERIC_TO_LABEL); a blank first option
    // means "no severity".
    var sevSelect = popover.querySelector(".mark-popover-severity");
    if (!sevSelect.options.length) {
      var blankOpt = document.createElement("option");
      blankOpt.value = "";
      blankOpt.textContent = "No severity";
      sevSelect.appendChild(blankOpt);
      var levels = CLIPGEN_CONFIG.severity || [];
      for (var s = 0; s < levels.length; s++) {
        var opt = document.createElement("option");
        opt.value = levels[s].label;
        opt.textContent = levels[s].label;
        sevSelect.appendChild(opt);
      }
    }
    sevSelect.value = markObj.severity || "";
    sevSelect.onchange = function () {
      // updateMarkSeverity owns the mark-state mutation (like updateMarkCategory)
      // so it can capture the previous value and roll back on failure. markObj is
      // the live seg.marks[0]/streaming-cache ref, so it stays in sync.
      updateMarkSeverity(markObj.id, sevSelect.value);
    };

    // Remove button
    var removeBtn = popover.querySelector(".mark-popover-remove");
    removeBtn.onclick = function (e) {
      e.stopPropagation();
      removeMark(markObj.id);
    };

    // Keyboard: arrows rove between category pills, Enter applies the focused
    // category (the pill's own click), Esc dismisses. Typing in the label input
    // keeps its own Enter/Esc handling and is skipped for arrow roving.
    popover.onkeydown = function (e) {
      if (e.key === "Escape") { e.preventDefault(); hideMarkPopover(); return; }
      if (document.activeElement === labelInput || document.activeElement === sevSelect) return;
      var idx = pills.indexOf(document.activeElement);
      if (e.key === "ArrowRight" || e.key === "ArrowDown") {
        e.preventDefault();
        pills[idx < 0 ? 0 : (idx + 1) % pills.length].focus();
      } else if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
        e.preventDefault();
        pills[idx < 0 ? pills.length - 1 : (idx - 1 + pills.length) % pills.length].focus();
      }
    };

    // Position below anchor
    var rect = anchorEl.getBoundingClientRect();
    popover.style.top = (rect.bottom + window.scrollY + 4) + "px";
    popover.style.left = (rect.left + window.scrollX - 4) + "px";
    popover.classList.remove("hidden");

    // Focus the current category so arrows/Enter work immediately. preventScroll
    // keeps the segment list from jumping when the popover opens via the keyboard.
    if (pills.length) {
      try { pills[activeIdx].focus({ preventScroll: true }); }
      catch (err) { pills[activeIdx].focus(); }
    }

    // Close on outside click (deferred so this click doesn't trigger it).
    // Track the timeout so a fast hideMarkPopover (e.g. Esc within the same
    // tick) cancels the pending attach instead of leaving a permanently
    // attached listener after the deferred .addEventListener fires.
    if (_popoverAttachTimer) clearTimeout(_popoverAttachTimer);
    document.removeEventListener("click", _popoverOutsideClick);
    _popoverAttachTimer = setTimeout(function () {
      _popoverAttachTimer = null;
      document.addEventListener("click", _popoverOutsideClick);
    }, 0);
  }

  function _popoverOutsideClick(e) {
    var popover = qs("#markPopover");
    if (popover && !popover.contains(e.target)) {
      hideMarkPopover();
    }
  }

  var _popoverAttachTimer = null;

  function hideMarkPopover() {
    var popover = qs("#markPopover");
    if (popover) popover.classList.add("hidden");
    if (_popoverAttachTimer) {
      clearTimeout(_popoverAttachTimer);
      _popoverAttachTimer = null;
    }
    document.removeEventListener("click", _popoverOutsideClick);
  }

  // ---- Search (impl in transcripts-search.js) ----
  // Thin hub delegator: boot wires initSearch. Search renders its own results
  // internally, so no renderSearchResults delegator is needed.
  function initSearch() {
    return TS.initSearch && TS.initSearch();
  }

  // ---- Participant pills (impl in transcripts-pills.js) ----
  // Thin hub delegators: loadParticipants / selectParticipant / pollTaskStatus
  // call renderPills; boot wires initPillOutsideClick / initPillWheelScroll; the
  // Transcribe All quick action enqueues through transcribeParticipants.
  function renderPills() { return TS.renderPills && TS.renderPills(); }
  function transcribeParticipants() { return TS.transcribeParticipants && TS.transcribeParticipants.apply(null, arguments); }
  function initPillOutsideClick() { return TS.initPillOutsideClick && TS.initPillOutsideClick(); }
  function initPillWheelScroll() { return TS.initPillWheelScroll && TS.initPillWheelScroll(); }

  // _confirmUncachedWhisperModels stays in the hub (model-install state:
  // _trModelsCache / _trModelsCachePromise / _whisperDownloadConfirmed +
  // confirmModelInstall). The pills satellite's _postTranscribe reaches it via
  // TS._confirmUncachedWhisperModels.

  // Confirm each distinct non-cached Whisper model in turn; any cancel aborts.
  function _confirmUncachedWhisperModels(uncached) {
    return uncached.reduce(function (chain, m) {
      return chain.then(function (okSoFar) {
        if (!okSoFar) return false;
        return confirmModelInstall({
          kind: "whisper",
          model: m.model,
          sizeMb: m.size_mb,
        }).then(function (ok) {
          if (ok) {
            _whisperDownloadConfirmed[m.model] = true;
            _trModelsCache = null;
            _trModelsCachePromise = null;
          }
          return ok;
        });
      });
    }, Promise.resolve(true));
  }

  // ---- Task polling ----
  // state.pollPoller → pollTaskStatus for Whisper jobs; summary/citations use
  // separate pollers (see file header).

  function startPolling() {
    if (state.pollPoller) return;
    // runImmediately is false to match the previous setInterval (first poll after POLL_INTERVAL).
    state.pollPoller = createPoller(pollTaskStatus, POLL_INTERVAL, {
      runImmediately: false,
      label: "transcripts.tasks",
    });
    state.pollPoller.start();
  }

  function stopPolling() {
    if (state.pollPoller) {
      state.pollPoller.stop();
      state.pollPoller = null;
    }
  }

  // Keyed by task id (NOT participant) so each task's completion is handled
  // exactly once. A participant can have several completed tasks over a session
  // (re-transcription creates a new task while old completed tasks linger in the
  // worker and are restored from the manifest), so a participant key would let a
  // stale completed task suppress the new run's completion transition.
  var _refreshedCompletedTaskIds = {};
  // After a whisper task flips to "completed", keep the main poll loop alive
  // for a few cycles so the summary "Generating…" state surfaces even if the
  // next /api/participants or /api/summary response races the in-flight slot.
  var _postCompletionGrace = 0;
  var POST_COMPLETION_GRACE_CYCLES = 4; // ~12s at POLL_INTERVAL=3000ms

  // Ceiling on how long a pill may sit in "Cancelling…". Nothing normal comes
  // close — the worst real case is a cold model load (~10s, uninterruptible)
  // plus one poll. This exists only so a wedged worker thread, whose task stays
  // "running" forever, hands the stop button back instead of leaving a dead pill.
  var CANCEL_PENDING_MAX_MS = 30000;

  // Drop the pending-cancel flag for every task the server no longer reports as
  // active. One predicate covers all four exits: the status flipped to cancelled
  // (the happy path), it flipped to completed or failed (the cancel raced the
  // finish line), or the task vanished from the list entirely (dismissed, worker
  // restart).
  function _sweepCancellingTasks() {
    var active = {};
    for (var i = 0; i < state.tasks.length; i++) {
      var t = state.tasks[i];
      if (t.status === "running" || t.status === "queued") active[t.id] = true;
    }
    var now = Date.now();
    for (var id in state.cancellingTasks) {
      if (!active[id] || now - state.cancellingTasks[id].at > CANCEL_PENDING_MAX_MS) {
        delete state.cancellingTasks[id];
      }
    }
  }

  function _anyAgentActive() {
    for (var i = 0; i < state.participants.length; i++) {
      var ag = state.participants[i].agents;
      if (ag && (ag.summary === "running" || ag.citations === "running" || ag.friction === "running")) return true;
    }
    return false;
  }

  // During the post-completion grace window an agent may register as running a
  // cycle or two after the whisper task completes. Re-load the selected
  // participant's summary so renderSummaryGenerating()/_startSummaryPoll fire
  // once the backend reports it generating; also refresh friction when its dep
  // is met. Guarded so it won't stomp an already-armed/rendered panel.
  function _rearmSelectedAgentPanels() {
    var pid = state.selectedParticipant;
    if (!pid) return;
    var p = _currentParticipant();
    var summaryRunning = !!(p && p.agents && p.agents.summary === "running");
    // _summaryPoller lives in the agents satellite; ask it whether the summary
    // poll is already armed (TS.isSummaryPolling) rather than reading the var.
    if (summaryRunning && !(TS.isSummaryPolling && TS.isSummaryPolling()) && !state.summaryText) loadSummary(pid);
    // Reload friction when the panel is empty, or when the deterministic-only
    // placeholder is showing and the summary-gated agent has now started/finished
    // — otherwise the deterministic blob keeps this guard shut and
    // the AI moments would never replace it (see agents/CODE-REVIEW.md poll gates).
    var frictionActive = !!(p && p.agents && (p.agents.friction === "running" || p.agents.friction === "done"));
    var showingDeterministic = !!(state.frictionData && state.frictionData.deterministic);
    if (!state.frictionGenerating && _frictionDepMet() &&
        (!state.frictionData || (showingDeterministic && frictionActive))) {
      loadFriction(pid);
    }
  }

  // Called from pill agent onStart/onStop after the API POST resolves.
  // Reloads participants so pill state reflects the new running/idle state
  // immediately, and starts the poll loop so the pill keeps updating without
  // waiting for the next external trigger (poll loop is otherwise gated on
  // `_anyAgentActive()` which reads stale state right after a manual run).
  function _refreshAgentStateNow() {
    loadParticipants().then(function () {
      if (_anyAgentActive()) startPolling();
    });
  }

  // Reconcile a participant still showing the streaming view against the backend:
  // once its whisper task is done and the transcript ready, swap the
  // "Transcribing… X%" footer for the finalized transcript and reveal the analysis
  // panel, mirroring selectParticipant's has_transcript path. Self-healing because
  // state.streamingParticipant clears only after the finalized transcript really
  // rendered, so a transient API failure retries next poll instead of freezing.
  function _finalizeStreamingIfComplete(pid) {
    var running = false;
    var completed = false;
    for (var i = 0; i < state.tasks.length; i++) {
      var t = state.tasks[i];
      if (t.participant !== pid) continue;
      if (t.status === "running" || t.status === "queued") running = true;
      else if (t.status === "completed") completed = true;
    }
    // Still transcribing → leave the streaming view up; a later poll re-enters.
    if (running) return;
    // No running task and nothing completed (failed/cancelled/dismissed) → drop
    // the streaming flag so the normal empty/failed rendering can take over and
    // the poll loop can wind down (matches the prior behaviour here).
    if (!completed) {
      state.streamingParticipant = null;
      // The partial rows stay — they are real transcript the user may still want
      // to read — but the footer under them has to go, or a cancelled run leaves
      // a frozen "Cancelling…" line forever: nothing calls renderSegments() on
      // this path, and renderSegmentsImpl is the only other place the indicator
      // is removed. Cancel the queued insert first, or a RAF scheduled while the
      // tab was hidden re-inserts the footer right after we remove it.
      _cancelStreamingIndicator();
      var ind = document.querySelector("#segmentList .streaming-indicator");
      if (ind) ind.parentNode.removeChild(ind);
      return;
    }
    var ver = state.participantReqVer;
    apiGet("api/transcript/" + pid).then(function (data) {
      if (ver !== state.participantReqVer || state.selectedParticipant !== pid) return;
      // Transcript not merged yet (rare race) → keep streamingParticipant set so
      // the next poll retries rather than leaving a frozen footer.
      if (!(data.ok && data.segments && data.segments.length > 0)) return;
      state.streamingParticipant = null;
      state.segments = data.segments;
      state.activeSegmentIndex = -1;
      renderSegments();
      renderTimeline();
      _setAnalysisPanelVisible(true);
      _restoreActiveTab(pid);
      loadSummary(pid);
      loadFriction(pid);
    });
  }

  function pollTaskStatus() {
    apiGet("api/transcribe/status").then(function (data) {
      if (!data.ok) return;
      state.tasks = data.tasks;
      // Before updateTranscribeFill / updateStatusIndicator / renderPills below
      // — all three read the flag, and a stale one would keep the band off and
      // the pill inert for a full extra tick.
      _sweepCancellingTasks();
      if (_anyTxEtaActive()) _txEtaTicker.ensure();

      // Fill the timeline in sync with the selected participant's transcription
      // progress (starts/updates/clears the band purely from state.tasks).
      updateTranscribeFill();

      // Snapshot before _finalizeStreamingIfComplete can clear streamingParticipant
      // in its async /api/transcript callback; the newlyCompleted refresh must not
      // double-fire loadTranscript/loadSummary/loadFriction for that transition.
      var wasStreamingSelected =
        state.streamingParticipant === state.selectedParticipant;

      // Re-render the status circle immediately so completed tasks reflect
      // before the async loadParticipants()/loadTranscript() chain resolves.
      // This is what keeps the indicator from freezing at "95%".
      updateStatusIndicator();

      // Stream partial segments for the selected participant's running task
      var selectedRunningTask = null;
      if (state.selectedParticipant) {
        data.tasks.forEach(function (t) {
          if (t.participant === state.selectedParticipant && t.status === "running" && t.partial_count) {
            selectedRunningTask = t;
          }
        });
        // Keep the empty pane's wait text current (queued → loading model →
        // starting) while nothing has streamed yet; harmless when hidden.
        // Same running-over-queued pick as the status indicator, so the two
        // never disagree when duplicate tasks exist for the participant.
        _setTranscriptEmptyText(_taskForSelectedParticipant());
      }
      if (selectedRunningTask) {
        state.streamingParticipant = state.selectedParticipant;
        _loadStreamingMarks(state.selectedParticipant);
        _syncStreamSegs(selectedRunningTask, function (segs) {
          if (segs.length > 0) renderPartialSegments(segs, selectedRunningTask.progress);
        });
      } else if (state.streamingParticipant) {
        // The selected participant's streaming view is up but it has no running
        // task. Finalize in place once its transcript is ready. This is
        // state-driven and retried every poll (not a one-shot tied to the
        // newlyCompleted de-dup), so a transient /api/participants or
        // /api/transcript failure can't leave the "Transcribing… X%" footer
        // frozen forever. It owns clearing state.streamingParticipant.
        _finalizeStreamingIfComplete(state.streamingParticipant);
      }

      var hasActive = false;
      var newlyCompleted = [];
      data.tasks.forEach(function (t) {
        if (t.status === "queued" || t.status === "running") hasActive = true;
        if (t.status === "completed" && !_refreshedCompletedTaskIds[t.id]) {
          newlyCompleted.push(t.participant);
          _refreshedCompletedTaskIds[t.id] = true;
        }
      });

      // Refresh participants and transcript as each task completes.
      // Thinking-agents (summary → citations) are spawned on whisper completion
      // and on server startup, so we always refresh after anything completes
      // or if any agent is currently running on any pill.
      // Refresh during the grace window too, so participants are re-fetched each
      // cycle until the summary agent registers as running (or grace expires).
      var needsRefresh =
        newlyCompleted.length > 0 || _anyAgentActive() || _postCompletionGrace > 0;
      if (newlyCompleted.length > 0) {
        // Re-arm on every fresh completion so a multi-participant queue keeps
        // extending the grace window.
        _postCompletionGrace = POST_COMPLETION_GRACE_CYCLES;
        _streamingMarks = {};
        _streamingMarksLoadedByPid = {};
        _bumpStreamingMarksVersion();
      }
      if (needsRefresh) {
        loadParticipants().then(function () {
          if (newlyCompleted.length > 0 && state.selectedParticipant &&
              newlyCompleted.indexOf(state.selectedParticipant) >= 0 &&
              !wasStreamingSelected) {
            // The selected participant finished but was NOT mid-stream in this
            // view (e.g. it was queued/idle when it completed), so reveal the
            // analysis panel and load the finalized transcript here, mirroring
            // selectParticipant's has_transcript path. The streaming→done case
            // is owned by _finalizeStreamingIfComplete (state-driven + retried),
            // so skip when wasStreamingSelected to avoid double-firing after it
            // clears streamingParticipant and loads summary/friction first.
            _setAnalysisPanelVisible(true);
            _restoreActiveTab(state.selectedParticipant);
            loadTranscript(state.selectedParticipant);
            loadSummary(state.selectedParticipant);
            loadFriction(state.selectedParticipant);
          } else if (state.selectedParticipant) {
            // Summary/citations/friction may be auto-chaining server-side after
            // an earlier completion, or registering during the grace window;
            // re-arm the selected participant's panels so the running state
            // surfaces without a manual reload.
            _rearmSelectedAgentPanels();
          }
          updateStatusIndicator();
          // Agents typically kick in right after whisper completes; keep the
          // poll alive across the grace window so dot transitions (running →
          // done → next) are seen even before the agent registers as running.
          // Also stay alive while a streaming view still needs finalizing, so
          // _finalizeStreamingIfComplete keeps retrying until it renders.
          if (_anyAgentActive() || _postCompletionGrace > 0 || state.streamingParticipant) startPolling();
          else if (!hasActive) {
            stopPolling();
            _refreshedCompletedTaskIds = {};
          }
        });
      }

      if (hasActive || _anyAgentActive() || _postCompletionGrace > 0 || state.streamingParticipant) {
        startPolling();
      } else if (!needsRefresh) {
        stopPolling();
        _refreshedCompletedTaskIds = {};
      }

      if (!hasActive && _hadActiveTranscriptionLastPoll) {
        refreshTranscriptionModelHintOnce();
        // Transcription just finished — re-validate downloads against real
        // cache state (downloaded → no prompt, failed → prompt again).
        _forgetWhisperDownloadAgreements();
      }
      _hadActiveTranscriptionLastPoll = hasActive;

      // Count down the grace window at the end so the current cycle still
      // counts as "in grace" for the decisions above.
      if (_postCompletionGrace > 0) _postCompletionGrace--;

      renderPills();
    });
  }

  // ---- Corrections modal (impl in transcripts-corrections.js) ----
  // Thin hub delegators forward to the satellite; the inline-edit
  // saveCorrections() flow (loadCorrections) and boot (initCorrectionsModal)
  // keep calling these bare names.
  function initCorrectionsModal() {
    return TS.initCorrectionsModal && TS.initCorrectionsModal();
  }
  function loadCorrections() {
    return TS.loadCorrections && TS.loadCorrections();
  }

  // ---- Settings (shared modal lives in settings-modal.js) ----

  // Models are also fetched for per-pill model overrides; keep a tiny
  // cached fetcher here. The shared modal maintains its own cache.
  var _trModelsCache = null;
  var _trModelsCachePromise = null;
  // Whisper models the user has already agreed to download this session. The
  // models response is cached and a just-confirmed model won't read back as
  // cached until its background download finishes, so without this we'd
  // re-prompt on every transcription. Keyed by model name.
  var _whisperDownloadConfirmed = {};
  // Serializes confirmModelInstall() calls: there is one shared modal element,
  // so overlapping callers (e.g. prewarm + an agent run) must take turns.
  var _modelInstallChain = Promise.resolve();

  function _trFetchModels() {
    if (_trModelsCache) return Promise.resolve(_trModelsCache);
    if (_trModelsCachePromise) return _trModelsCachePromise;
    _trModelsCachePromise = apiGet("/api/models")
      .then(function (data) {
        // Don't pin a result where Ollama wasn't reachable — otherwise the
        // agent gate and pickers stay blind to installed models for the whole
        // session. Reset so the next call re-fetches once the server is up.
        if (data && data.ok && !(data.ollama && data.ollama.available === false)) {
          _trModelsCache = data;
        } else {
          _trModelsCachePromise = null;
        }
        return data;
      })
      .catch(function () { _trModelsCachePromise = null; return null; });
    return _trModelsCachePromise;
  }

  // Audio-track layout per participant, shared by the volume mixer (which only
  // ever wants the *selected* participant) and the pill transcribe picker (which
  // opens for any pill). One endpoint, one cache — two independent fetchers of
  // /api/audio-info would drift. Keyed by pid + video_version (an mtime sum from
  // /api/participants) so replacing a source file re-probes; bounded by the
  // participant count. Never cached on a failed response.
  var _trAudioInfoCache = {};
  var _trAudioInfoPromises = {};

  function _trAudioInfoKey(pid, videoVersion) {
    return pid + ":" + (videoVersion == null ? "" : videoVersion);
  }

  // Synchronous peek — returns the cached layout or null. Lets the pill popover
  // render its track row in the same tick when warm, so the 3 s poll's pane
  // rebuild doesn't strobe the row in and out.
  function audioInfoCached(pid, videoVersion) {
    return _trAudioInfoCache[_trAudioInfoKey(pid, videoVersion)] || null;
  }

  function _trFetchAudioInfo(pid, videoVersion) {
    var key = _trAudioInfoKey(pid, videoVersion);
    if (_trAudioInfoCache[key]) return Promise.resolve(_trAudioInfoCache[key]);
    if (_trAudioInfoPromises[key]) return _trAudioInfoPromises[key];
    _trAudioInfoPromises[key] = apiGet("api/audio-info/" + encodeURIComponent(pid))
      .then(function (data) {
        if (!data || !data.ok) {
          delete _trAudioInfoPromises[key];
          return null;
        }
        var info = {
          tracks: data.audio_tracks || [],
          count: data.audio_track_count || 0,
          auto: data.auto_index || 0
        };
        _trAudioInfoCache[key] = info;
        return info;
      })
      .catch(function () { delete _trAudioInfoPromises[key]; return null; });
    return _trAudioInfoPromises[key];
  }

  // ---- Local-model install confirmation ----
  // Both whisper transcription models and Ollama agent models are "local
  // models" that get installed on demand. We never download one silently:
  // confirmModelInstall() gates every install behind an explicit dialog.

  function _trFormatModelSize(mb) {
    if (!mb || mb <= 0) return "";
    if (mb >= 1024) return (mb / 1024).toFixed(1) + " GB";
    return Math.round(mb) + " MB";
  }

  // Stream an Ollama pull, reporting progress dicts via onProgress. Resolves
  // true on success, false on failure/cancellation. isCancelled() is polled
  // each tick so a dismissed dialog stops the poll (the server-side pull keeps
  // running and resumes from cached layers). The poll also self-terminates
  // after a run of unanswered status checks so it can never leak forever.
  function installOllamaModel(model, onProgress, isCancelled) {
    return apiPost("api/models/ollama/pull", { model: model }).then(function (data) {
      if (!data || !data.ok) return false;
      return new Promise(function (resolve) {
        // createPoller, not a raw setInterval: the install dialog can sit
        // open in a backgrounded tab for minutes, and this must pause with
        // it instead of streaming 1 Hz requests at the server.
        var misses = 0;
        var poller = createPoller(function () {
          if (isCancelled && isCancelled()) {
            poller.stop();
            resolve(false);
            return;
          }
          apiGet("api/models/ollama/pull-status?model=" + encodeURIComponent(model))
            .then(function (st) {
              if (!st || !st.ok || !st.found) {
                if (++misses >= 20) { poller.stop(); resolve(false); }
                return;
              }
              misses = 0;
              if (onProgress) onProgress(st);
              if (st.done) {
                poller.stop();
                resolve(!!st.succeeded);
              }
            })
            .catch(function () {
              if (++misses >= 20) { poller.stop(); resolve(false); }
            });
        }, 1000, { runImmediately: true, label: "transcripts.ollamaPull" });
        poller.start();
      });
    }).catch(function () { return false; });
  }

  // Stream the managed install of the Ollama CLI itself (macOS): same
  // start-then-poll shape as installOllamaModel, against the unkeyed
  // install/install-status endpoints. `already_installing` attaches to the
  // in-flight install rather than failing, so a reopened dialog resumes its
  // progress display.
  function installOllamaRuntime(onProgress, isCancelled) {
    return apiPost("api/models/ollama/install", {}).then(function (data) {
      if (!data || !data.ok) return false;
      if (data.already_installed) return true;
      return new Promise(function (resolve) {
        // Same createPoller rationale as installOllamaModel above.
        var misses = 0;
        var poller = createPoller(function () {
          if (isCancelled && isCancelled()) {
            poller.stop();
            resolve(false);
            return;
          }
          apiGet("api/models/ollama/install-status")
            .then(function (st) {
              if (!st || !st.ok || !st.found) {
                if (++misses >= 20) { poller.stop(); resolve(false); }
                return;
              }
              misses = 0;
              if (onProgress) onProgress(st);
              if (st.done) {
                poller.stop();
                resolve(!!st.succeeded);
              }
            })
            .catch(function () {
              if (++misses >= 20) { poller.stop(); resolve(false); }
            });
        }, 1000, { runImmediately: true, label: "transcripts.ollamaInstall" });
        poller.start();
      });
    }).catch(function () { return false; });
  }

  // Show the confirm/install dialog. Resolves true when the model is available
  // to use (whisper: user agreed to the download; ollama: pull succeeded),
  // false when the user cancels or the install fails. Calls are serialized
  // (see _modelInstallChain) so concurrent callers never share the one modal.
  function confirmModelInstall(opts) {
    var run = function () { return _confirmModelInstallNow(opts); };
    var result = _modelInstallChain.then(run, run);
    // Advance the chain when this dialog settles, swallowing its outcome so a
    // cancelled/failed dialog doesn't break the queue for the next caller.
    _modelInstallChain = result.then(function () {}, function () {});
    return result;
  }

  function _confirmModelInstallNow(opts) {
    return new Promise(function (resolve) {
      var cancelled = false;
      var modal = qs("#modelInstallModal");
      var titleEl = qs("#modelInstallTitle");
      var msgEl = qs("#modelInstallMessage");
      var progress = qs("#modelInstallProgress");
      var barFill = qs("#modelInstallBarFill");
      var progressText = qs("#modelInstallProgressText");
      var cancelBtn = qs("#modelInstallCancel");
      var confirmBtn = qs("#modelInstallConfirm");
      var altBtn = qs("#modelInstallAlt");
      var hintEl = qs("#modelInstallHint");

      // The progress line shimmers only while the install/check is actually
      // moving; pass working=false for anything terminal (failed, installed,
      // "close and retry") so a stalled dialog stops claiming to be busy.
      function setProgressText(text, working) {
        progressText.classList.toggle("cg-shimmer", working !== false);
        progressText.textContent = text;
      }

      progress.classList.add("hidden");
      barFill.style.width = "0%";
      setProgressText("", false);
      hintEl.classList.add("hidden");
      hintEl.textContent = "";
      cancelBtn.disabled = false;
      cancelBtn.textContent = "Cancel";
      confirmBtn.disabled = false;
      confirmBtn.classList.remove("hidden");
      altBtn.classList.add("hidden");
      altBtn.disabled = false;

      if (opts.kind === "ollama-runtime") {
        // Ollama missing or down is a different problem from "the model isn't
        // pulled", and the one case the gate used to stay silent about. On macOS
        // clipgen can download the CLI itself (this dialog *is* the consent);
        // elsewhere show the commands and offer a re-check. Not status.message —
        // that is written for a panel banner and ends in "then Refresh", the wrong
        // instruction beside a button that does the re-check itself.
        if (opts.state === "missing") {
          titleEl.textContent = "Ollama isn't installed";
          if (opts.canInstall) {
            var dlSize = opts.installSizeMb ? " (~" + _trFormatModelSize(opts.installSizeMb) + ")" : "";
            msgEl.textContent = "clipgen couldn't find Ollama on this machine. " +
              "The AI summaries, citations and reports need it — everything else works without it. " +
              "clipgen can download it for you" + dlSize + ", or install it yourself:";
            confirmBtn.textContent = "Download & install";
            altBtn.classList.remove("hidden");
          } else {
            msgEl.textContent = "clipgen couldn't find Ollama on this machine. " +
              "The AI summaries, citations and reports need it — everything else works without it.";
            confirmBtn.textContent = "I've installed it — retry";
          }
          if (opts.hint && opts.hint.length) {
            hintEl.textContent = opts.hint.join("\n");
            hintEl.classList.remove("hidden");
          }
        } else {
          titleEl.textContent = "Ollama isn't running";
          msgEl.textContent = "Ollama is installed but isn't answering at " +
            (opts.baseUrl || "localhost") + ". clipgen can start it for you.";
          confirmBtn.textContent = "Start Ollama";
        }
      } else if (opts.kind === "whisper") {
        titleEl.textContent = "Download transcription model?";
        var size = opts.sizeMb ? " (~" + _trFormatModelSize(opts.sizeMb) + ")" : "";
        if (opts.prewarm) {
          msgEl.textContent = 'The "' + opts.model + '" transcription model' + size +
            " isn't downloaded yet. Download it now so transcription is ready to start? It will be stored locally.";
        } else {
          msgEl.textContent = 'The "' + opts.model + '" transcription model' + size +
            " isn't downloaded yet. It will be downloaded and stored locally before transcription begins.";
        }
        confirmBtn.textContent = "Download";
      } else {
        titleEl.textContent = "Install AI model?";
        msgEl.textContent = 'The Ollama model "' + opts.model + '" used by the ' +
          (opts.agentKey || "analysis") +
          " agent isn't installed. Install it now? This downloads the model locally and may take several minutes.";
        confirmBtn.textContent = "Install";
      }

      function cleanup() {
        cancelBtn.removeEventListener("click", onCancel);
        confirmBtn.removeEventListener("click", onConfirm);
        altBtn.removeEventListener("click", onRuntimeConfirm);
        closeBlockingModal(modal);
      }
      function close(result) {
        cancelled = true; // stop any in-flight pull poll and its late callbacks
        cleanup();
        modal.classList.add("hidden");
        resolve(result);
      }
      function onCancel() { close(false); }

      // "Start Ollama" / "I've installed it — retry": both end in the same
      // question — is Ollama usable now? Re-fetch rather than trusting the
      // start call, so a server that spawned but never came up still reads as
      // a failure.
      function onRuntimeConfirm() {
        confirmBtn.disabled = true;
        altBtn.disabled = true;
        progress.classList.remove("hidden");
        setProgressText(opts.state === "stopped" ? "Starting…" : "Checking…");
        var step = opts.state === "stopped"
          ? apiPost("api/models/ollama/start", {}).catch(function () { return null; })
          : Promise.resolve(null);
        // Re-enabling the button and relabelling Cancel is the only way out of
        // the "Checking…" state, so it has to happen on *every* ending — a
        // /api/models that rejects left the dialog stuck mid-check with both
        // buttons dead.
        function stillUnavailable(text) {
          if (cancelled) return;
          setProgressText(text, false);
          confirmBtn.disabled = false;
          altBtn.disabled = false;
          cancelBtn.textContent = "Close";
        }

        step.then(function () {
          _trModelsCache = null;
          _trModelsCachePromise = null;
          return _trFetchModels();
        }).then(function (data) {
          if (cancelled) return;
          // _trFetchModels resolves null rather than rejecting when the fetch
          // fails, and clipgenOllamaStatus(null) is "ok" by design (an unknown
          // state must never block an action elsewhere). Here that default is
          // wrong in the other direction: the user asked "is it ready now?" and
          // a failed check is not a yes.
          if (!data || !data.ok) {
            stillUnavailable("Couldn't check — clipgen didn't answer. Try again.");
            return;
          }
          if (clipgenOllamaStatus(data.ollama).state === "ok") {
            showToast("Ollama is ready");
            close(true);
            return;
          }
          stillUnavailable(opts.state === "stopped"
            ? "Ollama still isn't responding."
            : "Still not finding Ollama. Open a new terminal and check `ollama --version`.");
        }).catch(function () {
          stillUnavailable("Couldn't check — clipgen didn't answer. Try again.");
        });
      }

      // Download the Ollama CLI itself, then chain into the same "is it usable
      // now?" sequence the Start button runs: serve → refetch → status check.
      // The server-side install keeps running if the dialog is dismissed, and
      // reopening re-attaches via `already_installing` + the status poll.
      function onManagedInstall() {
        confirmBtn.classList.add("hidden");
        altBtn.classList.add("hidden");
        hintEl.classList.add("hidden");
        progress.classList.remove("hidden");
        setProgressText("Starting download…");
        installOllamaRuntime(function (st) {
          if (st.total > 0) {
            var pct = Math.max(0, Math.min(100, Math.round((st.completed / st.total) * 100)));
            barFill.style.width = pct + "%";
            setProgressText((st.status || "Downloading") + ": " + pct + "%");
          } else {
            setProgressText(st.status || "Working…");
          }
        }, function () { return cancelled; }).then(function (installed) {
          if (cancelled) return;
          if (!installed) {
            setProgressText("Installation failed. Retry, or install Ollama yourself:", false);
            if (opts.hint && opts.hint.length) {
              hintEl.textContent = opts.hint.join("\n");
              hintEl.classList.remove("hidden");
            }
            // Both buttons come back, same as stillUnavailable() above: hiding
            // them left the only recovery path behind a dismiss-and-redo of
            // the whole AI action. Retrying is safe now that the server no
            // longer reports a broken install as already_installed.
            progress.classList.add("hidden");
            confirmBtn.classList.remove("hidden");
            confirmBtn.disabled = false;
            altBtn.classList.remove("hidden");
            altBtn.disabled = false;
            cancelBtn.textContent = "Close";
            return;
          }
          setProgressText("Installed. Starting Ollama…");
          apiPost("api/models/ollama/start", {}).catch(function () { return null; })
            .then(function () {
              _trModelsCache = null;
              _trModelsCachePromise = null;
              return _trFetchModels();
            }).then(function (data) {
              if (cancelled) return;
              if (data && data.ok && clipgenOllamaStatus(data.ollama).state === "ok") {
                showToast("Ollama installed and running");
                close(true);
                return;
              }
              setProgressText("Installed, but Ollama isn't answering yet. Close and retry the action.", false);
              cancelBtn.textContent = "Close";
            }).catch(function () {
              if (cancelled) return;
              setProgressText("Installed, but couldn't confirm Ollama is running. Close and retry the action.", false);
              cancelBtn.textContent = "Close";
            });
        });
      }

      function onConfirm() {
        if (opts.kind === "ollama-runtime") {
          if (opts.state === "missing" && opts.canInstall) { onManagedInstall(); return; }
          onRuntimeConfirm();
          return;
        }
        if (opts.kind === "whisper") { close(true); return; }
        // Ollama: kick off the pull and show progress in place. Cancel stays
        // enabled so the user can dismiss while it runs.
        confirmBtn.classList.add("hidden");
        progress.classList.remove("hidden");
        setProgressText("Starting…");
        installOllamaModel(opts.model, function (st) {
          if (st.total > 0) {
            var pct = Math.max(0, Math.min(100, Math.round((st.completed / st.total) * 100)));
            barFill.style.width = pct + "%";
            setProgressText((st.status || "Downloading") + ": " + pct + "%");
          } else {
            setProgressText(st.status || "Working…");
          }
        }, function () { return cancelled; }).then(function (ok) {
          if (cancelled) return; // dialog dismissed mid-pull — no toast, no re-close
          if (ok) {
            _trModelsCache = null;
            _trModelsCachePromise = null;
            showToast("Model installed");
            close(true);
          } else {
            // Leave the dialog open so the user can read the failure and
            // dismiss it; Cancel now resolves false.
            setProgressText("Installation failed. Check that Ollama is running.", false);
            cancelBtn.textContent = "Close";
          }
        });
      }

      cancelBtn.addEventListener("click", onCancel);
      confirmBtn.addEventListener("click", onConfirm);
      altBtn.addEventListener("click", onRuntimeConfirm);
      modal.classList.remove("hidden");
      // Escape and backdrop click both cancel (no focus trap — matches prior
      // behavior for this lightweight progress dialog).
      openBlockingModal(modal, { onEscape: onCancel, onBackdropClick: onCancel });
    });
  }

  // Gate an agent run on Ollama being usable and its model installed; resolves
  // true to proceed, false to abort.
  //
  // An unreachable Ollama gets its own dialog rather than falling through to the
  // downstream "model unavailable" error, which only appears on two agent
  // surfaces and never says how to fix anything — and it is the state a
  // first-time user is actually in. An *unknown* state (fetch failed) still
  // passes: never block an action on a question we couldn't ask.
  function ensureAgentModelInstalled(agentKey) {
    return _trFetchModels().then(function (data) {
      var oll = data && data.ollama;
      if (!oll) return true;
      var status = clipgenOllamaStatus(oll);
      if (status.state !== "ok") {
        return confirmModelInstall({
          kind: "ollama-runtime",
          state: status.state,
          baseUrl: status.baseUrl,
          hint: status.hint,
          canInstall: status.canInstall,
          installSizeMb: status.installSizeMb,
        }).then(function (recovered) {
          if (!recovered) return false;
          // Ollama is up now, but that says nothing about the agent's model —
          // the payload we started from listed no models at all. Ask again
          // against the fresh one the dialog just refetched.
          return _trFetchModels().then(function (fresh) {
            return _ensureModelFromPayload(fresh, agentKey);
          });
        });
      }
      return _ensureModelFromPayload(data, agentKey);
    }).catch(function () { return true; });
  }

  // The "is this agent's model pulled?" half of the gate, against an already
  // fetched /api/models payload. Resolves true to proceed.
  function _ensureModelFromPayload(data, agentKey) {
    var oll = data && data.ollama;
    if (!oll) return true;
    var agents = oll.agents || [];
    var info = null;
    for (var i = 0; i < agents.length; i++) {
      if (agents[i].key === agentKey) { info = agents[i]; break; }
    }
    if (!info || info.installed || !info.model) return true;
    return confirmModelInstall({ kind: "ollama", agentKey: agentKey, model: info.model });
  }

  function _applySettingsSnapshot(applied, settings) {
    var nextCats = null;
    if (applied && applied.MARK_CATEGORIES) {
      nextCats = applied.MARK_CATEGORIES;
    } else if (settings) {
      for (var i = 0; i < settings.length; i++) {
        if (settings[i].name === "MARK_CATEGORIES") {
          nextCats = settings[i].value;
          break;
        }
      }
    }
    if (nextCats) {
      setMarkCategories(nextCats);
      if (!MARK_CATEGORIES[state.lastMarkCategory]) {
        var firstKey = Object.keys(MARK_CATEGORIES)[0];
        state.lastMarkCategory = firstKey || "bookmark";
      }
      // Refresh streaming mark colors and re-render the visible transcript.
      for (var sid in _streamingMarks) {
        var sm = _streamingMarks[sid];
        var cat = MARK_CATEGORIES[sm.category];
        if (cat) sm.color = cat.color;
      }
      _bumpStreamingMarksVersion();
      hideMarkPopover();
      if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
    }
    if (applyCrossRefSetting(applied, settings)) rerenderCrossRefs();
  }

  // Shared by the settings modal and the command palette's cross-ref command.
  function rerenderCrossRefs() {
    if (state.segments.length > 0) renderSegments();
  }
  window.clipgenRerenderCrossRefs = rerenderCrossRefs;

  // A changed transcription model invalidates the prewarm guards: the new
  // model may be uncached and must get its own download confirmation, so we
  // clear the "already posted/declined" state and let prewarm re-offer it.
  // Clearing the declined model also means re-selecting a previously-declined
  // model gets a fresh prompt (the user explicitly chose it again).
  function _onTranscribeModelMaybeChanged(newModel) {
    if (newModel === undefined || newModel === null) return;
    if (newModel !== _lastTranscribeModel) {
      _transcriptionWarmupPosted = false;
      _prewarmDownloadPrompting = false;
      _prewarmDeclinedModel = null;
    }
    _lastTranscribeModel = newModel;
  }

  function _settingValueFromRecords(settings, name) {
    if (!settings) return undefined;
    for (var i = 0; i < settings.length; i++) {
      if (settings[i].name === name) return settings[i].value;
    }
    return undefined;
  }

  function initTranscriptSettings() {
    if (!window.wireSettingsButton) return;
    window.wireSettingsButton({
      initialTab: "Transcription",
      onApply: function (applied, settings) {
        _trModelsCache = null;
        _trModelsCachePromise = null;
        _onTranscribeModelMaybeChanged(
          applied && applied.TRANSCRIBE_MODEL !== undefined
            ? applied.TRANSCRIBE_MODEL
            : _settingValueFromRecords(settings, "TRANSCRIBE_MODEL")
        );
        _applySettingsSnapshot(applied, settings);
      },
    });
  }

  // ---- Clip marked lines ----
  //
  // Cuts one clip per cluster of marked lines through Studio's
  // ../studio/api/generate-intake — the endpoint Studio's Transcript Intake tab
  // uses — so output lands in clipgen_manifest.json as if queued there and this
  // page needs no generation backend. Unlike Studio's queue path it also sends
  // `text`/`label`, which is what gives each artifact a readable description
  // instead of a bare category (see _process_intake_item in server.py).

  // Mirrors Studio's #trIntakeClusterThreshold default so identical marks
  // cluster identically on both pages. Padding defaults to 0 for the same
  // reason — the spans then match Studio's exactly. Segment boundaries are
  // word/energy-tight since TRANSCRIBE_WORD_TIMESTAMPS + TRANSCRIBE_EDGE_SNAP,
  // so pad-0 cuts no longer clip the first or last word; the pad remains for
  // users who want breathing room around the speech.
  var CLIP_MARKS_DEFAULT_GAP_SECONDS = 10;
  var CLIP_MARKS_DEFAULT_PAD_SECONDS = 0;

  // { done, failed, total, abort } while a batch streams; null when idle. The
  // modal can be dismissed mid-run (the run continues) and reopened onto the
  // live progress, so this outlives the dialog.
  var _clipMarksRun = null;
  // Valid resolved marks, refetched every time the modal opens.
  var _clipMarksMarks = [];

  function _clipMarksScopedMarks() {
    var scope = (qs("#clipMarksScope") || {}).value;
    if (scope !== "current") return _clipMarksMarks;
    var pid = state.selectedParticipant;
    return _clipMarksMarks.filter(function (m) { return m.participant === pid; });
  }

  function _clipMarksNumber(sel, fallback, min, max) {
    var raw = parseFloat((qs(sel) || {}).value);
    if (isNaN(raw)) return fallback;
    return Math.min(max, Math.max(min, raw));
  }

  // The preview and the payload cluster through the same shared helper, so the
  // count the user reads is the number of clips they get.
  function _clipMarksClusters() {
    var marks = _clipMarksScopedMarks();
    if (!marks.length) return [];
    var gap = _clipMarksNumber("#clipMarksGap", CLIP_MARKS_DEFAULT_GAP_SECONDS, 0, 120);
    return window.ClipgenIntakeCluster.clusterTranscriptMarks(marks, gap);
  }

  function renderClipMarksSummary() {
    var summaryEl = qs("#clipMarksSummary");
    var confirmBtn = qs("#clipMarksConfirm");
    if (!summaryEl || !confirmBtn) return;
    if (_clipMarksRun) return; // progress block owns the copy while a run streams
    summaryEl.classList.remove("cg-shimmer"); // the "Loading marks…" fetch landed
    var marks = _clipMarksScopedMarks();
    if (!marks.length) {
      var pid = state.selectedParticipant;
      var scope = (qs("#clipMarksScope") || {}).value;
      summaryEl.textContent =
        scope === "current" && pid
          ? "No marked lines in " + pid + " yet."
          : "No marked lines yet — mark a line with M or the gutter dot.";
      confirmBtn.disabled = true;
      return;
    }
    var clusters = _clipMarksClusters();
    summaryEl.textContent =
      clipgenPluralUnit(marks.length, "marked line", "marked lines") +
      " → " +
      clipgenPluralUnit(clusters.length, "clip", "clips");
    confirmBtn.disabled = false;
  }

  // Both option labels carry live counts so the scope choice and its
  // consequence are legible in one place.
  function _renderClipMarksScopeOptions() {
    var sel = qs("#clipMarksScope");
    if (!sel || sel.options.length < 2) return;
    var pid = state.selectedParticipant;
    var mine = pid
      ? _clipMarksMarks.filter(function (m) { return m.participant === pid; }).length
      : 0;
    sel.options[0].textContent =
      (pid ? "Current participant (" + pid + ")" : "Current participant") +
      " — " + clipgenPluralUnit(mine, "mark", "marks");
    sel.options[0].disabled = !pid;
    sel.options[1].textContent =
      "All participants — " + clipgenPluralUnit(_clipMarksMarks.length, "mark", "marks");
    if (!pid) sel.value = "all";
  }

  function _renderClipMarksProgress() {
    var wrap = qs("#clipMarksProgress");
    var fill = qs("#clipMarksBarFill");
    var text = qs("#clipMarksProgressText");
    var confirmBtn = qs("#clipMarksConfirm");
    var cancelBtn = qs("#clipMarksCancel");
    if (!wrap || !fill || !text || !confirmBtn || !cancelBtn) return;
    var run = _clipMarksRun;
    wrap.classList.toggle("hidden", !run);
    confirmBtn.classList.toggle("hidden", !!run);
    cancelBtn.textContent = run ? "Stop" : "Cancel";
    if (!run) return;
    var pct = run.total ? Math.round((run.done / run.total) * 100) : 0;
    fill.style.width = pct + "%";
    text.textContent =
      "Clipping… " + run.done + "/" + run.total +
      (run.failed ? " (" + run.failed + " failed)" : "");
  }

  function openClipMarksModal() {
    var modal = qs("#clipMarksModal");
    if (!modal) return;
    modal.classList.remove("hidden");
    openBlockingModal(modal, {
      onEscape: closeClipMarksModal,
      onBackdropClick: closeClipMarksModal,
    });
    _renderClipMarksProgress();
    // A run in flight owns the dialog's copy; only refresh the pickers when idle.
    if (_clipMarksRun) return;
    qs("#clipMarksSummary").classList.add("cg-shimmer");
    qs("#clipMarksSummary").textContent = "Loading marks…";
    qs("#clipMarksConfirm").disabled = true;
    apiGet("api/marks")
      .then(function (data) {
        _clipMarksMarks = data.ok
          ? (data.marks || []).filter(function (m) { return m.valid; })
          : [];
        _renderClipMarksScopeOptions();
        renderClipMarksSummary();
      })
      .catch(function () {
        qs("#clipMarksSummary").classList.remove("cg-shimmer");
        qs("#clipMarksSummary").textContent = "Could not load marks.";
      });
  }

  function closeClipMarksModal() {
    var modal = qs("#clipMarksModal");
    if (!modal) return;
    closeBlockingModal(modal);
    modal.classList.add("hidden");
  }

  function submitClipMarks() {
    if (_clipMarksRun) return;
    var clusters = _clipMarksClusters();
    if (!clusters.length) return;
    var pad = _clipMarksNumber("#clipMarksPad", CLIP_MARKS_DEFAULT_PAD_SECONDS, 0, 10);
    // Only the start needs clamping — ffmpeg stops at EOF, and a multi-video
    // participant's span is bounded when it is mapped onto the timeline.
    var items = clusters.map(function (c) {
      return {
        participant: c.participant,
        start: Math.max(0, c.start - pad),
        end: c.end + pad,
        event_type: c.category || "transcript",
        event_ids: [],
        source: "transcript",
        mark_ids: c.marks.map(function (m) { return m.id; }),
        text: c.text || "",
        label: c.label || "",
      };
    });

    var run = { done: 0, failed: 0, total: items.length, abort: new AbortController() };
    _clipMarksRun = run;
    _renderClipMarksProgress();

    function handleLine(line) {
      var data;
      try { data = JSON.parse(line); } catch (_) { return; }
      // The trailing {"cancelled": true} line carries no index — this also
      // keeps it out of the completion tally.
      if (!data || typeof data.index !== "number") return;
      run.done++;
      if (!data.ok) run.failed++;
      _renderClipMarksProgress();
    }

    function finish(message) {
      _clipMarksRun = null;
      _renderClipMarksProgress();
      closeClipMarksModal();
      showToast(message);
    }

    apiPostNDJSON(
      "../studio/api/generate-intake",
      { items: items, format: "clip" },
      { signal: run.abort.signal, onLine: handleLine }
    )
      .then(function () {
        var made = run.done - run.failed;
        finish(
          run.failed
            ? clipgenPluralUnit(made, "clip", "clips") + " generated, " + run.failed + " failed"
            : clipgenPluralUnit(made, "clip", "clips") + " generated — open Studio to review"
        );
      })
      .catch(function (err) {
        var aborted = err && (err.name === "AbortError" || err.code === 20);
        finish(aborted ? "Clip generation cancelled" : "Clip generation failed: " + (err && err.message));
      });
  }

  function onClipMarksCancel() {
    if (!_clipMarksRun) {
      closeClipMarksModal();
      return;
    }
    apiPost("../studio/api/generate-intake/cancel", {}).catch(function () {});
    _clipMarksRun.abort.abort();
  }

  function initClipMarksModal() {
    qs("#clipMarksCancel").addEventListener("click", onClipMarksCancel);
    qs("#clipMarksConfirm").addEventListener("click", submitClipMarks);
    qs("#clipMarksScope").addEventListener("change", renderClipMarksSummary);
    qs("#clipMarksGap").addEventListener("input", renderClipMarksSummary);
  }

  // ---- Embed subtitles ----
  //
  // Muxes each participant's transcript back into a copy of their source video as
  // a soft subtitle track (stream copy, no re-encode). Structurally a twin of the
  // Clip Marked Lines dialog above — same scope picker, live summary and NDJSON
  // progress bar — with the track's default disposition as its only extra
  // parameter; container codec, track title and language are derived server-side.
  //
  // No fetch on open, unlike clip-marks: state.participants already carries
  // has_transcript and video_paths from the poll.

  // { done, failed, total, abort } while a batch streams; null when idle. Lives
  // out here so the dialog can be dismissed mid-run and reopened onto the live
  // progress, exactly like _clipMarksRun.
  var _embedSubsRun = null;

  // The server refuses a participant whose transcript spans several source
  // files (muxing it back would mean concatenating the parts first), and
  // video_paths already ships on /api/participants — so drop them here rather
  // than spending a whole ffmpeg round-trip to surface a failure line.
  function _embedSubsCandidates() {
    var ps = state.participants || [];
    var out = [];
    for (var i = 0; i < ps.length; i++) {
      if (ps[i].has_transcript) out.push(ps[i]);
    }
    return out;
  }

  function _embedSubsIsMultiPart(p) {
    return !!(p.video_paths && p.video_paths.length > 1);
  }

  function _embedSubsScoped() {
    var all = _embedSubsCandidates();
    if ((qs("#embedSubsScope") || {}).value !== "current") return all;
    var pid = state.selectedParticipant;
    return all.filter(function (p) { return p.id === pid; });
  }

  // Container extension of a participant's first source file, lowercased.
  function _embedSubsExt(p) {
    var path = (p.video_paths && p.video_paths[0]) || "";
    var dot = path.lastIndexOf(".");
    return dot === -1 ? "" : path.slice(dot).toLowerCase();
  }

  function _embedSubsContainers() {
    return CLIPGEN_CONFIG.subtitleContainers || { supported: [], alwaysDefault: [] };
  }

  // A container mux_subtitles has no codec for is rejected server-side at its
  // `codec is None` guard. Filtering here for the same reason multi-part is
  // filtered here: otherwise the summary promises "8 subtitled videos" and the
  // run comes back with 8 failure lines.
  function _embedSubsIsUnsupported(p) {
    var supported = _embedSubsContainers().supported || [];
    return supported.indexOf(_embedSubsExt(p)) === -1;
  }

  function _embedSubsTargets() {
    return _embedSubsScoped().filter(function (p) {
      return !_embedSubsIsMultiPart(p) && !_embedSubsIsUnsupported(p);
    });
  }

  // The mp4 family always ships its subtitle track enabled — measured on ffmpeg
  // 8.1.2, the muxer reports default=1 whatever -disposition:s:0 is given (an
  // ISOBMFF track is enabled or absent; only .mkv/.webm have a present-but-off
  // state). Unticking the box is therefore a no-op for those files, and a
  // control that silently does nothing is worse than one that says so.
  function _embedSubsAlwaysDefault(targets) {
    var always = _embedSubsContainers().alwaysDefault || [];
    return targets.filter(function (p) {
      return always.indexOf(_embedSubsExt(p)) !== -1;
    });
  }

  function renderEmbedSubsSummary() {
    var summaryEl = qs("#embedSubsSummary");
    var confirmBtn = qs("#embedSubsConfirm");
    if (!summaryEl || !confirmBtn) return;
    if (_embedSubsRun) return; // progress block owns the copy while a run streams
    var targets = _embedSubsTargets();
    var scoped = _embedSubsScoped();
    var skipped = scoped.filter(_embedSubsIsMultiPart);
    var unsupported = scoped.filter(function (p) {
      return !_embedSubsIsMultiPart(p) && _embedSubsIsUnsupported(p);
    });
    var skipNote = skipped.length
      ? " " + skipped.map(function (p) { return p.id; }).join(", ") +
        " skipped (multi-part " + (skipped.length === 1 ? "recording" : "recordings") + ")."
      : "";
    if (unsupported.length) {
      skipNote += " " + unsupported.map(function (p) { return p.id; }).join(", ") +
        " skipped (subtitles cannot be muxed into " +
        _embedSubsExt(unsupported[0]) + ").";
    }
    if (!targets.length) {
      var scope = (qs("#embedSubsScope") || {}).value;
      // Three different dead ends: nothing transcribed yet, transcripts that
      // exist but span several files, or a container the muxer cannot write.
      // Only the first is fixed by running a transcription, so saying that when
      // the real blocker is the footage sends the user at a no-op.
      if (unsupported.length && !skipped.length) {
        summaryEl.textContent =
          (unsupported.length === 1
            ? unsupported[0].id + "'s recording is " + _embedSubsExt(unsupported[0])
            : "These recordings are in a container") +
          ", which cannot carry an embedded subtitle track. Supported: " +
          (_embedSubsContainers().supported || []).join(", ") + ".";
      } else if (skipped.length) {
        summaryEl.textContent =
          (skipped.length === 1
            ? skipped[0].id + "'s transcript spans several video files"
            : "Every transcript here spans several video files") +
          ", which cannot be muxed back into one subtitled copy.";
      } else {
        summaryEl.textContent =
          scope === "current" && state.selectedParticipant
            ? "No transcript for " + state.selectedParticipant + " yet."
            : "No transcripts yet — transcribe a video first.";
      }
      confirmBtn.disabled = true;
      return;
    }
    // Only worth saying when the box is unticked: with it ticked the mp4
    // muxer's behaviour and the user's request agree.
    var stuckOn = (qs("#embedSubsDefault") || {}).checked
      ? []
      : _embedSubsAlwaysDefault(targets);
    var stuckNote = stuckOn.length
      ? " " + (stuckOn.length === targets.length ? "The track" : stuckOn.length + " of these")
        + " will still be on by default — .mp4/.mov cannot carry a subtitle track that is off."
      : "";
    summaryEl.textContent =
      clipgenPluralUnit(targets.length, "transcript", "transcripts") +
      " → " +
      clipgenPluralUnit(targets.length, "subtitled video", "subtitled videos") +
      "." + skipNote + stuckNote;
    confirmBtn.disabled = false;
  }

  // Both option labels carry live counts so the scope choice and its
  // consequence are legible in one place (mirrors the clip-marks picker).
  function _renderEmbedSubsScopeOptions() {
    var sel = qs("#embedSubsScope");
    if (!sel || sel.options.length < 2) return;
    var pid = state.selectedParticipant;
    var all = _embedSubsCandidates();
    var mine = pid
      ? all.filter(function (p) { return p.id === pid; }).length
      : 0;
    sel.options[0].textContent =
      (pid ? "Current participant (" + pid + ")" : "Current participant") +
      " — " + clipgenPluralUnit(mine, "transcript", "transcripts");
    sel.options[0].disabled = !mine;
    sel.options[1].textContent =
      "All participants — " + clipgenPluralUnit(all.length, "transcript", "transcripts");
    if (!mine) sel.value = "all";
  }

  function _renderEmbedSubsProgress() {
    var wrap = qs("#embedSubsProgress");
    var fill = qs("#embedSubsBarFill");
    var text = qs("#embedSubsProgressText");
    var confirmBtn = qs("#embedSubsConfirm");
    var cancelBtn = qs("#embedSubsCancel");
    if (!wrap || !fill || !text || !confirmBtn || !cancelBtn) return;
    var run = _embedSubsRun;
    wrap.classList.toggle("hidden", !run);
    confirmBtn.classList.toggle("hidden", !!run);
    cancelBtn.textContent = run ? "Stop" : "Cancel";
    if (!run) return;
    var pct = run.total ? Math.round((run.done / run.total) * 100) : 0;
    fill.style.width = pct + "%";
    text.textContent =
      "Embedding… " + run.done + "/" + run.total +
      (run.failed ? " (" + run.failed + " failed)" : "");
  }

  function openEmbedSubsModal() {
    var modal = qs("#embedSubsModal");
    if (!modal) return;
    modal.classList.remove("hidden");
    openBlockingModal(modal, {
      onEscape: closeEmbedSubsModal,
      onBackdropClick: closeEmbedSubsModal,
    });
    _renderEmbedSubsProgress();
    // A run in flight owns the dialog's copy; only refresh the pickers when idle.
    if (_embedSubsRun) return;
    _renderEmbedSubsScopeOptions();
    renderEmbedSubsSummary();
  }

  function closeEmbedSubsModal() {
    var modal = qs("#embedSubsModal");
    if (!modal) return;
    closeBlockingModal(modal);
    modal.classList.add("hidden");
  }

  function submitEmbedSubs() {
    if (_embedSubsRun) return;
    var targets = _embedSubsTargets();
    if (!targets.length) return;
    var pids = targets.map(function (p) { return p.id; });
    var defaultTrack = !!(qs("#embedSubsDefault") || {}).checked;

    var run = { done: 0, failed: 0, total: pids.length, abort: new AbortController() };
    _embedSubsRun = run;
    _renderEmbedSubsProgress();
    var outputDir = "";

    var sawDone = false;

    function handleLine(line) {
      var data;
      try { data = JSON.parse(line); } catch (_) { return; }
      if (!data) return;
      // The header line carries the destination and the trailing
      // {"cancelled": true} / {"done": true} carry nothing — none has an
      // index, which is also what keeps them out of the completion tally.
      if (data.output_dir) outputDir = data.output_dir;
      if (data.token) run.token = data.token; // echoed by Stop to scope the cancel
      if (data.done) sawDone = true;
      if (typeof data.index !== "number") return;
      run.done++;
      if (!data.ok) run.failed++;
      _renderEmbedSubsProgress();
    }

    function finish(message) {
      _embedSubsRun = null;
      _renderEmbedSubsProgress();
      closeEmbedSubsModal();
      showToast(message);
    }

    apiPostNDJSON(
      "api/embed-subtitles",
      { participants: pids, default_track: defaultTrack },
      { signal: run.abort.signal, onLine: handleLine }
    )
      .then(function () {
        var made = run.done - run.failed;
        var where = outputDir ? " to " + outputDir : "";
        // No sentinel means the server died partway and the body simply
        // stopped. The stream reader cannot tell that from a clean end, so
        // without this check a run that blew up at participant 4 of 10
        // reported "3 subtitled videos written" and nothing else.
        if (!sawDone) {
          finish(
            "Subtitle embedding stopped early — " +
              clipgenPluralUnit(made, "video was", "videos were") + " written" + where +
              " of " + run.total + ". Check the clipgen log."
          );
          return;
        }
        finish(
          run.failed
            ? clipgenPluralUnit(made, "subtitled video", "subtitled videos") + " written" + where + ", " + run.failed + " failed"
            : clipgenPluralUnit(made, "subtitled video", "subtitled videos") + " written" + where
        );
      })
      .catch(function (err) {
        var aborted = err && (err.name === "AbortError" || err.code === 20);
        finish(aborted ? "Subtitle embedding cancelled" : "Subtitle embedding failed: " + (err && err.message));
      });
  }

  function onEmbedSubsCancel() {
    if (!_embedSubsRun) {
      closeEmbedSubsModal();
      return;
    }
    // The server stops between files (a mux cannot be interrupted mid-copy), so
    // the abort just drops our end of a stream that is already winding down.
    apiPost("api/embed-subtitles/cancel", { token: _embedSubsRun.token || null }).catch(function () {});
    _embedSubsRun.abort.abort();
  }

  function initEmbedSubsModal() {
    qs("#embedSubsCancel").addEventListener("click", onEmbedSubsCancel);
    qs("#embedSubsConfirm").addEventListener("click", submitEmbedSubs);
    qs("#embedSubsScope").addEventListener("change", renderEmbedSubsSummary);
    // The checkbox does not change the file count, but it does decide whether
    // the .mp4 caveat applies — so the summary has to re-render for it.
    qs("#embedSubsDefault").addEventListener("change", renderEmbedSubsSummary);
  }

  // ---- Normalize audio ----
  //
  // Rewrites source videos in place with loudness-normalized audio: the picture
  // (and any unselected track) is stream-copied, the chosen tracks are
  // re-encoded through the server's loudnorm preset, and the original is parked
  // beside the source as .orig — the same kept-original flow the remux banner
  // manages, so Restore/Delete come for free after the run. Structurally a twin
  // of the Embed Subtitles dialog above; its one extra wrinkle is the track
  // field, which swaps between a coarse mode select (scope: all — per-participant
  // track layouts are heterogeneous, so explicit indices would be ambiguous) and
  // per-track checkboxes fetched from api/audio-info (scope: current).

  // { done, failed, total, abort } while a batch streams; null when idle. Lives
  // out here so the dialog can be dismissed mid-run and reopened onto the live
  // progress, exactly like _embedSubsRun.
  var _normAudioRun = null;

  // pid -> how many of the participant's parts have a kept .orig occupying
  // their backup slot (from api/remux/status on open). Only a *fully* occupied
  // participant is excluded: the server counts already-rewritten parts as done,
  // so a run that failed or was stopped halfway through a multi-part
  // participant resumes from the parts that are left rather than being locked
  // out until the successful backups are deleted or restored.
  var _normAudioKept = {};

  // The current-scope track checkboxes are built from an async audio-info
  // fetch; these pin which participant the rendered list (and its indices)
  // belong to, so a stale response can never dress the wrong participant.
  var _normAudioTrackPid = null;
  var _normAudioTrackInfo = null;

  function _normAudioCandidates() {
    var ps = state.participants || [];
    var out = [];
    for (var i = 0; i < ps.length; i++) {
      // A transcript is not required — normalization reads only the media.
      if (ps[i].has_video) out.push(ps[i]);
    }
    return out;
  }

  function _normAudioScoped() {
    var all = _normAudioCandidates();
    if ((qs("#normAudioScope") || {}).value !== "current") return all;
    var pid = state.selectedParticipant;
    return all.filter(function (p) { return p.id === pid; });
  }

  function _normAudioKeptCount(p) {
    return _normAudioKept[p.id] || 0;
  }

  function _normAudioIsFullyKept(p) {
    var parts = (p.video_paths || []).length || 1;
    return _normAudioKeptCount(p) >= parts;
  }

  function _normAudioTargets() {
    return _normAudioScoped().filter(function (p) {
      return !_normAudioIsFullyKept(p);
    });
  }

  // Files a run would actually rewrite: parts whose backup slot is free.
  function _normAudioFileCount(targets) {
    var n = 0;
    for (var i = 0; i < targets.length; i++) {
      var parts = (targets[i].video_paths || []).length || 1;
      n += Math.max(0, parts - _normAudioKeptCount(targets[i]));
    }
    return n;
  }

  // Which tracks the submit should send: the mode select's word for scope=all,
  // the checked checkbox indices for a multi-track current participant, and
  // "auto" when there is nothing to choose (single track).
  function _normAudioTracksSpec() {
    if ((qs("#normAudioScope") || {}).value !== "current") {
      return (qs("#normAudioTrackMode") || {}).value || "auto";
    }
    var info = _normAudioTrackInfo;
    if (!info || info.count <= 1) return "auto";
    var boxes = document.querySelectorAll("#normAudioTrackList input");
    var picked = [];
    for (var i = 0; i < boxes.length; i++) {
      if (boxes[i].checked) picked.push(parseInt(boxes[i].value, 10));
    }
    return picked;
  }

  function renderNormAudioSummary() {
    var summaryEl = qs("#normAudioSummary");
    var confirmBtn = qs("#normAudioConfirm");
    if (!summaryEl || !confirmBtn) return;
    if (_normAudioRun) return; // progress block owns the copy while a run streams
    var scope = (qs("#normAudioScope") || {}).value;
    var scoped = _normAudioScoped();
    var targets = _normAudioTargets();
    var fullyKept = scoped.filter(_normAudioIsFullyKept);
    var resuming = targets.filter(function (p) { return _normAudioKeptCount(p) > 0; });
    var keptNote = fullyKept.length
      ? " " + fullyKept.map(function (p) { return p.id; }).join(", ") +
        " skipped (an earlier original is still kept — delete or restore it first)."
      : "";
    if (resuming.length) {
      // Partially-rewritten multi-part participants: the server skips the
      // parts whose backup slot is occupied, so the run finishes what's left.
      keptNote += " " + resuming.map(function (p) { return p.id; }).join(", ") +
        (resuming.length === 1
          ? " resumes where it stopped"
          : " resume where they stopped") +
        " — already-rewritten parts are skipped.";
    }
    if (!targets.length) {
      summaryEl.textContent = fullyKept.length
        ? "Nothing to normalize —" + keptNote
        : scope === "current" && state.selectedParticipant
          ? "No source video for " + state.selectedParticipant + "."
          : "No source videos yet.";
      confirmBtn.disabled = true;
      return;
    }
    if (scope === "current" && _normAudioTrackPid && !_normAudioTrackInfo) {
      summaryEl.textContent = "Checking audio tracks…";
      confirmBtn.disabled = true;
      return;
    }
    var spec = _normAudioTracksSpec();
    if (Object.prototype.toString.call(spec) === "[object Array]" && !spec.length) {
      summaryEl.textContent = "Select at least one track to normalize." + keptNote;
      confirmBtn.disabled = true;
      return;
    }
    var files = _normAudioFileCount(targets);
    var trackNote =
      scope === "current" && _normAudioTrackInfo && _normAudioTrackInfo.count === 1
        ? " (1 audio track)"
        : "";
    summaryEl.textContent =
      clipgenPluralUnit(files, "video", "videos") + trackNote +
      " → rewritten in place; " +
      (files === 1 ? "the original is" : "originals are") +
      " kept beside " + (files === 1 ? "it" : "them") + "." + keptNote;
    confirmBtn.disabled = false;
  }

  // Both option labels carry live counts so the scope choice and its
  // consequence are legible in one place (mirrors the embed picker).
  function _renderNormAudioScopeOptions() {
    var sel = qs("#normAudioScope");
    if (!sel || sel.options.length < 2) return;
    var pid = state.selectedParticipant;
    var all = _normAudioCandidates();
    var mine = pid
      ? all.filter(function (p) { return p.id === pid; }).length
      : 0;
    sel.options[0].textContent =
      (pid ? "Current participant (" + pid + ")" : "Current participant") +
      " — " + clipgenPluralUnit(mine, "video", "videos");
    sel.options[0].disabled = !mine;
    sel.options[1].textContent =
      "All participants — " + clipgenPluralUnit(all.length, "video", "videos");
    if (!mine) sel.value = "all";
  }

  // Swap the track field between the coarse mode select (scope: all) and the
  // per-track checkboxes (scope: current, multi-track file).
  function _renderNormAudioTrackField() {
    var modeLabel = qs("#normAudioTrackModeLabel");
    var list = qs("#normAudioTrackList");
    if (!modeLabel || !list) return;
    var scope = (qs("#normAudioScope") || {}).value;
    var pid = scope === "current" ? state.selectedParticipant : null;
    var p = pid
      ? _normAudioCandidates().filter(function (c) { return c.id === pid; })[0]
      : null;
    if (!p) {
      _normAudioTrackPid = null;
      _normAudioTrackInfo = null;
      modeLabel.classList.toggle("hidden", scope === "current");
      list.classList.add("hidden");
      list.innerHTML = "";
      renderNormAudioSummary();
      return;
    }
    modeLabel.classList.add("hidden");
    list.classList.add("hidden");
    list.innerHTML = "";
    _normAudioTrackPid = pid;
    _normAudioTrackInfo = null;
    renderNormAudioSummary(); // "Checking audio tracks…" while the probe runs
    _trFetchAudioInfo(pid, p.video_version).then(function (info) {
      // A scope flip or participant change while the probe ran owns the field.
      if (_normAudioTrackPid !== pid) return;
      _normAudioTrackInfo = info || { tracks: [], count: 1, auto: 0 };
      if (_normAudioTrackInfo.count > 1) {
        for (var i = 0; i < _normAudioTrackInfo.tracks.length; i++) {
          var row = document.createElement("label");
          row.className = "param-modal-label";
          var text = document.createElement("span");
          // Late-bound: transcripts-pills.js loads after the hub and publishes
          // the label helper on the namespace.
          text.textContent = TS.trackOptionLabel
            ? TS.trackOptionLabel(_normAudioTrackInfo.tracks[i], i)
            : "Track " + (i + 1);
          var box = document.createElement("input");
          box.type = "checkbox";
          box.className = "param-modal-checkbox";
          box.value = String(i);
          box.checked = i === _normAudioTrackInfo.auto;
          row.appendChild(text);
          row.appendChild(box);
          list.appendChild(row);
        }
        list.classList.remove("hidden");
      }
      renderNormAudioSummary();
    });
  }

  function _renderNormAudioProgress() {
    var wrap = qs("#normAudioProgress");
    var fill = qs("#normAudioBarFill");
    var text = qs("#normAudioProgressText");
    var confirmBtn = qs("#normAudioConfirm");
    var cancelBtn = qs("#normAudioCancel");
    if (!wrap || !fill || !text || !confirmBtn || !cancelBtn) return;
    var run = _normAudioRun;
    wrap.classList.toggle("hidden", !run);
    confirmBtn.classList.toggle("hidden", !!run);
    cancelBtn.textContent = run ? "Stop" : "Cancel";
    if (!run) return;
    var pct = run.total ? Math.round((run.done / run.total) * 100) : 0;
    fill.style.width = pct + "%";
    text.textContent =
      "Normalizing… " + run.done + "/" + run.total +
      (run.failed ? " (" + run.failed + " failed)" : "");
  }

  function openNormalizeAudioModal() {
    var modal = qs("#normAudioModal");
    if (!modal) return;
    modal.classList.remove("hidden");
    openBlockingModal(modal, {
      onEscape: closeNormalizeAudioModal,
      onBackdropClick: closeNormalizeAudioModal,
    });
    _renderNormAudioProgress();
    // A run in flight owns the dialog's copy; only refresh the pickers when idle.
    if (_normAudioRun) return;
    _normAudioKept = {};
    _renderNormAudioScopeOptions();
    _renderNormAudioTrackField();
    // Kept-original state lives on disk, not in state.participants — the remux
    // status endpoint re-probes it on every call, so one fetch on open is
    // always current.
    apiGet("api/remux/status")
      .then(function (data) {
        if (!data || !data.ok || _normAudioRun) return;
        var kept = data.kept || {};
        var map = {};
        for (var pid in kept) {
          if (kept[pid] && kept[pid].length) map[pid] = kept[pid].length;
        }
        _normAudioKept = map;
        renderNormAudioSummary();
      })
      .catch(function () {});
  }

  function closeNormalizeAudioModal() {
    var modal = qs("#normAudioModal");
    if (!modal) return;
    closeBlockingModal(modal);
    modal.classList.add("hidden");
  }

  function submitNormalizeAudio() {
    if (_normAudioRun) return;
    var targets = _normAudioTargets();
    if (!targets.length) return;
    var pids = targets.map(function (p) { return p.id; });
    var spec = _normAudioTracksSpec();

    var run = {
      done: 0,
      failed: 0,
      changed: 0,
      total: pids.length,
      abort: new AbortController(),
    };
    _normAudioRun = run;
    _renderNormAudioProgress();

    var sawDone = false;

    function handleLine(line) {
      var data;
      try { data = JSON.parse(line); } catch (_) { return; }
      if (!data) return;
      // The header and the trailing {"cancelled": true} / {"done": true} lines
      // carry no index, which is also what keeps them out of the tally.
      if (data.token) run.token = data.token; // echoed by Stop to scope the cancel
      if (data.done) sawDone = true;
      if (typeof data.index !== "number") return;
      run.done++;
      if (!data.ok) run.failed++;
      // Files actually swapped — can be non-zero on an ok=false line when a
      // later part of a multi-part participant failed or was stopped.
      if (typeof data.parts_done === "number") run.changed += data.parts_done;
      _renderNormAudioProgress();
    }

    function finish(message) {
      _normAudioRun = null;
      _renderNormAudioProgress();
      closeNormalizeAudioModal();
      showToast(message);
      // Any swapped file was pulled out from under the page: the <video> is
      // mid-stream on a renamed-away inode and the per-track mixers point at
      // stale extracts. Keyed on parts swapped (run.changed), not successful
      // participants — a participant that failed on part 2 still replaced
      // part 1. media-banner.js reloads for the identical file swap
      // (reloadAfterFileSwap); the delay lets the batch toast — which, unlike
      // remux's single-file one, carries failure counts — be seen first.
      // Playback position survives via videoTimeByParticipant.
      if (run.changed > 0) {
        setTimeout(function () { window.location.reload(); }, 1500);
      }
    }

    apiPostNDJSON(
      "api/normalize-audio",
      { participants: pids, tracks: spec },
      { signal: run.abort.signal, onLine: handleLine }
    )
      .then(function () {
        var made = run.done - run.failed;
        // No sentinel means the server died partway and the body simply
        // stopped; without this check a run that blew up at participant 4 of
        // 10 would report 3 rewrites as a clean finish.
        if (!sawDone) {
          finish(
            "Audio normalization stopped early — " +
              clipgenPluralUnit(made, "video was", "videos were") + " rewritten of " +
              run.total + ". Check the clipgen log."
          );
          return;
        }
        finish(
          run.failed
            ? clipgenPluralUnit(made, "video", "videos") + " normalized, " + run.failed + " failed"
            : clipgenPluralUnit(made, "video", "videos") + " normalized; originals kept beside the sources"
        );
      })
      .catch(function (err) {
        var aborted = err && (err.name === "AbortError" || err.code === 20);
        finish(aborted ? "Audio normalization cancelled" : "Audio normalization failed: " + (err && err.message));
      });
  }

  function onNormalizeAudioCancel() {
    if (!_normAudioRun) {
      closeNormalizeAudioModal();
      return;
    }
    // Unlike the embed run, the cancel event is threaded into ffmpeg itself, so
    // Stop interrupts the current file mid-encode; the abort just drops our end
    // of a stream that is already winding down.
    apiPost("api/normalize-audio/cancel", { token: _normAudioRun.token || null }).catch(function () {});
    _normAudioRun.abort.abort();
  }

  function initNormalizeAudioModal() {
    qs("#normAudioCancel").addEventListener("click", onNormalizeAudioCancel);
    qs("#normAudioConfirm").addEventListener("click", submitNormalizeAudio);
    // Scope decides which track control is shown, so it re-renders the field
    // (which re-renders the summary itself once the async probe settles).
    qs("#normAudioScope").addEventListener("change", _renderNormAudioTrackField);
    qs("#normAudioTrackMode").addEventListener("change", renderNormAudioSummary);
    // Delegated: the checkbox rows are rebuilt per participant.
    qs("#normAudioTrackList").addEventListener("change", renderNormAudioSummary);
  }

  // ---- Boot ----

  // Participants the Transcribe All action would enqueue: a source video, no
  // transcript yet, and nothing already queued or running for them. That last
  // filter is not optional — /api/transcribe has no in-flight guard, so without
  // it a second click while the batch runs enqueues every pending pid twice.
  function _untranscribedParticipants() {
    var ps = state.participants || [];
    var tasks = state.tasks || [];
    var busy = {};
    for (var t = 0; t < tasks.length; t++) {
      if (tasks[t].status === "queued" || tasks[t].status === "running") busy[tasks[t].participant] = true;
    }
    var pids = [];
    for (var i = 0; i < ps.length; i++) {
      if (ps[i].has_video && !ps[i].has_transcript && !busy[ps[i].id]) pids.push(ps[i].id);
    }
    return pids;
  }

  // One POST for the whole list; the worker thread runs them sequentially and the
  // pills render each task's progress off the existing poller. _postTranscribe
  // owns the toast, the polling restart, and the uncached-model confirm.
  function runTranscribeAll() {
    var pids = _untranscribedParticipants();
    if (!pids.length) return;
    transcribeParticipants(pids, false);
  }

  var _rebuildTopNavActions = function () {};

  // Published on TS for transcripts-agents.js, which gates its panel-visible
  // refreshes on it. The topnav no longer reads it — the Embed Subtitles dialog
  // does its own scoping — but the satellite still does.
  function _currentParticipantHasTranscript() {
    var pid = state.selectedParticipant;
    if (!pid || !state.participants) return false;
    for (var i = 0; i < state.participants.length; i++) {
      if (state.participants[i].id === pid) return !!state.participants[i].has_transcript;
    }
    return false;
  }

  function refreshTopNavActions() {
    _rebuildTopNavActions();
  }

  function initTopNavActions() {
    if (!window.ClipgenTopNav) return;
    function rebuild() {
      var pending = _untranscribedParticipants().length;
      window.ClipgenTopNav.setQuickActions([
        {
          icon: "microphone",
          label: "Transcribe All",
          action: runTranscribeAll,
          disabled: pending === 0,
          title: pending
            ? "Queue transcription for the " + clipgenPluralUnit(pending, "participant", "participants") + " without a transcript"
            : "Every participant with a video already has a transcript or is queued.",
        },
        {
          icon: "language",
          label: "Embed Subtitles…",
          action: openEmbedSubsModal,
          // Never gated, same as Clip Marked Lines below: the modal reports the
          // transcript count and disables its own Embed button, so the scope
          // choice lives where its consequence is visible.
          title: "Write a subtitled copy of each source video into the output folder; the originals are never modified",
        },
        {
          icon: "speaker-wave",
          label: "Normalize Audio…",
          action: openNormalizeAudioModal,
          // Never gated, same as the neighbours: the modal reports the video
          // count and disables its own Normalize button.
          title: "Rewrite source videos in place with loudness-normalized audio; the original is kept beside each file until you delete it",
        },
        {
          icon: "scissors",
          label: "Clip Marked Lines…",
          action: openClipMarksModal,
          // Never gated: the modal reports the mark count and disables its own
          // Generate button, which keeps the menu free of an async mark fetch.
          title: "Cut a clip for every manually marked line",
        },
        window.ClipgenExportActions.exportQuickAction(),
      ]);
    }
    _rebuildTopNavActions = rebuild;
    rebuild();
    window.ClipgenExportActions.refreshExportStatus(rebuild);
    // Always rebuild on menu open so Transcribe All re-counts against in-flight
    // tasks; also refresh export-enabled state.
    window.ClipgenTopNav.onBeforeOpen(function () {
      rebuild();
      window.ClipgenExportActions.refreshExportStatus(rebuild);
    });
  }

  // Command palette (command-palette.js): additions beyond the auto-ingested
  // quick actions — search focus, cheatsheet, panel tabs, participant jumps
  // (the provider runs on every palette open, so it tracks state.participants).
  function initCommandPalette() {
    if (!window.ClipgenCommandPalette) return;
    function clickCommand(id, title, icon, keywords, elId) {
      return {
        id: id,
        title: title,
        icon: icon,
        keywords: keywords,
        section: "Transcripts",
        visible: function () { return !!document.getElementById(elId); },
        run: function () { document.getElementById(elId).click(); },
      };
    }
    window.ClipgenCommandPalette.setParticipants(function () {
      return (state.participants || []).map(function (p) { return p.id; });
    });
    window.ClipgenCommandPalette.register("transcripts", function () {
      var cmds = [
        {
          id: "transcripts:search",
          title: "Focus transcript search",
          icon: "magnifying-glass",
          keywords: "find text query",
          section: "Transcripts",
          visible: function () { return !!document.getElementById("searchInput"); },
          // Runs after the palette closes and restores focus, so this focus
          // call wins.
          run: function () { document.getElementById("searchInput").focus(); },
        },
        clickCommand("transcripts:shortcuts", "Keyboard shortcuts", "command-line",
          "cheatsheet keys help", "shortcutsBtn"),
        clickCommand("transcripts:tab-summary", "Show Summary tab", "table-cells",
          "panel agents", "tabBtnSummary"),
        clickCommand("transcripts:tab-friction", "Show Friction tab", "table-cells",
          "panel analysis moments", "tabBtnFriction"),
        clickCommand("transcripts:toggle-video", "Toggle video panel", "video-camera",
          "hide show collapse drawer player", "videoCollapseBtn"),
        clickCommand("transcripts:toggle-captions", "Toggle captions", "language",
          "subtitles cc text", "videoCcBtn"),
        {
          id: "transcripts:cycle-friction-mode",
          title: "Cycle friction mode (off / highlight / isolate)",
          icon: "fire",
          keywords: "friction analysis overlay heatmap filter isolate highlight timeline",
          section: "Transcripts",
          // Gated on data, not on a control: the mode track only means anything
          // once the selected participant has friction scores.
          enabled: function () {
            var fd = state.frictionData;
            return !!(fd && fd.segments && fd.segments.length);
          },
          run: function () { cycleFrictionMode(); },
        },
      ];
      // "Jump to … in Transcripts" = stays here and selects in place; the
      // palette's built-in provider adds the cross-page "Open … in <Page>".
      (state.participants || []).forEach(function (p) {
        cmds.push({
          id: "transcripts:p:" + p.id,
          title: "Jump to " + p.id + " in Transcripts",
          icon: "user",
          keywords: "participant select transcript",
          section: "Participants",
          run: function () { selectParticipant(p.id); },
        });
      });
      return cmds;
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    initThemeToggle();
    initStatusIndicatorTooltip();
    checkNavLinks();
    initFrontendSwitcher();
    initSearch();
    initPillOutsideClick();
    initPillWheelScroll();
    initCorrectionsModal();
    initClipMarksModal();
    initEmbedSubsModal();
    initNormalizeAudioModal();
    initVideoPlayer();
    initVideoSync();
    initTimelineCanvas();
    initPipScroll();
    initPlayerKeyboard();
    initPanelTabs();
    // /transcripts/#tab=friction style deep links (command palette). The
    // participant hash form (#P07) is handled separately in loadParticipants.
    var hashTab = clipgenHashTab();
    if (hashTab === "summary" || hashTab === "friction") {
      var hashTabBtn = qs(hashTab === "friction" ? "#tabBtnFriction" : "#tabBtnSummary");
      if (hashTabBtn) hashTabBtn.click();
    }
    initSummaryActions();
    initFriction();
    initFrictionMode();
    initTranscriptSettings();
    initTopNavActions();
    initCommandPalette();

    // Pause every poller when tab is hidden; resume what was active on focus.
    // Without this, summary/citations/model-hint pollers (1.5–3 s cadence)
    // keep hammering the backend from background tabs.
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) {
        stopPolling();
        stopXrefPolling();
        stopModelHintPoll();
        _stopSummaryPoll();
        _stopCitationsPoll();
        _stopFrictionPoll();
        _txEtaTicker.stop();
      } else {
        pollTaskStatus();
        startXrefPolling();
        if (_anyTxEtaActive()) _txEtaTicker.ensure();
        // Re-check summary + citations on tab refocus. Background-running
        // Ollama agents finish without notifying the frontend; if the citations
        // poll already gave up (or summary completed after we stopped polling)
        // the manifest result would only surface on a full page reload. This
        // catches the common "user goes to another tab, comes back" case.
        // loadSummary re-arms the summary/citations polls if generation is
        // still in flight, and the transcribe-status poll re-arms the
        // model-hint poll on the next active task transition. loadFriction does
        // the same for the friction pass.
        if (state.selectedParticipant) {
          loadSummary(state.selectedParticipant);
          loadFriction(state.selectedParticipant);
        }
      }
    });

    // Window focus is a separate signal from tab visibility: switching to
    // another window/app (Cmd-Tab) leaves the tab "visible" (document.hidden
    // stays false, so the visibilitychange handler above never fires) yet
    // browsers — Safari most aggressively — still pause/throttle setInterval for
    // the unfocused window. That freezes the streaming "Transcribing… X%"
    // progress until the next throttled tick or a manual reload. Re-poll
    // immediately on focus so the transcription progress resyncs without a
    // reload (the per-agent summary/friction pollers self-resume on focus).
    window.addEventListener("focus", function () {
      if (document.hidden) return;
      pollTaskStatus();
      if (_anyTxEtaActive()) _txEtaTicker.ensure();
    });

    // Load initial data
    loadParticipants();

    // Fetch live mark categories so the popover/pill renders match overrides.
    apiGet("api/marks").then(function (data) {
      if (data && data.ok && data.categories) {
        setMarkCategories(data.categories);
        if (!MARK_CATEGORIES[state.lastMarkCategory]) {
          var firstKey = Object.keys(MARK_CATEGORIES)[0];
          state.lastMarkCategory = firstKey || "bookmark";
        }
      }
    });

    // Check for active tasks on load
    pollTaskStatus();
  });

  // ---- Satellite interface (window.ClipgenTranscripts) ----
  // Published for the transcripts-*.js satellite files (corrections, search,
  // video, pills, agents) that load after this script. They read the hub's
  // shared `state` + helpers through this object and attach their own published
  // functions back onto it — mirrors screenspace.js / window.ClipgenScreenspace.
  // Assigned synchronously here (during the hub script's load) so the object is
  // fully populated before any satellite IIFE runs; the DOMContentLoaded init
  // above and all user-event handlers fire later still, by which point
  // satellites have registered their functions.
  //
  // Hub helpers the satellites call outward are published below as they are
  // needed by each carved satellite. Functions the hub calls that live in a
  // satellite are reached through thin guarded delegators (see the delegators
  // section) so the hub degrades to a no-op if a satellite fails to load.
  var TS = (window.ClipgenTranscripts = window.ClipgenTranscripts || {});
  TS.state = state;
  TS.showToast = showToast;
  // Hub helpers the satellites call outward.
  TS.loadTranscript = loadTranscript; // corrections, search, agents
  TS.findOverlapsForSearch = findOverlapsForSearch; // search
  TS.selectParticipant = selectParticipant; // search, pills
  TS.cycleParticipant = cycleParticipant; // video (Z/X participant cycle)
  // Accumulated streaming segments for the currently-streamed participant. The
  // status poll no longer carries partial_segments, so search reads them here.
  TS.streamingSegmentsFor = function (pid) {
    return pid && pid === state.streamingParticipant ? _streamSegs.segments : [];
  }; // search
  TS.getMarkForSegment = getMarkForSegment; // video (timeline markers + tooltip)
  TS.toggleMark = toggleMark; // video (keyboard marking)
  TS.updateMarkCategory = updateMarkCategory; // video (keyboard category set)
  TS.showMarkPopover = showMarkPopover; // video (keyboard marking on already-marked segment)
  TS.maybeWarmOnPillHover = maybeWarmOnPillHover; // pills
  TS.tryPostTranscriptionWarmup = tryPostTranscriptionWarmup; // pills
  TS.pollTaskStatus = pollTaskStatus; // pills
  TS.refreshTranscribeWording = refreshTranscribeWording; // pills (cancel repaints before the next poll)
  TS.startPolling = startPolling; // pills
  TS._refreshAgentStateNow = _refreshAgentStateNow; // pills
  TS._trFetchModels = _trFetchModels; // pills
  TS._trFetchAudioInfo = _trFetchAudioInfo; // pills (audio-track picker)
  TS.audioInfoCached = audioInfoCached; // pills (synchronous warm-cache peek)
  TS.ensureAgentModelInstalled = ensureAgentModelInstalled; // pills, agents
  TS.confirmModelInstall = confirmModelInstall; // shot.py state probing; future satellites
  TS._confirmUncachedWhisperModels = _confirmUncachedWhisperModels; // pills (model-install kept in hub)
  // Hub helpers the agents satellite calls outward (loadFriction is owned by the
  // agents satellite now and reached through the delegator above).
  TS.renderSegments = renderSegments; // agents (heatmap toggle, friction mark-all)
  TS._txEtaTicker = _txEtaTicker; // agents (summary/citations/friction elapsed)
  TS._summaryEtaTracker = _summaryEtaTracker; // agents
  TS._citationsEtaTracker = _citationsEtaTracker; // agents
  TS._frictionEtaTracker = _frictionEtaTracker; // agents
  TS._updateAgentElapsed = _updateAgentElapsed; // agents
  TS._currentParticipantHasTranscript = _currentParticipantHasTranscript; // agents (panel-visible guard)

})();
