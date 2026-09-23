/* clipgen Transcripts batch-job satellite — transcripts-batch.js
 *
 * The three parameter-modal quick actions that run one NDJSON batch over the
 * study: Clip Marked Lines (clusters marks with intake-cluster.js and streams
 * them to Studio's generate-intake), Embed Subtitles (soft-muxes transcripts
 * into subtitled copies) and Normalize Audio (loudnorm in place, .orig kept).
 * One createBatchJobModal factory owns the shared skeleton — scope picker with
 * live counts, progress bar, cancel-with-token, run-outlives-dialog — and each
 * flow supplies named candidate / summary / request / toast functions. Loads
 * last; reads hub state + helpers through window.ClipgenTranscripts (TS) and
 * publishes the three open functions plus initBatchModals for the hub's quick
 * actions and boot. TS.trackOptionLabel (pills) is late-bound. Plain utils.js
 * globals (qs/apiGet/apiPost/apiPostNDJSON/clipgenPluralUnit/openBlockingModal)
 * are reached via the scope chain.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var showToast = TS.showToast,
    _trFetchAudioInfo = TS._trFetchAudioInfo;

  // ---- Batch-job modals ----
  // One NDJSON run per dialog; it outlives a dismissed dialog.
  function createBatchJobModal(cfg) {
    var run = null; // { done, failed, total, abort, token? } while streaming, null idle
    var api;

    function el(suffix) {
      return qs("#" + cfg.prefix + suffix);
    }
    function scopeValue() {
      return (el("Scope") || {}).value;
    }
    function ownedBy(all, pid) {
      return all.filter(function (item) { return cfg.itemPid(item) === pid; });
    }
    function scoped() {
      var all = cfg.candidates();
      if (scopeValue() !== "current") return all;
      return ownedBy(all, state.selectedParticipant);
    }
    function targets() {
      var items = scoped();
      return cfg.targets ? cfg.targets(items) : items;
    }

    function renderSummary() {
      var summaryEl = el("Summary");
      var confirmBtn = el("Confirm");
      if (!summaryEl || !confirmBtn) return;
      if (run) return; // progress block owns the copy while a run streams
      summaryEl.classList.remove("cg-shimmer");
      var result = cfg.summary({
        scoped: scoped(),
        targets: targets(),
        scope: scopeValue(),
        pid: state.selectedParticipant,
      });
      summaryEl.textContent = result.text;
      confirmBtn.disabled = !result.ready;
    }

    // Option labels carry live counts.
    function renderScopeOptions() {
      var sel = el("Scope");
      if (!sel || sel.options.length < 2) return;
      var pid = state.selectedParticipant;
      var all = cfg.candidates();
      var count = pid ? ownedBy(all, pid).length : 0;
      var enabled = cfg.currentEnabled ? cfg.currentEnabled(pid, count) : count > 0;
      sel.options[0].textContent =
        (pid ? "Current participant (" + pid + ")" : "Current participant") +
        " — " + clipgenPluralUnit(count, cfg.unit[0], cfg.unit[1]);
      sel.options[0].disabled = !enabled;
      sel.options[1].textContent =
        "All participants — " + clipgenPluralUnit(all.length, cfg.unit[0], cfg.unit[1]);
      if (!enabled) sel.value = "all";
    }

    function renderProgress() {
      var wrap = el("Progress");
      var fill = el("BarFill");
      var text = el("ProgressText");
      var confirmBtn = el("Confirm");
      var cancelBtn = el("Cancel");
      if (!wrap || !fill || !text || !confirmBtn || !cancelBtn) return;
      wrap.classList.toggle("hidden", !run);
      confirmBtn.classList.toggle("hidden", !!run);
      cancelBtn.textContent = run ? "Stop" : "Cancel";
      if (!run) return;
      var pct = run.total ? Math.round((run.done / run.total) * 100) : 0;
      fill.style.width = pct + "%";
      text.textContent =
        cfg.verb + " " + run.done + "/" + run.total +
        (run.failed ? " (" + run.failed + " failed)" : "");
    }

    function open() {
      var modal = el("Modal");
      if (!modal) return;
      modal.classList.remove("hidden");
      openBlockingModal(modal, { onEscape: close, onBackdropClick: close });
      renderProgress();
      // A run in flight owns the dialog's copy; only refresh the pickers when idle.
      if (run) return;
      if (cfg.onOpen) {
        cfg.onOpen(api);
        return;
      }
      renderScopeOptions();
      renderSummary();
    }

    function close() {
      var modal = el("Modal");
      if (!modal) return;
      closeBlockingModal(modal);
      modal.classList.add("hidden");
    }

    function submit() {
      if (run) return;
      var req = cfg.request(targets());
      if (!req) return;
      var job = { done: 0, failed: 0, total: req.total, abort: new AbortController() };
      run = job;
      var sawDone = false;
      renderProgress();

      function handleLine(line) {
        var data;
        try { data = JSON.parse(line); } catch (_) { return; }
        if (!data) return;
        if (data.token) job.token = data.token; // echoed by Stop to scope the cancel
        if (data.done) sawDone = true;
        if (cfg.onLine) cfg.onLine(data, job);
        // Header and trailing sentinel lines have no index; they stay out of the tally.
        if (typeof data.index !== "number") return;
        job.done++;
        if (!data.ok) job.failed++;
        renderProgress();
      }

      function finish(message) {
        run = null;
        renderProgress();
        close();
        showToast(message);
        if (cfg.afterFinish) cfg.afterFinish(job);
      }

      apiPostNDJSON(req.url, req.body, { signal: job.abort.signal, onLine: handleLine })
        .then(function () {
          // No sentinel means the server died partway; the reader can't tell otherwise.
          finish(cfg.toast(cfg.expectsDone && !sawDone ? "stopped" : "done", job));
        })
        .catch(function (err) {
          var aborted = err && (err.name === "AbortError" || err.code === 20);
          finish(cfg.toast(aborted ? "cancelled" : "error", job, err));
        });
    }

    function cancel() {
      if (!run) {
        close();
        return;
      }
      var request = cfg.cancel(run);
      apiPost(request.url, request.body).catch(function () {});
      run.abort.abort();
    }

    function init() {
      el("Cancel").addEventListener("click", cancel);
      el("Confirm").addEventListener("click", submit);
      el("Scope").addEventListener("change", cfg.onScopeChange || renderSummary);
      if (cfg.initExtra) cfg.initExtra(api);
    }

    api = {
      open: open,
      close: close,
      init: init,
      renderSummary: renderSummary,
      renderScopeOptions: renderScopeOptions,
      isRunning: function () { return !!run; },
    };
    return api;
  }

  // ---- Clip marked lines ----
  // One clip per mark cluster via Studio's generate-intake.

  // Mirror Studio's #trIntakeClusterThreshold and pad-0 so both pages cut identical spans.
  var CLIP_MARKS_DEFAULT_GAP_SECONDS = 10;
  var CLIP_MARKS_DEFAULT_PAD_SECONDS = 0;

  // Valid resolved marks, refetched every time the modal opens.
  var _clipMarksMarks = [];

  function _clipMarksNumber(sel, fallback, min, max) {
    var raw = parseFloat((qs(sel) || {}).value);
    if (isNaN(raw)) return fallback;
    return Math.min(max, Math.max(min, raw));
  }

  // Preview and payload share this, so the shown count is the clip count.
  function _clipMarksClusters(marks) {
    if (!marks.length) return [];
    var gap = _clipMarksNumber("#clipMarksGap", CLIP_MARKS_DEFAULT_GAP_SECONDS, 0, 120);
    return window.ClipgenIntakeCluster.clusterTranscriptMarks(marks, gap);
  }

  function _clipMarksSummary(ctx) {
    var marks = ctx.scoped;
    if (!marks.length) {
      return {
        ready: false,
        text: ctx.scope === "current" && ctx.pid
          ? "No marked lines in " + ctx.pid + " yet."
          : "No marked lines yet — mark a line with M or the gutter dot.",
      };
    }
    var clusters = _clipMarksClusters(marks);
    return {
      ready: true,
      text: clipgenPluralUnit(marks.length, "marked line", "marked lines") +
        " → " + clipgenPluralUnit(clusters.length, "clip", "clips"),
    };
  }

  function _clipMarksOnOpen(api) {
    qs("#clipMarksSummary").classList.add("cg-shimmer");
    qs("#clipMarksSummary").textContent = "Loading marks…";
    qs("#clipMarksConfirm").disabled = true;
    apiGet("api/marks")
      .then(function (data) {
        _clipMarksMarks = data.ok
          ? (data.marks || []).filter(function (m) { return m.valid; })
          : [];
        api.renderScopeOptions();
        api.renderSummary();
      })
      .catch(function () {
        qs("#clipMarksSummary").classList.remove("cg-shimmer");
        qs("#clipMarksSummary").textContent = "Could not load marks.";
      });
  }

  function _clipMarksRequest(targets) {
    var clusters = _clipMarksClusters(targets);
    if (!clusters.length) return null;
    var pad = _clipMarksNumber("#clipMarksPad", CLIP_MARKS_DEFAULT_PAD_SECONDS, 0, 10);
    // Only the start needs clamping; ffmpeg stops at EOF.
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
    return {
      url: "../studio/api/generate-intake",
      body: { items: items, format: "clip" },
      total: items.length,
    };
  }

  function _clipMarksToast(kind, job, err) {
    if (kind === "cancelled") return "Clip generation cancelled";
    if (kind === "error") return "Clip generation failed: " + (err && err.message);
    var made = job.done - job.failed;
    return job.failed
      ? clipgenPluralUnit(made, "clip", "clips") + " generated, " + job.failed + " failed"
      : clipgenPluralUnit(made, "clip", "clips") + " generated — open Studio to review";
  }

  var clipMarks = createBatchJobModal({
    prefix: "clipMarks",
    verb: "Clipping…",
    unit: ["mark", "marks"],
    candidates: function () { return _clipMarksMarks; },
    itemPid: function (m) { return m.participant; },
    currentEnabled: function (pid) { return !!pid; },
    summary: _clipMarksSummary,
    onOpen: _clipMarksOnOpen,
    request: _clipMarksRequest,
    toast: _clipMarksToast,
    cancel: function () {
      return { url: "../studio/api/generate-intake/cancel", body: {} };
    },
    initExtra: function (api) {
      qs("#clipMarksGap").addEventListener("input", api.renderSummary);
    },
  });

  // ---- Embed subtitles ----
  // Soft-mux transcripts into video copies; twin of Clip Marked Lines.

  // Multi-part participants are filtered client-side; the server would refuse them anyway.
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

  // Container extension of a participant's first source file, lowercased.
  function _embedSubsExt(p) {
    var path = (p.video_paths && p.video_paths[0]) || "";
    var dot = path.lastIndexOf(".");
    return dot === -1 ? "" : path.slice(dot).toLowerCase();
  }

  function _embedSubsContainers() {
    return CLIPGEN_CONFIG.subtitleContainers || { supported: [], alwaysDefault: [] };
  }

  // Mirrors mux_subtitles' `codec is None` guard so the summary count is honest.
  function _embedSubsIsUnsupported(p) {
    var supported = _embedSubsContainers().supported || [];
    return supported.indexOf(_embedSubsExt(p)) === -1;
  }

  function _embedSubsTargets(scoped) {
    return scoped.filter(function (p) {
      return !_embedSubsIsMultiPart(p) && !_embedSubsIsUnsupported(p);
    });
  }

  // ISOBMFF ignores -disposition:s:0 (measured, ffmpeg 8.1.2); only .mkv/.webm can be off.
  function _embedSubsAlwaysDefault(targets) {
    var always = _embedSubsContainers().alwaysDefault || [];
    return targets.filter(function (p) {
      return always.indexOf(_embedSubsExt(p)) !== -1;
    });
  }

  function _embedSubsSummary(ctx) {
    var targets = ctx.targets;
    var scoped = ctx.scoped;
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
      var text;
      // Name the real blocker; only "no transcript" is fixed by transcribing.
      if (unsupported.length && !skipped.length) {
        text =
          (unsupported.length === 1
            ? unsupported[0].id + "'s recording is " + _embedSubsExt(unsupported[0])
            : "These recordings are in a container") +
          ", which cannot carry an embedded subtitle track. Supported: " +
          (_embedSubsContainers().supported || []).join(", ") + ".";
      } else if (skipped.length) {
        text =
          (skipped.length === 1
            ? skipped[0].id + "'s transcript spans several video files"
            : "Every transcript here spans several video files") +
          ", which cannot be muxed back into one subtitled copy.";
      } else {
        text =
          ctx.scope === "current" && ctx.pid
            ? "No transcript for " + ctx.pid + " yet."
            : "No transcripts yet — transcribe a video first.";
      }
      return { ready: false, text: text };
    }
    // Only when unticked; ticked agrees with the mp4 muxer anyway.
    var stuckOn = (qs("#embedSubsDefault") || {}).checked
      ? []
      : _embedSubsAlwaysDefault(targets);
    var stuckNote = stuckOn.length
      ? " " + (stuckOn.length === targets.length ? "The track" : stuckOn.length + " of these")
        + " will still be on by default — .mp4/.mov cannot carry a subtitle track that is off."
      : "";
    return {
      ready: true,
      text: clipgenPluralUnit(targets.length, "transcript", "transcripts") +
        " → " +
        clipgenPluralUnit(targets.length, "subtitled video", "subtitled videos") +
        "." + skipNote + stuckNote,
    };
  }

  function _embedSubsRequest(targets) {
    var pids = targets.map(function (p) { return p.id; });
    if (!pids.length) return null;
    return {
      url: "api/embed-subtitles",
      body: { participants: pids, default_track: !!(qs("#embedSubsDefault") || {}).checked },
      total: pids.length,
    };
  }

  function _embedSubsOnLine(data, job) {
    if (data.output_dir) job.outputDir = data.output_dir;
  }

  function _embedSubsToast(kind, job, err) {
    if (kind === "cancelled") return "Subtitle embedding cancelled";
    if (kind === "error") return "Subtitle embedding failed: " + (err && err.message);
    var made = job.done - job.failed;
    var where = job.outputDir ? " to " + job.outputDir : "";
    if (kind === "stopped") {
      return "Subtitle embedding stopped early — " +
        clipgenPluralUnit(made, "video was", "videos were") + " written" + where +
        " of " + job.total + ". Check the clipgen log.";
    }
    return job.failed
      ? clipgenPluralUnit(made, "subtitled video", "subtitled videos") + " written" + where + ", " + job.failed + " failed"
      : clipgenPluralUnit(made, "subtitled video", "subtitled videos") + " written" + where;
  }

  function _embedSubsInitExtra(api) {
    // The checkbox decides whether the .mp4 caveat shows.
    qs("#embedSubsDefault").addEventListener("change", api.renderSummary);
  }

  var embedSubs = createBatchJobModal({
    prefix: "embedSubs",
    verb: "Embedding…",
    unit: ["transcript", "transcripts"],
    candidates: _embedSubsCandidates,
    itemPid: function (p) { return p.id; },
    targets: _embedSubsTargets,
    summary: _embedSubsSummary,
    request: _embedSubsRequest,
    onLine: _embedSubsOnLine,
    expectsDone: true,
    toast: _embedSubsToast,
    // The server stops between files; the abort just drops our end.
    cancel: function (job) {
      return { url: "api/embed-subtitles/cancel", body: { token: job.token || null } };
    },
    initExtra: _embedSubsInitExtra,
  });

  // ---- Normalize audio ----
  // Loudnorm videos in place, keeping .orig like remux.

  // pid -> parts with a kept .orig. Only fully-kept participants are excluded.
  var _normAudioKept = {};

  // Pin which participant the async-built track checkboxes belong to.
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

  function _normAudioKeptCount(p) {
    return _normAudioKept[p.id] || 0;
  }

  function _normAudioIsFullyKept(p) {
    var parts = (p.video_paths || []).length || 1;
    return _normAudioKeptCount(p) >= parts;
  }

  function _normAudioTargets(scoped) {
    return scoped.filter(function (p) {
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

  // Mode word for scope=all, checked indices for multi-track current, else "auto".
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

  function _normAudioSummary(ctx) {
    var scope = ctx.scope;
    var scoped = ctx.scoped;
    var targets = ctx.targets;
    var fullyKept = scoped.filter(_normAudioIsFullyKept);
    var resuming = targets.filter(function (p) { return _normAudioKeptCount(p) > 0; });
    var keptNote = fullyKept.length
      ? " " + fullyKept.map(function (p) { return p.id; }).join(", ") +
        " skipped (an earlier original is still kept — delete or restore it first)."
      : "";
    if (resuming.length) {
      // The server skips parts whose backup slot is occupied.
      keptNote += " " + resuming.map(function (p) { return p.id; }).join(", ") +
        (resuming.length === 1
          ? " resumes where it stopped"
          : " resume where they stopped") +
        " — already-rewritten parts are skipped.";
    }
    if (!targets.length) {
      return {
        ready: false,
        text: fullyKept.length
          ? "Nothing to normalize —" + keptNote
          : scope === "current" && ctx.pid
            ? "No source video for " + ctx.pid + "."
            : "No source videos yet.",
      };
    }
    if (scope === "current" && _normAudioTrackPid && !_normAudioTrackInfo) {
      return { ready: false, text: "Checking audio tracks…" };
    }
    var spec = _normAudioTracksSpec();
    if (Object.prototype.toString.call(spec) === "[object Array]" && !spec.length) {
      return { ready: false, text: "Select at least one track to normalize." + keptNote };
    }
    var files = _normAudioFileCount(targets);
    var trackNote =
      scope === "current" && _normAudioTrackInfo && _normAudioTrackInfo.count === 1
        ? " (1 audio track)"
        : "";
    return {
      ready: true,
      text: clipgenPluralUnit(files, "video", "videos") + trackNote +
        " → rewritten in place; " +
        (files === 1 ? "the original is" : "originals are") +
        " kept beside " + (files === 1 ? "it" : "them") + "." + keptNote,
    };
  }

  // Mode select for scope=all; per-track checkboxes for a multi-track current participant.
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
      normAudio.renderSummary();
      return;
    }
    modeLabel.classList.add("hidden");
    list.classList.add("hidden");
    list.innerHTML = "";
    _normAudioTrackPid = pid;
    _normAudioTrackInfo = null;
    normAudio.renderSummary(); // "Checking audio tracks…" while the probe runs
    _trFetchAudioInfo(pid, p.video_version).then(function (info) {
      // A scope flip or participant change while the probe ran owns the field.
      if (_normAudioTrackPid !== pid) return;
      _normAudioTrackInfo = info || { tracks: [], count: 1, auto: 0 };
      if (_normAudioTrackInfo.count > 1) {
        for (var i = 0; i < _normAudioTrackInfo.tracks.length; i++) {
          var row = document.createElement("label");
          row.className = "param-modal-label";
          var text = document.createElement("span");
          // Late-bound; transcripts-pills.js publishes the helper after the hub loads.
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
      normAudio.renderSummary();
    });
  }

  function _normAudioOnOpen(api) {
    _normAudioKept = {};
    api.renderScopeOptions();
    _renderNormAudioTrackField();
    // Kept-original state lives on disk; remux/status re-probes on every call.
    apiGet("api/remux/status")
      .then(function (data) {
        if (!data || !data.ok || api.isRunning()) return;
        var kept = data.kept || {};
        var map = {};
        for (var pid in kept) {
          if (kept[pid] && kept[pid].length) map[pid] = kept[pid].length;
        }
        _normAudioKept = map;
        api.renderSummary();
      })
      .catch(function () {});
  }

  function _normAudioRequest(targets) {
    var pids = targets.map(function (p) { return p.id; });
    if (!pids.length) return null;
    return {
      url: "api/normalize-audio",
      body: { participants: pids, tracks: _normAudioTracksSpec() },
      total: pids.length,
    };
  }

  function _normAudioOnLine(data, job) {
    // Files swapped; can be non-zero on an ok=false multi-part line.
    if (typeof data.parts_done === "number") job.changed = (job.changed || 0) + data.parts_done;
  }

  function _normAudioAfterFinish(job) {
    // Swapped files invalidate the <video> stream; reload like media-banner.js, after the toast.
    if (job.changed > 0) {
      setTimeout(function () { window.location.reload(); }, 1500);
    }
  }

  function _normAudioToast(kind, job, err) {
    if (kind === "cancelled") return "Audio normalization cancelled";
    if (kind === "error") return "Audio normalization failed: " + (err && err.message);
    var made = job.done - job.failed;
    if (kind === "stopped") {
      return "Audio normalization stopped early — " +
        clipgenPluralUnit(made, "video was", "videos were") + " rewritten of " +
        job.total + ". Check the clipgen log.";
    }
    return job.failed
      ? clipgenPluralUnit(made, "video", "videos") + " normalized, " + job.failed + " failed"
      : clipgenPluralUnit(made, "video", "videos") + " normalized; originals kept beside the sources";
  }

  function _normAudioInitExtra(api) {
    qs("#normAudioTrackMode").addEventListener("change", api.renderSummary);
    // Delegated: the checkbox rows are rebuilt per participant.
    qs("#normAudioTrackList").addEventListener("change", api.renderSummary);
  }

  var normAudio = createBatchJobModal({
    prefix: "normAudio",
    verb: "Normalizing…",
    unit: ["video", "videos"],
    candidates: _normAudioCandidates,
    itemPid: function (p) { return p.id; },
    targets: _normAudioTargets,
    summary: _normAudioSummary,
    onOpen: _normAudioOnOpen,
    request: _normAudioRequest,
    onLine: _normAudioOnLine,
    expectsDone: true,
    toast: _normAudioToast,
    // Unlike embed, Stop interrupts ffmpeg mid-encode; the abort just drops our end.
    cancel: function (job) {
      return { url: "api/normalize-audio/cancel", body: { token: job.token || null } };
    },
    afterFinish: _normAudioAfterFinish,
    // Scope picks the track control; the field re-renders the summary itself.
    onScopeChange: _renderNormAudioTrackField,
    initExtra: _normAudioInitExtra,
  });

  // ---- Satellite exports ----
  TS.openClipMarksModal = clipMarks.open; // hub (quick action)
  TS.openEmbedSubsModal = embedSubs.open; // hub (quick action)
  TS.openNormalizeAudioModal = normAudio.open; // hub (quick action)
  TS.initBatchModals = function () {
    clipMarks.init();
    embedSubs.init();
    normAudio.init();
  }; // hub boot
})();
