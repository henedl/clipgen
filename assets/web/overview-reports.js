/* Overview Reports tab — per-participant mini-report (overview-reports.js).
 *
 * Feeds three data sources into the local Ollama "report" thinking agent:
 * sheet observations (hub state.sheetData), the transcript summary, and
 * marked transcript lines (hub state.trIntakeMarks). Generation is manual:
 * the backend agent is disabled by default and only runs through the generic
 * regenerate route (../transcripts/api/agent/report/<pid>/regenerate).
 * Missing upstream data is triggerable in place (Transcribe → Generate
 * summary); marks are hand-made, so their empty state deep-links to
 * /transcripts/#<pid> instead.
 *
 * The Clips section cuts the participant's sheet timestamps into real clips
 * through Studio's streaming ../studio/api/generate (skip-on-repeat is free —
 * the cell-keyed artifact cache serves prior runs instantly), lists them from
 * ../studio/api/manifest, posters them via ../studio/api/thumbnail, and plays
 * them from the /studio/media/ output-dir route. Report [M:SS] stamps that a
 * clip covers become playable chips; severity labels inside (category,
 * severity) groups are painted with the shared .sev-* palette.
 *
 * All hub data comes from the overview.js hub via window.ClipgenOverview
 * (lazy reads inside activate(), never top-level destructures). Lifecycle:
 * OV.reportsActivate / reportsDeactivate / reportsResize. Participant
 * selection is a single key on purpose — a future aggregate mode extends
 * rpState.selected to a set without reshaping the tab.
 */

(function () {
  "use strict";

  var state; // hub state, set lazily in activate()

  var REPORT_POLL_MS = 1200;
  var TASK_POLL_MS = 3000;
  // Consecutive idle participant-poll ticks required before the poll stops:
  // a just-fired trigger can race the worker/orchestrator claiming its slot,
  // so one idle reading straight after a POST is not proof nothing runs.
  var TASK_IDLE_TICKS_TO_STOP = 2;

  var rpState = {
    active: false,
    initialized: false,
    _snapshot: null, // { version: state.dataVersion } — hub staleness contract
    participants: [], // ../transcripts/api/participants ∪ sheet-only ids
    selected: null,
    gen: 0, // generation counter — bumped on participant switch
    ollama: null, // /api/models "ollama" payload (null = not fetched/unknown)
    report: null, // stored report payload for the selected participant
    reportGenerating: false,
    reportPartial: "",
    reportMissing: false,
    reportPoll: null,
    taskPoll: null,
    taskIdleTicks: 0,
    clips: [], // studio-manifest clip artifacts for the selected participant
    clipsLoaded: false,
    clipsGenerating: false,
    clipsDone: 0,
    clipsTotal: 0,
    clipsAbort: null, // AbortController for the streaming generate fetch
    playingClip: null, // artifact id in the inline player
  };

  var dom = {}; // scaffold refs, filled by initDom()

  // ---- Data helpers ----

  function rec() {
    for (var i = 0; i < rpState.participants.length; i++) {
      if (rpState.participants[i].id === rpState.selected) {
        return rpState.participants[i];
      }
    }
    return null;
  }

  function agentState(r, key) {
    return (r && r.agents && r.agents[key]) || "idle";
  }

  function countObservations(pid) {
    if (!state.sheetData || !state.sheetData.rows) return 0;
    var n = 0;
    for (var i = 0; i < state.sheetData.rows.length; i++) {
      var cell = state.sheetData.rows[i].cells[pid];
      if (cell && cell.hasText) n++;
    }
    return n;
  }

  function inSheet(pid) {
    var list = (state.sheetData && state.sheetData.participants) || [];
    return list.indexOf(pid) !== -1;
  }

  function marksFor(pid) {
    var out = [];
    for (var i = 0; i < state.trIntakeMarks.length; i++) {
      if (state.trIntakeMarks[i].participant === pid) out.push(state.trIntakeMarks[i]);
    }
    return out;
  }

  function markBreakdown(marks) {
    var counts = {};
    var order = [];
    for (var i = 0; i < marks.length; i++) {
      var cat = marks[i].category || "bookmark";
      if (counts[cat] === undefined) {
        counts[cat] = 0;
        order.push(cat);
      }
      counts[cat]++;
    }
    var parts = [];
    for (var j = 0; j < order.length && j < 3; j++) {
      var meta = MARK_CATEGORIES[order[j]];
      parts.push(counts[order[j]] + " " + ((meta && meta.label) || order[j]));
    }
    if (order.length > 3) parts.push("…");
    return parts.join(", ");
  }

  function mergeSheetParticipants(list) {
    var out = list.slice();
    var seen = {};
    for (var i = 0; i < out.length; i++) seen[out[i].id] = true;
    var sheetIds = (state.sheetData && state.sheetData.participants) || [];
    for (var j = 0; j < sheetIds.length; j++) {
      if (!seen[sheetIds[j]]) {
        out.push({ id: sheetIds[j], in_sheet: true, has_video: false, has_transcript: false, agents: {} });
      }
    }
    return out;
  }

  function anyUpstreamRunning() {
    for (var i = 0; i < rpState.participants.length; i++) {
      var p = rpState.participants[i];
      if (agentState(p, "transcription") === "running" || agentState(p, "summary") === "running") {
        return true;
      }
    }
    return false;
  }

  // Generate is blocked only when Ollama is positively unreachable or the
  // report model is positively missing (same stance as the Transcripts page:
  // an unknown /api/models state never blocks).
  function ollamaGate() {
    var o = rpState.ollama;
    if (!o) return null;
    var status = clipgenOllamaStatus(o);
    if (status.state !== "ok") {
      // The install command only helps the "missing" case; a stopped server
      // just needs starting, and this panel's Refresh re-checks either way.
      var extra = status.state === "missing" && status.hint.length
        ? " " + status.hint[0]
        : " Start it, then Refresh.";
      return status.message + extra;
    }
    var agents = o.agents || [];
    for (var i = 0; i < agents.length; i++) {
      if (agents[i].key === "report" && agents[i].installed === false) {
        return "Ollama model " + agents[i].model + " is not installed — install it from the Transcripts page, or pick another in Settings → Summaries.";
      }
    }
    return null;
  }

  // ---- Scaffold ----

  function initDom() {
    if (rpState.initialized) return;
    rpState.initialized = true;
    var panel = qs("#reportsPanel");
    if (!panel) return;

    var layout = el("div", "rp-layout");

    var sidebar = el("aside", "rp-sidebar");
    sidebar.appendChild(el("div", "rp-sidebar-header", "Participants"));
    dom.pills = el("div", "rp-pills");
    sidebar.appendChild(dom.pills);
    layout.appendChild(sidebar);

    var main = el("section", "rp-main");
    dom.empty = el("div", "rp-empty hidden");
    dom.empty.appendChild(el("p", "", "No participants yet."));
    var emptySub = el("p", "rp-empty-sub");
    emptySub.appendChild(document.createTextNode("Add videos to the input folder or open a spreadsheet in "));
    var studioLink = el("a", "", "Studio");
    studioLink.href = "/studio/";
    emptySub.appendChild(studioLink);
    emptySub.appendChild(document.createTextNode(", then come back."));
    dom.empty.appendChild(emptySub);
    main.appendChild(dom.empty);

    dom.content = el("div", "rp-content");
    dom.sources = el("div", "rp-sources");
    dom.content.appendChild(dom.sources);
    dom.note = el("div", "rp-note hidden");
    dom.content.appendChild(dom.note);

    var report = el("div", "rp-report");
    var head = el("div", "rp-report-head");
    head.appendChild(el("h2", "rp-report-title", "Mini-report"));
    dom.actions = el("div", "rp-report-actions");
    head.appendChild(dom.actions);
    report.appendChild(head);
    dom.meta = el("div", "rp-report-meta");
    report.appendChild(dom.meta);
    dom.body = el("div", "rp-report-body");
    report.appendChild(dom.body);
    dom.content.appendChild(report);

    // Report timestamps are re-rendered wholesale, so clip playback rides one
    // delegated listener instead of per-span handlers.
    dom.body.addEventListener("click", function (ev) {
      var target = ev.target && ev.target.closest ? ev.target.closest(".rp-ts--linked") : null;
      if (!target) return;
      var sec = parseFloat(target.getAttribute("data-seconds"));
      if (isNaN(sec)) return;
      var clip = findClipAt(sec);
      if (clip) playClip(clip, Math.max(0, sec - (clip.start || 0)));
    });

    dom.clipsSection = el("div", "rp-clips hidden");
    var clipsHead = el("div", "rp-clips-head");
    clipsHead.appendChild(el("h2", "rp-clips-title", "Clips"));
    dom.clipsMeta = el("span", "rp-clips-meta");
    clipsHead.appendChild(dom.clipsMeta);
    dom.clipsActions = el("div", "rp-clips-actions");
    clipsHead.appendChild(dom.clipsActions);
    dom.clipsSection.appendChild(clipsHead);
    dom.player = el("div", "rp-player hidden");
    dom.clipsSection.appendChild(dom.player);
    dom.clipsStrip = el("div", "rp-clips-strip");
    dom.clipsSection.appendChild(dom.clipsStrip);
    dom.content.appendChild(dom.clipsSection);

    main.appendChild(dom.content);
    layout.appendChild(main);
    panel.appendChild(layout);
  }

  // ---- Rendering ----

  function renderAll() {
    renderParticipants();
    renderMain();
  }

  function renderParticipants() {
    if (!dom.pills) return;
    var P = window.ClipgenPrimitives;
    var frag = document.createDocumentFragment();
    for (var i = 0; i < rpState.participants.length; i++) {
      var p = rpState.participants[i];
      var pill = P.createParticipantPill({
        id: p.id,
        active: p.id === rpState.selected,
        dataset: { pid: p.id },
        onClick: onPillClick,
      });
      var reportState = agentState(p, "report");
      if (reportState === "done" || reportState === "running") {
        pill.appendChild(el("span", "rp-pill-dot is-" + reportState));
      }
      frag.appendChild(pill);
    }
    dom.pills.innerHTML = "";
    dom.pills.appendChild(frag);
  }

  function onPillClick() {
    selectParticipant(this.dataset.pid);
  }

  function renderMain() {
    if (!dom.content) return;
    var hasParticipants = rpState.participants.length > 0;
    dom.empty.classList.toggle("hidden", hasParticipants);
    dom.content.classList.toggle("hidden", !hasParticipants);
    if (!hasParticipants) return;
    renderSources();
    renderReportArea();
  }

  function sourceRow(labelText, status, kind, actionEl) {
    var row = el("div", "rp-source-row");
    row.appendChild(el("span", "rp-source-dot is-" + kind));
    row.appendChild(el("span", "rp-source-label", labelText));
    var statusEl = el("span", "rp-source-status");
    if (typeof status === "string") statusEl.textContent = status;
    else statusEl.appendChild(status);
    row.appendChild(statusEl);
    if (actionEl) {
      var act = el("span", "rp-source-action");
      act.appendChild(actionEl);
      row.appendChild(act);
    }
    return row;
  }

  function transcriptsLink(pid, label) {
    var a = el("a", "", label);
    a.href = "/transcripts/#" + pid;
    return a;
  }

  function renderSources() {
    var pid = rpState.selected;
    var r = rec();
    dom.sources.innerHTML = "";
    if (!pid || !r) return;
    var P = window.ClipgenPrimitives;
    var frag = document.createDocumentFragment();

    // 1. Sheet observations — informational only (a sheet can't be triggered).
    if (!state.sheetData || !state.sheetData.sheet_loaded) {
      frag.appendChild(sourceRow("Sheet observations", "no spreadsheet loaded", "off", null));
    } else if (!inSheet(pid)) {
      frag.appendChild(sourceRow("Sheet observations", "not in sheet", "off", null));
    } else {
      var nObs = countObservations(pid);
      frag.appendChild(sourceRow(
        "Sheet observations",
        clipgenPluralUnit(nObs, "observation", "observations"),
        nObs > 0 ? "ok" : "warn",
        null
      ));
    }

    // 2. Transcript summary — the report's hard dependency; triggerable here.
    var transcription = agentState(r, "transcription");
    var summary = agentState(r, "summary");
    if (transcription === "running") {
      frag.appendChild(sourceRow("Transcript summary", "transcribing…", "busy", null));
    } else if (!r.has_transcript) {
      if (!r.has_video) {
        frag.appendChild(sourceRow("Transcript summary", "no video found", "off", null));
      } else {
        frag.appendChild(sourceRow(
          "Transcript summary", "no transcript", "warn",
          P.createBtn({ label: "Transcribe", icon: "microphone", size: "sm", onClick: function () { transcribe(false); } })
        ));
      }
    } else if (summary === "running") {
      frag.appendChild(sourceRow("Transcript summary", "generating summary…", "busy", null));
    } else if (summary === "done") {
      frag.appendChild(sourceRow("Transcript summary", "ready", "ok", null));
    } else {
      var summaryBtn = P.createBtn({
        label: "Generate summary",
        icon: "octicon/dependabot-16",
        size: "sm",
        onClick: generateSummary,
      });
      summaryBtn.setAttribute("data-tooltip", "Runs a local AI thinking agent");
      frag.appendChild(sourceRow("Transcript summary", "no summary", "warn", summaryBtn));
    }

    // 3. Marked lines — hand-made in Transcripts, so link instead of trigger.
    var marks = marksFor(pid);
    if (marks.length > 0) {
      frag.appendChild(sourceRow(
        "Marked lines",
        clipgenPluralUnit(marks.length, "marked line", "marked lines") + " · " + markBreakdown(marks),
        "ok",
        null
      ));
    } else {
      var linkStatus = el("span", "");
      linkStatus.appendChild(document.createTextNode("none — "));
      linkStatus.appendChild(transcriptsLink(pid, "mark lines in Transcripts"));
      frag.appendChild(sourceRow("Marked lines", linkStatus, "warn", null));
    }

    dom.sources.appendChild(frag);
  }

  function metaLine(report) {
    var parts = [];
    if (report.model) parts.push(report.model);
    if (report.generated_at) {
      var d = new Date(report.generated_at);
      if (!isNaN(d.getTime())) parts.push(d.toLocaleString());
    }
    var src = report.sources || {};
    parts.push(
      clipgenPluralUnit(src.observations || 0, "observation", "observations") +
      ", " + clipgenPluralUnit(src.bookmarks || 0, "marked line", "marked lines") + " used"
    );
    return parts.join(" · ");
  }

  function tsToSeconds(ts) {
    var parts = ts.split(":");
    var sec = 0;
    for (var i = 0; i < parts.length; i++) sec = sec * 60 + parseInt(parts[i], 10);
    return sec;
  }

  function findClipAt(sec) {
    // ±2 s tolerance: report timestamps are whole seconds while clip bounds
    // carry sub-second starts.
    for (var i = 0; i < rpState.clips.length; i++) {
      var c = rpState.clips[i];
      if (sec >= (c.start || 0) - 2 && sec <= (c.end || 0) + 2) return c;
    }
    return null;
  }

  // Wrap [M:SS] / [H:MM:SS] stamps in mono chips; ones a generated clip covers
  // become playable (the delegated dom.body handler). Runs on escaped HTML.
  function decorateTimestamps(html) {
    return html.replace(/\[(\d{1,2}:\d{2}(?::\d{2})?)\]/g, function (m, ts) {
      var sec = tsToSeconds(ts);
      var cls = findClipAt(sec) ? "rp-ts rp-ts--linked" : "rp-ts";
      return '<span class="' + cls + '" data-seconds="' + sec + '">[' + ts + "]</span>";
    });
  }

  // Color severity labels, but only inside tag-free parenthesized groups —
  // the "(category, severity)" shape the prompt's source lines carry — so a
  // bare "high" in prose is never painted. Longest label first, one hit per
  // group ("Very Positive" must not be re-matched by "Positive").
  function decorateSeverity(html) {
    var sevs = (CLIPGEN_CONFIG.severity || []).slice();
    sevs.sort(function (a, b) { return b.label.length - a.label.length; });
    return html.replace(/\(([^()<]*)\)/g, function (m, inner) {
      for (var i = 0; i < sevs.length; i++) {
        var esc = sevs[i].label.replace(/[.*+?^${}()|[\]\\/]/g, "\\$&");
        var re = new RegExp("(^|[\\s,])(" + esc + ")(?=$|[\\s,])", "i");
        if (re.test(inner)) {
          inner = inner.replace(
            re,
            '$1<span class="rp-sev ' + sevs[i].cssClass + '">$2</span>'
          );
          break;
        }
      }
      return "(" + inner + ")";
    });
  }

  function decorateReportInline(text) {
    return decorateSeverity(decorateTimestamps(clipgenRenderInlineMarkdown(text)));
  }

  // Minimal markdown for the report shape the prompt asks for (## headings,
  // "- " bullets, paragraphs) plus inline emphasis, timestamp chips, and
  // severity coloring. Everything passes through escapeHtml (inside
  // clipgenRenderInlineMarkdown); the model never gets to inject markup.
  function renderReportText(text) {
    var lines = String(text || "").split("\n");
    var html = "";
    var inList = false;
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) {
        if (inList) { html += "</ul>"; inList = false; }
        continue;
      }
      var heading = line.match(/^(#{2,4})\s+(.*)$/);
      if (heading) {
        if (inList) { html += "</ul>"; inList = false; }
        var tag = heading[1].length === 2 ? "h3" : "h4";
        html += "<" + tag + ">" + clipgenRenderInlineMarkdown(heading[2]) + "</" + tag + ">";
      } else if (line.indexOf("- ") === 0 || line.indexOf("* ") === 0) {
        if (!inList) { html += "<ul>"; inList = true; }
        html += "<li>" + decorateReportInline(line.substring(2)) + "</li>";
      } else {
        if (inList) { html += "</ul>"; inList = false; }
        html += "<p>" + decorateReportInline(line) + "</p>";
      }
    }
    if (inList) html += "</ul>";
    return html;
  }

  function renderReportBodyPartial() {
    dom.body.innerHTML = "";
    var pre = el("div", "rp-report-partial");
    pre.textContent = rpState.reportPartial || "Waiting for the model…";
    dom.body.appendChild(pre);
  }

  function renderReportArea() {
    if (!dom.actions) return;
    dom.actions.innerHTML = "";
    dom.meta.textContent = "";
    var pid = rpState.selected;
    var r = rec();
    if (!pid || !r) {
      dom.body.innerHTML = "";
      return;
    }
    var P = window.ClipgenPrimitives;

    if (rpState.reportGenerating) {
      dom.actions.appendChild(P.createBtn({ label: "Stop", icon: "stop", size: "sm", onClick: stopReport }));
      dom.meta.textContent = "Generating…";
      renderReportBodyPartial();
      return;
    }

    var canGenerate = !!r.has_transcript && agentState(r, "summary") === "done";
    var gate = ollamaGate();
    var genBtn = P.createBtn({
      label: rpState.report ? "Regenerate" : "Generate report",
      icon: "octicon/dependabot-16",
      variant: "solid",
      size: "sm",
      disabled: !canGenerate || !!gate,
      onClick: generateReport,
    });
    genBtn.setAttribute("data-tooltip", "Runs a local AI thinking agent");
    dom.actions.appendChild(genBtn);

    if (rpState.report) {
      dom.meta.textContent = metaLine(rpState.report);
      dom.body.innerHTML = renderReportText(rpState.report.text);
      return;
    }

    var hint;
    if (!canGenerate) {
      hint = "Needs a transcript summary first — trigger the missing steps above.";
    } else if (gate) {
      hint = gate;
    } else if (rpState.reportMissing) {
      hint = "No report yet. Generate one from the sources above.";
    } else {
      hint = "Loading…";
    }
    dom.body.innerHTML = "";
    dom.body.appendChild(el("p", "rp-report-hint", hint));
  }

  // ---- Notices (transcribe download gate, trigger failures) ----

  function showNotice(text, actionLabel, onAction, linkPid) {
    dom.note.innerHTML = "";
    dom.note.classList.remove("hidden");
    dom.note.appendChild(el("span", "rp-note-text", text));
    if (actionLabel) {
      dom.note.appendChild(window.ClipgenPrimitives.createBtn({
        label: actionLabel, size: "sm", onClick: onAction,
      }));
    }
    if (linkPid) dom.note.appendChild(transcriptsLink(linkPid, "Open Transcripts"));
  }

  function clearNotice() {
    if (!dom.note) return;
    dom.note.classList.add("hidden");
    dom.note.innerHTML = "";
  }

  // ---- Selection + data loading ----

  function selectParticipant(pid) {
    if (rpState.selected === pid) return;
    if (rpState.clipsGenerating) cancelClips();
    rpState.selected = pid;
    rpState.gen++;
    stopReportPoll();
    clearNotice();
    closePlayer();
    rpState.report = null;
    rpState.reportGenerating = false;
    rpState.reportPartial = "";
    rpState.reportMissing = false;
    rpState.clips = [];
    rpState.clipsLoaded = false;
    renderAll();
    fetchReport(pid, rpState.gen);
    loadClips(pid, rpState.gen);
  }

  function ensureSelection() {
    var ids = {};
    for (var i = 0; i < rpState.participants.length; i++) ids[rpState.participants[i].id] = true;
    if (rpState.selected && !ids[rpState.selected]) rpState.selected = null;
    if (!rpState.selected && rpState.participants.length > 0) {
      selectParticipant(rpState.participants[0].id);
    }
  }

  function loadParticipants() {
    return apiGet("../transcripts/api/participants")
      .then(function (data) {
        if (!rpState.active) return;
        rpState.participants = mergeSheetParticipants((data && data.participants) || []);
        ensureSelection();
        renderAll();
      })
      .catch(function () {
        if (!rpState.active) return;
        rpState.participants = mergeSheetParticipants([]);
        ensureSelection();
        renderAll();
      });
  }

  function loadModels() {
    apiGet("/api/models")
      .then(function (data) {
        rpState.ollama = (data && data.ollama) || null;
        if (rpState.active) renderMain();
      })
      .catch(function () {
        rpState.ollama = null;
      });
  }

  // ---- Report fetch/poll ----

  function applyReportResponse(data, pid, g) {
    if (data && data.ok && data.report) {
      stopReportPoll();
      rpState.report = data.report;
      rpState.reportGenerating = false;
      rpState.reportPartial = "";
      rpState.reportMissing = false;
      renderReportArea();
      // The pill dot (agents.report) is served by the participants API; a
      // local nudge keeps it honest until the next refetch.
      var r = rec();
      if (r && r.agents) r.agents.report = "done";
      renderParticipants();
    } else if (data && data.generating) {
      rpState.reportGenerating = true;
      rpState.reportPartial = data.partial || "";
      if (!rpState.reportPoll) {
        renderReportArea();
        startReportPoll(pid, g);
      } else {
        renderReportBodyPartial();
      }
    } else {
      stopReportPoll();
      rpState.reportGenerating = false;
      rpState.reportMissing = true;
      renderReportArea();
    }
  }

  function fetchReport(pid, g) {
    apiGet("../transcripts/api/agent/report/" + pid)
      .then(function (data) {
        if (g !== rpState.gen || !rpState.active) return;
        applyReportResponse(data, pid, g);
      })
      .catch(function () {
        // 404 = no report stored for this participant (apiGet throws on it).
        if (g !== rpState.gen || !rpState.active) return;
        stopReportPoll();
        rpState.reportGenerating = false;
        rpState.reportMissing = true;
        renderReportArea();
      });
  }

  function startReportPoll(pid, g) {
    stopReportPoll();
    rpState.reportPoll = setInterval(function () {
      if (!rpState.active || g !== rpState.gen) {
        stopReportPoll();
        return;
      }
      if (document.hidden) return;
      fetchReport(pid, g);
    }, REPORT_POLL_MS);
  }

  function stopReportPoll() {
    if (rpState.reportPoll) {
      clearInterval(rpState.reportPoll);
      rpState.reportPoll = null;
    }
  }

  function generateReport() {
    var pid = rpState.selected;
    if (!pid) return;
    var g = rpState.gen;
    clearNotice();
    apiPost("../transcripts/api/agent/report/" + pid + "/regenerate", {})
      .then(function () {
        if (g !== rpState.gen || !rpState.active) return;
        rpState.report = null;
        rpState.reportGenerating = true;
        rpState.reportPartial = "";
        renderReportArea();
        startReportPoll(pid, g);
      })
      .catch(function () {
        if (g !== rpState.gen || !rpState.active) return;
        showNotice("Could not start the report — re-check the sources above.", null, null, null);
      });
  }

  function stopReport() {
    var pid = rpState.selected;
    if (!pid) return;
    var g = rpState.gen;
    apiPost("../transcripts/api/agent/report/" + pid + "/stop", {})
      .then(function () {
        if (g !== rpState.gen || !rpState.active) return;
        stopReportPoll();
        rpState.reportGenerating = false;
        rpState.reportPartial = "";
        fetchReport(pid, g);
      })
      .catch(function () {});
  }

  // ---- Clips (generated from the participant's sheet timestamps) ----

  function sheetCellsFor(pid) {
    var cells = [];
    if (!state.sheetData || !state.sheetData.rows) return cells;
    for (var i = 0; i < state.sheetData.rows.length; i++) {
      var row = state.sheetData.rows[i];
      var cell = row.cells[pid];
      if (cell && cell.valid) cells.push(pid + "." + row.rowNum);
    }
    return cells;
  }

  function loadClips(pid, g) {
    apiGet("../studio/api/manifest")
      .then(function (data) {
        if (g !== rpState.gen || !rpState.active) return;
        var arts = (data && data.artifacts) || [];
        var clips = [];
        for (var i = 0; i < arts.length; i++) {
          if (arts[i].participant === pid && arts[i].type === "clip") clips.push(arts[i]);
        }
        clips.sort(function (a, b) { return (a.start || 0) - (b.start || 0); });
        rpState.clips = clips;
        rpState.clipsLoaded = true;
        renderClips();
        // Timestamp chips in the report link up only where a clip covers them.
        renderReportArea();
      })
      .catch(function () {
        if (g !== rpState.gen || !rpState.active) return;
        rpState.clips = [];
        rpState.clipsLoaded = true;
        renderClips();
      });
  }

  function generateClips() {
    var pid = rpState.selected;
    var g = rpState.gen;
    var cells = sheetCellsFor(pid);
    if (!pid || !cells.length || rpState.clipsGenerating) return;
    clearNotice();
    rpState.clipsGenerating = true;
    rpState.clipsDone = 0;
    rpState.clipsTotal = cells.length;
    var ctrl = new AbortController();
    rpState.clipsAbort = ctrl;
    renderClips();
    fetch("../studio/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ cells: cells, format: "clip" }),
      signal: ctrl.signal,
    })
      .then(function (resp) {
        if (resp.status === 409) throw new Error("generate-busy");
        if (!resp.ok) throw new Error("Server error " + resp.status);
        return readNDJSONStream(resp, function (line) {
          if (!line || line.cancelled) return;
          rpState.clipsDone++;
          renderClipsProgress();
        });
      })
      .then(function () {
        resetClipsGeneration();
        if (g !== rpState.gen || !rpState.active) return;
        loadClips(pid, g);
      })
      .catch(function (err) {
        resetClipsGeneration();
        if (g !== rpState.gen || !rpState.active) return;
        if (err && err.message === "generate-busy") {
          showNotice("Studio is already generating — wait for it to finish, then try again.", null, null, null);
        } else if (!(err && err.name === "AbortError")) {
          showNotice("Clip generation failed.", null, null, null);
        }
        // Cancelled or partial runs still produced clips — pick them up.
        loadClips(pid, g);
      });
  }

  function resetClipsGeneration() {
    rpState.clipsGenerating = false;
    rpState.clipsAbort = null;
    rpState.clipsDone = 0;
    rpState.clipsTotal = 0;
  }

  function cancelClips() {
    apiPost("../studio/api/generate/cancel", {}).catch(function () {});
    if (rpState.clipsAbort) rpState.clipsAbort.abort();
  }

  function renderClipsProgress() {
    if (!rpState.clipsGenerating || !dom.clipsMeta) return;
    dom.clipsMeta.textContent = rpState.clipsDone + "/" + rpState.clipsTotal + " cells";
    if (dom.clipsGenBtn) {
      var frac = rpState.clipsTotal ? rpState.clipsDone / rpState.clipsTotal : 0;
      window.ClipgenPrimitives.setButtonProgress(dom.clipsGenBtn, frac);
    }
  }

  function clipTile(a) {
    var tile = el("button", "rp-clip");
    tile.type = "button";
    var sev = severityClass(a.severity);
    if (sev) tile.classList.add(sev);
    if (rpState.playingClip === a.id) tile.classList.add("is-playing");
    var img = document.createElement("img");
    img.className = "rp-clip-thumb";
    img.loading = "lazy";
    img.alt = "";
    img.src =
      "../studio/api/thumbnail/" +
      encodeURIComponent(a.participant) +
      "/" +
      Math.max(0, Math.floor(a.start || 0));
    img.addEventListener("error", function () {
      tile.classList.add("rp-clip--noposter");
    });
    tile.appendChild(img);
    var meta = el("span", "rp-clip-meta");
    meta.appendChild(el("span", "rp-clip-time cg-mono", formatTime(a.start || 0)));
    if (a.severity) {
      meta.appendChild(el("span", "rp-clip-sev " + (sev || ""), a.severity));
    }
    tile.appendChild(meta);
    if (a.description) tile.setAttribute("data-tooltip", a.description);
    tile.addEventListener("click", function () {
      playClip(a, 0);
    });
    return tile;
  }

  function renderClips() {
    if (!dom.clipsSection) return;
    var pid = rpState.selected;
    var cells = pid ? sheetCellsFor(pid) : [];
    var show = !!pid && (cells.length > 0 || rpState.clips.length > 0);
    dom.clipsSection.classList.toggle("hidden", !show);
    if (!show) {
      closePlayer();
      return;
    }

    dom.clipsActions.innerHTML = "";
    dom.clipsGenBtn = null;
    var P = window.ClipgenPrimitives;
    if (rpState.clipsGenerating) {
      dom.clipsGenBtn = P.createBtn({ label: "Generating…", icon: "film", size: "sm", disabled: true });
      dom.clipsActions.appendChild(dom.clipsGenBtn);
      dom.clipsActions.appendChild(P.createBtn({ label: "Cancel", size: "sm", onClick: cancelClips }));
      renderClipsProgress();
    } else {
      if (cells.length) {
        var genBtn = P.createBtn({ label: "Generate clips", icon: "film", size: "sm", onClick: generateClips });
        genBtn.setAttribute(
          "data-tooltip",
          "Cut this participant's sheet timestamps into clips (already-generated ones are skipped)"
        );
        dom.clipsActions.appendChild(genBtn);
      }
      dom.clipsMeta.textContent = rpState.clips.length
        ? clipgenPluralUnit(rpState.clips.length, "clip", "clips")
        : (rpState.clipsLoaded ? "no clips yet" : "");
    }

    dom.clipsStrip.innerHTML = "";
    if (rpState.clips.length) {
      var frag = document.createDocumentFragment();
      for (var i = 0; i < rpState.clips.length; i++) {
        frag.appendChild(clipTile(rpState.clips[i]));
      }
      dom.clipsStrip.appendChild(frag);
    } else if (rpState.clipsLoaded && !rpState.clipsGenerating && cells.length) {
      dom.clipsStrip.appendChild(
        el("p", "rp-clips-hint", "No clips yet — generate them from this participant's sheet timestamps.")
      );
    }
  }

  function playClip(a, offset) {
    rpState.playingClip = a.id;
    dom.player.innerHTML = "";
    dom.player.classList.remove("hidden");
    var head = el("div", "rp-player-head");
    head.appendChild(el("span", "rp-player-title", a.description || a.file || ""));
    head.appendChild(window.ClipgenPrimitives.createBtn({ label: "Close", size: "sm", onClick: closePlayer }));
    dom.player.appendChild(head);
    var vid = document.createElement("video");
    vid.className = "rp-player-video";
    vid.controls = true;
    vid.autoplay = true;
    vid.src = "../studio/media/" + encodeURIComponent(a.file);
    if (offset > 0.5) {
      vid.addEventListener("loadedmetadata", function () {
        if (vid.duration && offset < vid.duration) vid.currentTime = offset;
      });
    }
    dom.player.appendChild(vid);
    renderClips(); // move the is-playing highlight
  }

  function closePlayer() {
    if (rpState.playingClip === null) return;
    rpState.playingClip = null;
    if (dom.player) {
      dom.player.innerHTML = "";
      dom.player.classList.add("hidden");
    }
    renderClips();
  }

  // ---- Upstream triggers (transcription, summary) ----

  function generateSummary() {
    var pid = rpState.selected;
    if (!pid) return;
    clearNotice();
    apiPost("../transcripts/api/agent/summary/" + pid + "/regenerate", {})
      .then(function () {
        if (!rpState.active) return;
        // The backend clears dependent agent results (report included) when
        // the summary regenerates — drop the in-memory copy too, or the stale
        // mini-report keeps rendering until the participant is re-selected.
        stopReportPoll();
        rpState.report = null;
        rpState.reportGenerating = false;
        rpState.reportPartial = "";
        rpState.reportMissing = true;
        var r = rec();
        if (r && r.agents) {
          r.agents.summary = "running";
          r.agents.report = "idle"; // renderAll refreshes the pill dot too
        }
        renderAll();
        startTaskPoll();
      })
      .catch(function () {
        if (!rpState.active) return;
        showNotice("Could not start the summary — is there a transcript?", null, null, pid);
      });
  }

  function transcribe(allowDownload) {
    var pid = rpState.selected;
    if (!pid) return;
    clearNotice();
    apiPost("../transcripts/api/transcribe", { participants: [pid], allow_download: !!allowDownload })
      .then(function (data) {
        if (!rpState.active) return;
        if (data && data.ok) {
          var r = rec();
          if (r && r.agents) r.agents.transcription = "running";
          renderMain();
          startTaskPoll();
          return;
        }
        if (data && data.reason === "model_not_cached") {
          var un = (data.uncached && data.uncached[0]) || {};
          var size = un.size_mb ? " (~" + un.size_mb + " MB)" : "";
          showNotice(
            "Whisper model " + (un.model || "") + size + " is not downloaded.",
            "Download & transcribe",
            function () { transcribe(true); },
            pid
          );
          return;
        }
        showNotice("Could not start transcription.", null, null, pid);
      })
      .catch(function () {
        if (!rpState.active) return;
        showNotice("Could not start transcription.", null, null, pid);
      });
  }

  // While a triggered upstream task runs, re-poll the participants list (a
  // quiet-poll path) so the source rows track transcription → summary → done,
  // and the Generate button unlocks the moment the summary lands.
  function startTaskPoll() {
    rpState.taskIdleTicks = 0;
    if (rpState.taskPoll) return;
    rpState.taskPoll = setInterval(function () {
      if (!rpState.active) {
        stopTaskPoll();
        return;
      }
      if (document.hidden) return;
      apiGet("../transcripts/api/participants")
        .then(function (data) {
          if (!rpState.active) return;
          rpState.participants = mergeSheetParticipants((data && data.participants) || []);
          ensureSelection();
          renderAll();
          if (anyUpstreamRunning()) {
            rpState.taskIdleTicks = 0;
          } else {
            rpState.taskIdleTicks++;
            if (rpState.taskIdleTicks >= TASK_IDLE_TICKS_TO_STOP) stopTaskPoll();
          }
        })
        .catch(function () {});
    }, TASK_POLL_MS);
  }

  function stopTaskPoll() {
    if (rpState.taskPoll) {
      clearInterval(rpState.taskPoll);
      rpState.taskPoll = null;
    }
  }

  // ---- Staleness (hub dataVersion contract) ----

  function takeSnapshot() {
    rpState._snapshot = { version: state.dataVersion };
  }

  function checkStaleness() {
    if (!rpState._snapshot || !rpState.active) return;
    window.ClipgenOverview.setRefreshStale(rpState._snapshot.version !== state.dataVersion);
  }

  // ---- Lifecycle ----

  function activate() {
    rpState.active = true;
    if (!state) state = window.ClipgenOverview.state;
    initDom();
    if (rpState._snapshot) checkStaleness();
    window.ClipgenOverview.ensureData().then(function () {
      if (!rpState.active) return;
      loadParticipants().then(function () {
        if (!rpState.active) return;
        takeSnapshot();
        checkStaleness();
        // Re-activation keeps the selection, so selectParticipant's clip load
        // never fires — refresh the strip here (new clips may have landed).
        if (rpState.selected) loadClips(rpState.selected, rpState.gen);
      });
    });
    loadModels();
  }

  function deactivate() {
    rpState.active = false;
    stopReportPoll();
    stopTaskPoll();
    var vid = dom.player ? dom.player.querySelector("video") : null;
    if (vid) vid.pause();
  }

  function resize() {
    // Flow layout only — nothing measures the viewport.
  }

  document.addEventListener("visibilitychange", function () {
    if (!document.hidden && rpState.active) checkStaleness();
  });

  window.addEventListener("pagehide", function () {
    stopReportPoll();
    stopTaskPoll();
  });

  window.ClipgenOverview.reportsActivate = activate;
  window.ClipgenOverview.reportsDeactivate = deactivate;
  window.ClipgenOverview.reportsResize = resize;
})();
