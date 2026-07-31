/* Overview Summary tab — per-participant mini-report (overview-summary.js).
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
 * All hub data comes from the overview.js hub via window.ClipgenOverview
 * (lazy reads inside activate(), never top-level destructures). Lifecycle:
 * OV.summaryActivate / summaryDeactivate / summaryResize. Participant
 * selection is a single key on purpose — a future aggregate mode extends
 * smState.selected to a set without reshaping the tab.
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

  var smState = {
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
  };

  var dom = {}; // scaffold refs, filled by initDom()

  // ---- Data helpers ----

  function rec() {
    for (var i = 0; i < smState.participants.length; i++) {
      if (smState.participants[i].id === smState.selected) {
        return smState.participants[i];
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
    for (var i = 0; i < smState.participants.length; i++) {
      var p = smState.participants[i];
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
    var o = smState.ollama;
    if (!o) return null;
    if (o.available === false) {
      return "Ollama is not reachable at " + (o.base_url || "localhost") + ". Start it, then Refresh.";
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
    if (smState.initialized) return;
    smState.initialized = true;
    var panel = qs("#summaryPanel");
    if (!panel) return;

    var layout = el("div", "sm-layout");

    var sidebar = el("aside", "sm-sidebar");
    sidebar.appendChild(el("div", "sm-sidebar-header", "Participants"));
    dom.pills = el("div", "sm-pills");
    sidebar.appendChild(dom.pills);
    layout.appendChild(sidebar);

    var main = el("section", "sm-main");
    dom.empty = el("div", "sm-empty hidden");
    dom.empty.appendChild(el("p", "", "No participants yet."));
    var emptySub = el("p", "sm-empty-sub");
    emptySub.appendChild(document.createTextNode("Add videos to the input folder or open a spreadsheet in "));
    var studioLink = el("a", "", "Studio");
    studioLink.href = "/studio/";
    emptySub.appendChild(studioLink);
    emptySub.appendChild(document.createTextNode(", then come back."));
    dom.empty.appendChild(emptySub);
    main.appendChild(dom.empty);

    dom.content = el("div", "sm-content");
    dom.sources = el("div", "sm-sources");
    dom.content.appendChild(dom.sources);
    dom.note = el("div", "sm-note hidden");
    dom.content.appendChild(dom.note);

    var report = el("div", "sm-report");
    var head = el("div", "sm-report-head");
    head.appendChild(el("h2", "sm-report-title", "Mini-report"));
    dom.actions = el("div", "sm-report-actions");
    head.appendChild(dom.actions);
    report.appendChild(head);
    dom.meta = el("div", "sm-report-meta");
    report.appendChild(dom.meta);
    dom.body = el("div", "sm-report-body");
    report.appendChild(dom.body);
    dom.content.appendChild(report);

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
    for (var i = 0; i < smState.participants.length; i++) {
      var p = smState.participants[i];
      var pill = P.createParticipantPill({
        id: p.id,
        active: p.id === smState.selected,
        dataset: { pid: p.id },
        onClick: onPillClick,
      });
      var reportState = agentState(p, "report");
      if (reportState === "done" || reportState === "running") {
        pill.appendChild(el("span", "sm-pill-dot is-" + reportState));
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
    var hasParticipants = smState.participants.length > 0;
    dom.empty.classList.toggle("hidden", hasParticipants);
    dom.content.classList.toggle("hidden", !hasParticipants);
    if (!hasParticipants) return;
    renderSources();
    renderReportArea();
  }

  function sourceRow(labelText, status, kind, actionEl) {
    var row = el("div", "sm-source-row");
    row.appendChild(el("span", "sm-source-dot is-" + kind));
    row.appendChild(el("span", "sm-source-label", labelText));
    var statusEl = el("span", "sm-source-status");
    if (typeof status === "string") statusEl.textContent = status;
    else statusEl.appendChild(status);
    row.appendChild(statusEl);
    if (actionEl) {
      var act = el("span", "sm-source-action");
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
    var pid = smState.selected;
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
      frag.appendChild(sourceRow(
        "Transcript summary", "no summary", "warn",
        P.createBtn({ label: "Generate summary", icon: "sparkles", size: "sm", onClick: generateSummary })
      ));
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

  // Minimal markdown for the report shape the prompt asks for (## headings,
  // "- " bullets, paragraphs). Everything passes through escapeHtml; the model
  // never gets to inject markup.
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
        html += "<" + tag + ">" + escapeHtml(heading[2]) + "</" + tag + ">";
      } else if (line.indexOf("- ") === 0 || line.indexOf("* ") === 0) {
        if (!inList) { html += "<ul>"; inList = true; }
        html += "<li>" + escapeHtml(line.substring(2)) + "</li>";
      } else {
        if (inList) { html += "</ul>"; inList = false; }
        html += "<p>" + escapeHtml(line) + "</p>";
      }
    }
    if (inList) html += "</ul>";
    return html;
  }

  function renderReportBodyPartial() {
    dom.body.innerHTML = "";
    var pre = el("div", "sm-report-partial");
    pre.textContent = smState.reportPartial || "Waiting for the model…";
    dom.body.appendChild(pre);
  }

  function renderReportArea() {
    if (!dom.actions) return;
    dom.actions.innerHTML = "";
    dom.meta.textContent = "";
    var pid = smState.selected;
    var r = rec();
    if (!pid || !r) {
      dom.body.innerHTML = "";
      return;
    }
    var P = window.ClipgenPrimitives;

    if (smState.reportGenerating) {
      dom.actions.appendChild(P.createBtn({ label: "Stop", icon: "stop", size: "sm", onClick: stopReport }));
      dom.meta.textContent = "Generating…";
      renderReportBodyPartial();
      return;
    }

    var canGenerate = !!r.has_transcript && agentState(r, "summary") === "done";
    var gate = ollamaGate();
    dom.actions.appendChild(P.createBtn({
      label: smState.report ? "Regenerate" : "Generate report",
      icon: "sparkles",
      variant: "solid",
      size: "sm",
      disabled: !canGenerate || !!gate,
      onClick: generateReport,
    }));

    if (smState.report) {
      dom.meta.textContent = metaLine(smState.report);
      dom.body.innerHTML = renderReportText(smState.report.text);
      return;
    }

    var hint;
    if (!canGenerate) {
      hint = "Needs a transcript summary first — trigger the missing steps above.";
    } else if (gate) {
      hint = gate;
    } else if (smState.reportMissing) {
      hint = "No report yet. Generate one from the sources above.";
    } else {
      hint = "Loading…";
    }
    dom.body.innerHTML = "";
    dom.body.appendChild(el("p", "sm-report-hint", hint));
  }

  // ---- Notices (transcribe download gate, trigger failures) ----

  function showNotice(text, actionLabel, onAction, linkPid) {
    dom.note.innerHTML = "";
    dom.note.classList.remove("hidden");
    dom.note.appendChild(el("span", "sm-note-text", text));
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
    if (smState.selected === pid) return;
    smState.selected = pid;
    smState.gen++;
    stopReportPoll();
    clearNotice();
    smState.report = null;
    smState.reportGenerating = false;
    smState.reportPartial = "";
    smState.reportMissing = false;
    renderAll();
    fetchReport(pid, smState.gen);
  }

  function ensureSelection() {
    var ids = {};
    for (var i = 0; i < smState.participants.length; i++) ids[smState.participants[i].id] = true;
    if (smState.selected && !ids[smState.selected]) smState.selected = null;
    if (!smState.selected && smState.participants.length > 0) {
      selectParticipant(smState.participants[0].id);
    }
  }

  function loadParticipants() {
    return apiGet("../transcripts/api/participants")
      .then(function (data) {
        if (!smState.active) return;
        smState.participants = mergeSheetParticipants((data && data.participants) || []);
        ensureSelection();
        renderAll();
      })
      .catch(function () {
        if (!smState.active) return;
        smState.participants = mergeSheetParticipants([]);
        ensureSelection();
        renderAll();
      });
  }

  function loadModels() {
    apiGet("/api/models")
      .then(function (data) {
        smState.ollama = (data && data.ollama) || null;
        if (smState.active) renderMain();
      })
      .catch(function () {
        smState.ollama = null;
      });
  }

  // ---- Report fetch/poll ----

  function applyReportResponse(data, pid, g) {
    if (data && data.ok && data.report) {
      stopReportPoll();
      smState.report = data.report;
      smState.reportGenerating = false;
      smState.reportPartial = "";
      smState.reportMissing = false;
      renderReportArea();
      // The pill dot (agents.report) is served by the participants API; a
      // local nudge keeps it honest until the next refetch.
      var r = rec();
      if (r && r.agents) r.agents.report = "done";
      renderParticipants();
    } else if (data && data.generating) {
      smState.reportGenerating = true;
      smState.reportPartial = data.partial || "";
      if (!smState.reportPoll) {
        renderReportArea();
        startReportPoll(pid, g);
      } else {
        renderReportBodyPartial();
      }
    } else {
      stopReportPoll();
      smState.reportGenerating = false;
      smState.reportMissing = true;
      renderReportArea();
    }
  }

  function fetchReport(pid, g) {
    apiGet("../transcripts/api/agent/report/" + pid)
      .then(function (data) {
        if (g !== smState.gen || !smState.active) return;
        applyReportResponse(data, pid, g);
      })
      .catch(function () {
        // 404 = no report stored for this participant (apiGet throws on it).
        if (g !== smState.gen || !smState.active) return;
        stopReportPoll();
        smState.reportGenerating = false;
        smState.reportMissing = true;
        renderReportArea();
      });
  }

  function startReportPoll(pid, g) {
    stopReportPoll();
    smState.reportPoll = setInterval(function () {
      if (!smState.active || g !== smState.gen) {
        stopReportPoll();
        return;
      }
      if (document.hidden) return;
      fetchReport(pid, g);
    }, REPORT_POLL_MS);
  }

  function stopReportPoll() {
    if (smState.reportPoll) {
      clearInterval(smState.reportPoll);
      smState.reportPoll = null;
    }
  }

  function generateReport() {
    var pid = smState.selected;
    if (!pid) return;
    var g = smState.gen;
    clearNotice();
    apiPost("../transcripts/api/agent/report/" + pid + "/regenerate", {})
      .then(function () {
        if (g !== smState.gen || !smState.active) return;
        smState.report = null;
        smState.reportGenerating = true;
        smState.reportPartial = "";
        renderReportArea();
        startReportPoll(pid, g);
      })
      .catch(function () {
        if (g !== smState.gen || !smState.active) return;
        showNotice("Could not start the report — re-check the sources above.", null, null, null);
      });
  }

  function stopReport() {
    var pid = smState.selected;
    if (!pid) return;
    var g = smState.gen;
    apiPost("../transcripts/api/agent/report/" + pid + "/stop", {})
      .then(function () {
        if (g !== smState.gen || !smState.active) return;
        stopReportPoll();
        smState.reportGenerating = false;
        smState.reportPartial = "";
        fetchReport(pid, g);
      })
      .catch(function () {});
  }

  // ---- Upstream triggers (transcription, summary) ----

  function generateSummary() {
    var pid = smState.selected;
    if (!pid) return;
    clearNotice();
    apiPost("../transcripts/api/agent/summary/" + pid + "/regenerate", {})
      .then(function () {
        if (!smState.active) return;
        var r = rec();
        if (r && r.agents) r.agents.summary = "running";
        renderMain();
        startTaskPoll();
      })
      .catch(function () {
        if (!smState.active) return;
        showNotice("Could not start the summary — is there a transcript?", null, null, pid);
      });
  }

  function transcribe(allowDownload) {
    var pid = smState.selected;
    if (!pid) return;
    clearNotice();
    apiPost("../transcripts/api/transcribe", { participants: [pid], allow_download: !!allowDownload })
      .then(function (data) {
        if (!smState.active) return;
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
        if (!smState.active) return;
        showNotice("Could not start transcription.", null, null, pid);
      });
  }

  // While a triggered upstream task runs, re-poll the participants list (a
  // quiet-poll path) so the source rows track transcription → summary → done,
  // and the Generate button unlocks the moment the summary lands.
  function startTaskPoll() {
    smState.taskIdleTicks = 0;
    if (smState.taskPoll) return;
    smState.taskPoll = setInterval(function () {
      if (!smState.active) {
        stopTaskPoll();
        return;
      }
      if (document.hidden) return;
      apiGet("../transcripts/api/participants")
        .then(function (data) {
          if (!smState.active) return;
          smState.participants = mergeSheetParticipants((data && data.participants) || []);
          ensureSelection();
          renderAll();
          if (anyUpstreamRunning()) {
            smState.taskIdleTicks = 0;
          } else {
            smState.taskIdleTicks++;
            if (smState.taskIdleTicks >= TASK_IDLE_TICKS_TO_STOP) stopTaskPoll();
          }
        })
        .catch(function () {});
    }, TASK_POLL_MS);
  }

  function stopTaskPoll() {
    if (smState.taskPoll) {
      clearInterval(smState.taskPoll);
      smState.taskPoll = null;
    }
  }

  // ---- Staleness (hub dataVersion contract) ----

  function takeSnapshot() {
    smState._snapshot = { version: state.dataVersion };
  }

  function checkStaleness() {
    if (!smState._snapshot || !smState.active) return;
    window.ClipgenOverview.setRefreshStale(smState._snapshot.version !== state.dataVersion);
  }

  // ---- Lifecycle ----

  function activate() {
    smState.active = true;
    if (!state) state = window.ClipgenOverview.state;
    initDom();
    if (smState._snapshot) checkStaleness();
    window.ClipgenOverview.ensureData().then(function () {
      if (!smState.active) return;
      loadParticipants().then(function () {
        if (!smState.active) return;
        takeSnapshot();
        checkStaleness();
      });
    });
    loadModels();
  }

  function deactivate() {
    smState.active = false;
    stopReportPoll();
    stopTaskPoll();
  }

  function resize() {
    // Flow layout only — nothing measures the viewport.
  }

  document.addEventListener("visibilitychange", function () {
    if (!document.hidden && smState.active) checkStaleness();
  });

  window.addEventListener("pagehide", function () {
    stopReportPoll();
    stopTaskPoll();
  });

  window.ClipgenOverview.summaryActivate = activate;
  window.ClipgenOverview.summaryDeactivate = deactivate;
  window.ClipgenOverview.summaryResize = resize;
})();
