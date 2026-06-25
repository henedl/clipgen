/* Workflows — run/results panel (satellite of workflows.js).
 *
 * Owns the Run lifecycle on the client: POST a run, subscribe to its per-run SSE
 * stream (with a polling fallback), render the run-history list + per-node status
 * rows in #wfRuns, and tint the canvas node cards by status with a progress bar.
 * Mirrors the screenspace-tasks SSE+poller pattern. Reads shared state through
 * WF.state — never re-`var`s a divergent `state` (the carve gotcha).
 */

(function () {
  "use strict";

  var WF = window.ClipgenWorkflows;
  var state = WF.state;

  // Satellite-local transport handles (only this file touches them, so they stay
  // module-local rather than on WF.state).
  var _stream = null; // EventSource for the active run
  var _poller = null; // createPoller fallback when SSE drops

  var TERMINAL = { completed: 1, failed: 1, cancelled: 1 };

  function isTerminal(status) {
    return !!TERMINAL[status];
  }

  // ---- Transport ------------------------------------------------------------

  function stopStream() {
    if (_stream) {
      _stream.close();
      _stream = null;
    }
  }

  function stopPolling() {
    if (_poller) {
      _poller.stop();
      _poller = null;
    }
  }

  function stopTransport() {
    stopStream();
    stopPolling();
  }

  function subscribeRun(runId) {
    stopTransport();
    if (!window.EventSource) {
      startPolling(runId);
      return;
    }
    var es = new EventSource("api/runs/" + encodeURIComponent(runId) + "/stream");
    _stream = es;
    es.onmessage = function (e) {
      var data;
      try {
        data = JSON.parse(e.data);
      } catch (_) {
        return;
      }
      if (data && data.run) handleRunData(data.run);
    };
    es.onerror = function () {
      // SSE dropped — fall back to polling so progress still flows.
      stopStream();
      startPolling(runId);
    };
  }

  function startPolling(runId) {
    if (_poller) return;
    _poller = createPoller(
      function () {
        apiGet("api/runs/" + encodeURIComponent(runId))
          .then(function (res) {
            if (res && res.run) handleRunData(res.run);
          })
          .catch(function () {});
      },
      POLL_INTERVAL,
      { runImmediately: true },
    );
    _poller.start();
  }

  // ---- Run lifecycle --------------------------------------------------------

  function activeRunInFlight() {
    var run = findRun(state.activeRunId);
    return run && !isTerminal(run.status);
  }

  function startRun() {
    if (!state.ready || !state.activeBlueprintId) return;
    if (activeRunInFlight()) return; // one run at a time
    setRunningUI(true);
    // Flush pending canvas edits so the server runs the latest blueprint.
    Promise.resolve(WF.flushSave ? WF.flushSave() : null)
      .then(function () {
        return apiPost("api/runs", { blueprintId: state.activeBlueprintId });
      })
      .then(function (res) {
        if (!res || !res.ok || !res.run) {
          setRunningUI(false);
          showToast("Couldn't start run");
          return;
        }
        upsertRun(res.run);
        state.activeRunId = res.run.id;
        renderRuns();
        subscribeRun(res.run.id);
      })
      .catch(function (err) {
        setRunningUI(false);
        // A cycle (or missing blueprint) comes back as a non-2xx → fetch throws.
        showToast((err && err.message) || "Couldn't start run");
      });
  }

  function stopRun() {
    var id = state.activeRunId;
    if (!id) return;
    apiPost("api/runs/" + encodeURIComponent(id) + "/cancel", {}).catch(
      function () {},
    );
    // UI flips to idle when the stream/poll reports the cancelled status.
  }

  // Fetch this blueprint's run history (and reattach to an in-flight run).
  function refreshRuns() {
    var bpId = state.activeBlueprintId;
    if (!bpId) return;
    stopTransport();
    state.activeRunId = null;
    apiGet("api/runs?blueprintId=" + encodeURIComponent(bpId))
      .then(function (res) {
        state.runs = (res && res.runs) || [];
        var live = null;
        for (var i = 0; i < state.runs.length; i++) {
          if (!isTerminal(state.runs[i].status)) {
            live = state.runs[i];
            break;
          }
        }
        if (live) {
          state.activeRunId = live.id;
          setRunningUI(true);
          subscribeRun(live.id);
        } else {
          setRunningUI(false);
        }
        renderRuns();
        annotateCanvas(state.runs[0] || null);
      })
      .catch(function () {});
  }

  // ---- Data handling --------------------------------------------------------

  function findRun(id) {
    for (var i = 0; i < state.runs.length; i++) {
      if (state.runs[i].id === id) return state.runs[i];
    }
    return null;
  }

  function upsertRun(run) {
    var existing = findRun(run.id);
    if (existing) {
      var idx = state.runs.indexOf(existing);
      state.runs[idx] = run;
    } else {
      state.runs.unshift(run);
    }
  }

  function handleRunData(run) {
    if (!run) return;
    upsertRun(run);
    if (run.id === state.activeRunId && isTerminal(run.status)) {
      stopTransport();
      setRunningUI(false);
    }
    renderRuns();
    annotateCanvas(run);
  }

  // ---- Canvas tinting -------------------------------------------------------

  var NODE_RUN_CLASSES = [
    "run-queued",
    "run-running",
    "run-completed",
    "run-failed",
    "run-skipped",
  ];

  // Toggle per-node status classes + a progress bar on the canvas cards. Only
  // tints when the run belongs to the blueprint currently on the canvas (a stale
  // run from another blueprint must not paint these cards).
  function annotateCanvas(run) {
    var cards = qsa("#wfWorld .wf-node");
    if (!run || run.blueprintId !== state.activeBlueprintId) {
      for (var c = 0; c < cards.length; c++) clearNodeRun(cards[c]);
      return;
    }
    var states = run.nodeStates || {};
    for (var i = 0; i < cards.length; i++) {
      var card = cards[i];
      var ns = states[card.getAttribute("data-node-id")];
      if (!ns) {
        clearNodeRun(card);
        continue;
      }
      card.classList.remove.apply(card.classList, NODE_RUN_CLASSES);
      card.classList.add("run-" + ns.status);
      setNodeProgress(card, ns.status === "running" ? ns.progress || 0 : 0);
    }
  }

  function clearNodeRun(card) {
    card.classList.remove.apply(card.classList, NODE_RUN_CLASSES);
    setNodeProgress(card, 0);
  }

  function setNodeProgress(card, fraction) {
    var bar = card.querySelector(".wf-node-progress");
    if (fraction > 0) {
      if (!bar) {
        bar = el("div", "wf-node-progress");
        bar.appendChild(el("div", "wf-node-progress-fill"));
        card.appendChild(bar);
      }
      bar.firstChild.style.width = Math.round(fraction * 100) + "%";
    } else if (bar) {
      bar.parentNode.removeChild(bar);
    }
  }

  // ---- Rendering ------------------------------------------------------------

  // Build a label map for the active blueprint's nodes (run rows show the node's
  // catalog label, falling back to its id for a run from another blueprint).
  function nodeLabel(nodeId) {
    var nodes = state.nodes || [];
    for (var i = 0; i < nodes.length; i++) {
      if (nodes[i].id === nodeId) {
        var type = state.catalogById[nodes[i].type];
        return (type && type.label) || nodes[i].type;
      }
    }
    return nodeId;
  }

  function statusCounts(run) {
    var counts = {};
    var states = run.nodeStates || {};
    Object.keys(states).forEach(function (k) {
      var s = states[k].status;
      counts[s] = (counts[s] || 0) + 1;
    });
    return counts;
  }

  function buildResultChips(run) {
    // Surface terminal pointers/counts (viewer path, artifact/event counts).
    var results = run.results || {};
    var chips = el("div", "wf-run-results");
    var any = false;
    Object.keys(results).forEach(function (nodeId) {
      var ports = results[nodeId] || {};
      Object.keys(ports).forEach(function (port) {
        var val = ports[port];
        if (val == null) return;
        var text = null;
        if (typeof val === "object") {
          if (val.path) text = port + ": " + basename(val.path);
          else if (typeof val.count === "number") text = port + ": " + val.count;
        } else if (typeof val !== "boolean") {
          text = port + ": " + val;
        }
        if (text) {
          chips.appendChild(el("span", "wf-run-chip", text));
          any = true;
        }
      });
    });
    return any ? chips : null;
  }

  function basename(path) {
    var parts = String(path).split(/[\\/]/);
    return parts[parts.length - 1] || path;
  }

  function buildRunCard(run, expanded) {
    var card = el("div", "wf-run-card");
    card.dataset.runId = run.id;

    var head = el("div", "wf-run-head");
    head.appendChild(el("span", "wf-run-status wf-run-status-" + run.status, run.status));
    var counts = statusCounts(run);
    var total = Object.keys(run.nodeStates || {}).length;
    var done = (counts.completed || 0) + (counts.skipped || 0);
    head.appendChild(el("span", "wf-run-meta", done + "/" + total + " nodes"));
    card.appendChild(head);

    if (expanded) {
      var rows = el("div", "wf-run-nodes");
      var states = run.nodeStates || {};
      Object.keys(states).forEach(function (nodeId) {
        var ns = states[nodeId];
        var row = el("div", "wf-run-node wf-run-node-" + ns.status);
        row.appendChild(el("span", "wf-run-node-label", nodeLabel(nodeId)));
        var detail =
          ns.status === "running"
            ? Math.round((ns.progress || 0) * 100) + "%"
            : ns.status;
        row.appendChild(el("span", "wf-run-node-detail", detail));
        if (ns.error) row.title = ns.error;
        rows.appendChild(row);
      });
      card.appendChild(rows);

      var chips = buildResultChips(run);
      if (chips) card.appendChild(chips);
    }
    return card;
  }

  function renderRuns() {
    var container = qs("#wfRuns");
    if (!container) return;
    container.innerHTML = "";
    if (!state.runs.length) {
      container.appendChild(
        el(
          "p",
          "wf-empty-hint",
          "Press Run to execute this workflow. Per-node progress and results appear here.",
        ),
      );
      return;
    }
    var frag = document.createDocumentFragment();
    // The newest (or active) run shows full per-node detail; the rest are compact.
    state.runs.forEach(function (run, idx) {
      frag.appendChild(buildRunCard(run, idx === 0));
    });
    container.appendChild(frag);
  }

  // ---- Running UI -----------------------------------------------------------

  function setRunningUI(running) {
    var runBtn = qs("#wfRunBtn");
    var stopBtn = qs("#wfStopBtn");
    if (runBtn) runBtn.disabled = running || !state.ready;
    if (stopBtn) stopBtn.classList.toggle("hidden", !running);
  }

  // Pause/resume the live stream when the tab is hidden (the poller already
  // self-pauses; the EventSource is reopened on return if a run is in flight).
  function onVisibility() {
    if (document.hidden) {
      stopStream();
      return;
    }
    if (activeRunInFlight() && !_stream && !_poller) {
      subscribeRun(state.activeRunId);
    }
  }

  function initRuns() {
    document.addEventListener("visibilitychange", onVisibility);
    setRunningUI(false);
  }

  // ---- Satellite interface ----
  WF.initRuns = initRuns;
  WF.startRun = startRun;
  WF.stopRun = stopRun;
  WF.refreshRuns = refreshRuns;
  WF.renderRuns = renderRuns;
})();
