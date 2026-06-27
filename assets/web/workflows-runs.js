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
  // module-local rather than on WF.state). The run pair streams the focused child
  // (or a single run); the batch pair streams the batch summary — both can run at
  // once during a batch (summary + the drilled-in child's per-node detail).
  var _stream = null; // EventSource for the active/focused run
  var _poller = null; // createPoller fallback when run SSE drops
  var _batchStream = null; // EventSource for the active batch
  var _batchPoller = null; // createPoller fallback when batch SSE drops
  var _discoverPoller = null; // low-freq poll surfacing runs THIS client didn't start

  var TERMINAL = { completed: 1, failed: 1, cancelled: 1 };

  function isTerminal(status) {
    return !!TERMINAL[status];
  }

  // A blueprint fans out when any Video Source is set to "All participants" — the
  // single Run button then launches a batch instead of one run.
  function blueprintWantsBatch() {
    var nodes = state.nodes || [];
    for (var i = 0; i < nodes.length; i++) {
      if (
        nodes[i].type === "video_source" &&
        nodes[i].params &&
        nodes[i].params.participant === WF.ALL_PARTICIPANTS
      ) {
        return true;
      }
    }
    return false;
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

  // ---- Batch transport ------------------------------------------------------

  function stopBatchStream() {
    if (_batchStream) {
      _batchStream.close();
      _batchStream = null;
    }
  }

  function stopBatchPolling() {
    if (_batchPoller) {
      _batchPoller.stop();
      _batchPoller = null;
    }
  }

  function stopBatchTransport() {
    stopBatchStream();
    stopBatchPolling();
  }

  function subscribeBatch(batchId) {
    stopBatchTransport();
    if (!window.EventSource) {
      startBatchPolling(batchId);
      return;
    }
    var es = new EventSource(
      "api/batches/" + encodeURIComponent(batchId) + "/stream",
    );
    _batchStream = es;
    es.onmessage = function (e) {
      var data;
      try {
        data = JSON.parse(e.data);
      } catch (_) {
        return;
      }
      if (data && data.batch) handleBatchData(data.batch);
    };
    es.onerror = function () {
      stopBatchStream();
      startBatchPolling(batchId);
    };
  }

  function startBatchPolling(batchId) {
    if (_batchPoller) return;
    _batchPoller = createPoller(
      function () {
        apiGet("api/batches/" + encodeURIComponent(batchId))
          .then(function (res) {
            if (res && res.batch) handleBatchData(res.batch);
          })
          .catch(function () {});
      },
      POLL_INTERVAL,
      { runImmediately: true },
    );
    _batchPoller.start();
  }

  // ---- Run lifecycle --------------------------------------------------------

  function activeRunInFlight() {
    var run = findRun(state.activeRunId);
    return run && !isTerminal(run.status);
  }

  function activeBatchInFlight() {
    var batch = findBatch(state.activeBatchId);
    return batch && !isTerminal(batch.status);
  }

  function startRun(targetNodeId) {
    if (!state.ready || !state.activeBlueprintId) return;
    if (activeRunInFlight() || activeBatchInFlight()) return; // one at a time
    // Errors gate the run (the button is already disabled; this guards the
    // programmatic path). Warnings never block.
    if (state.validation && state.validation.errors.length) return;
    // A Video Source set to "All participants" makes Run fan out over the study —
    // but a partial "run to here" is always a single run.
    if (!targetNodeId && blueprintWantsBatch()) {
      startBatch();
      return;
    }
    setRunningUI(true);
    // Flush pending canvas edits so the server runs the latest blueprint.
    Promise.resolve(WF.flushSave ? WF.flushSave() : null)
      .then(function () {
        var body = { blueprintId: state.activeBlueprintId };
        if (targetNodeId) body.targetNodeId = targetNodeId;
        return apiPost("api/runs", body);
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
    // The single Stop button cancels whichever is in flight — a batch cancels the
    // whole fan-out (current child + remaining), else a single run.
    if (activeBatchInFlight()) {
      apiPost(
        "api/batches/" + encodeURIComponent(state.activeBatchId) + "/cancel",
        {},
      ).catch(function () {});
      return;
    }
    var id = state.activeRunId;
    if (!id) return;
    apiPost("api/runs/" + encodeURIComponent(id) + "/cancel", {}).catch(
      function () {},
    );
    // UI flips to idle when the stream/poll reports the cancelled status.
  }

  // Fan the active blueprint out across every participant (P3). One run per
  // participant, sequential, grouped under one batch card. Reached from startRun
  // when a Video Source is set to "All participants".
  function startBatch() {
    if (!state.ready || !state.activeBlueprintId) return;
    if (activeRunInFlight() || activeBatchInFlight()) return; // one at a time
    setRunningUI(true);
    Promise.resolve(WF.flushSave ? WF.flushSave() : null)
      .then(function () {
        return apiPost("api/batches", { blueprintId: state.activeBlueprintId });
      })
      .then(function (res) {
        if (!res || !res.ok || !res.batch) {
          setRunningUI(false);
          showToast("Couldn't start batch");
          return;
        }
        upsertBatch(res.batch);
        state.activeBatchId = res.batch.id;
        state.activeRunId = null;
        renderRuns();
        subscribeBatch(res.batch.id);
      })
      .catch(function (err) {
        setRunningUI(false);
        showToast((err && err.message) || "Couldn't start batch");
      });
  }

  // Fetch this blueprint's run + batch history (and reattach to in-flight work).
  function refreshRuns() {
    var bpId = state.activeBlueprintId;
    if (!bpId) return;
    stopTransport();
    stopBatchTransport();
    state.activeRunId = null;
    state.activeBatchId = null;
    var q = encodeURIComponent(bpId);
    Promise.all([
      apiGet("api/runs?blueprintId=" + q).catch(function () {
        return null;
      }),
      apiGet("api/batches?blueprintId=" + q).catch(function () {
        return null;
      }),
    ])
      .then(function (results) {
        state.runs = (results[0] && results[0].runs) || [];
        state.batches = (results[1] && results[1].batches) || [];

        // A live batch owns the Run/Stop buttons; reattach to it first and focus
        // its running child for the canvas tint.
        var liveBatch = firstNonTerminal(state.batches);
        if (liveBatch) {
          state.activeBatchId = liveBatch.id;
          setRunningUI(true);
          subscribeBatch(liveBatch.id);
          var runningChild = (liveBatch.children || []).filter(function (c) {
            return c.status === "running";
          })[0];
          renderRuns();
          if (runningChild) focusChild(runningChild.runId);
          else annotateCanvas(null);
          return;
        }

        // Else reattach to a loose (non-batch) in-flight run, as before.
        var looseRuns = state.runs.filter(function (r) {
          return !r.batchId;
        });
        var live = null;
        for (var i = 0; i < looseRuns.length; i++) {
          if (!isTerminal(looseRuns[i].status)) {
            live = looseRuns[i];
            break;
          }
        }
        state.activeRunId = live
          ? live.id
          : (looseRuns[0] && looseRuns[0].id) || null;
        if (live) {
          setRunningUI(true);
          subscribeRun(live.id);
        } else {
          setRunningUI(false);
        }
        renderRuns();
        annotateCanvas(findRun(state.activeRunId));
      })
      .catch(function () {});
  }

  function firstNonTerminal(list) {
    for (var i = 0; i < (list || []).length; i++) {
      if (!isTerminal(list[i].status)) return list[i];
    }
    return null;
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
      // During a batch the batch summary owns the Run/Stop buttons — a finished
      // child must not flip them back to idle while siblings are still running.
      if (!activeBatchInFlight()) setRunningUI(false);
    }
    renderRuns();
    annotateCanvas(run);
  }

  // ---- Batch data handling --------------------------------------------------

  function findBatch(id) {
    for (var i = 0; i < state.batches.length; i++) {
      if (state.batches[i].id === id) return state.batches[i];
    }
    return null;
  }

  function upsertBatch(batch) {
    var existing = findBatch(batch.id);
    if (existing) {
      state.batches[state.batches.indexOf(existing)] = batch;
    } else {
      state.batches.unshift(batch);
    }
  }

  function handleBatchData(batch) {
    if (!batch) return;
    upsertBatch(batch);
    if (batch.id === state.activeBatchId && isTerminal(batch.status)) {
      stopBatchTransport();
      setRunningUI(false);
    }
    renderRuns();
  }

  // Drill into one participant's run: stream its per-node detail + tint the canvas
  // by it, without touching the batch's ownership of the Run/Stop buttons.
  function focusChild(runId) {
    if (!runId) return;
    state.activeRunId = runId;
    var existing = findRun(runId);
    if (existing) annotateCanvas(existing);
    renderRuns();
    subscribeRun(runId); // live per-node updates if the child is still running
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

  // Per-node detail rows + result chips for an expanded run (shared by single-run
  // cards and a drilled-in batch child). Rows whose snapshot says `hasResult`
  // expand to lazily fetch + render the node's stored result sidecar (P5).
  function buildNodeDetail(run) {
    var wrap = document.createDocumentFragment();
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
      if (ns.hasResult) {
        row.classList.add("wf-run-node-expandable");
        var panel = el("div", "wf-result-panel hidden");
        row.addEventListener("click", function (e) {
          e.stopPropagation();
          toggleNodeResult(run, nodeId, panel);
        });
        rows.appendChild(panel);
      }
    });
    wrap.appendChild(rows);
    var chips = buildResultChips(run);
    if (chips) wrap.appendChild(chips);
    return wrap;
  }

  // ---- Lazy per-node result (P5) -------------------------------------------

  // Expand/collapse one node's result panel. The full payload is fetched once
  // and cached on the run object, so re-expanding (even after a card rebuild)
  // never re-hits the endpoint.
  function toggleNodeResult(run, nodeId, panel) {
    if (!panel.classList.contains("hidden")) {
      panel.classList.add("hidden");
      return;
    }
    panel.classList.remove("hidden");
    if (panel.dataset.loaded === "1" || panel.dataset.loading === "1") return;
    var cached = run._nodeResults && run._nodeResults[nodeId];
    if (cached) {
      panel.innerHTML = "";
      appendResultBody(panel, cached);
      panel.dataset.loaded = "1";
      return;
    }
    panel.dataset.loading = "1";
    panel.innerHTML = "";
    panel.appendChild(el("div", "wf-result-line", "Loading…"));
    apiGet(
      "api/runs/" +
        encodeURIComponent(run.id) +
        "/nodes/" +
        encodeURIComponent(nodeId) +
        "/result",
    )
      .then(function (res) {
        panel.dataset.loading = "";
        panel.innerHTML = "";
        if (res && res.ok && res.result) {
          run._nodeResults = run._nodeResults || {};
          run._nodeResults[nodeId] = res.result;
          appendResultBody(panel, res.result);
          panel.dataset.loaded = "1";
        } else {
          panel.appendChild(el("div", "wf-result-line", "No stored result."));
        }
      })
      .catch(function () {
        panel.dataset.loading = "";
        panel.innerHTML = "";
        panel.appendChild(el("div", "wf-result-line", "Failed to load result."));
      });
  }

  // Render the stored result (a {port: value} map) into `container`, branching on
  // the value's shape — artifacts/events/segments lists, reel manifest, viewer
  // path, summary/citation/friction text, scalar.
  function appendResultBody(container, result) {
    var ports = Object.keys(result || {});
    if (!ports.length) {
      container.appendChild(el("div", "wf-result-line", "No inspectable output."));
      return;
    }
    var frag = document.createDocumentFragment();
    ports.forEach(function (port) {
      var box = el("div", "wf-result-port");
      box.appendChild(el("span", "wf-result-port-name", port));
      var body = el("div", "wf-result-port-body");
      renderResultValue(body, result[port]);
      box.appendChild(body);
      frag.appendChild(box);
    });
    container.appendChild(frag);
  }

  function renderResultValue(body, val) {
    if (val == null) {
      body.appendChild(el("div", "wf-result-line", "—"));
    } else if (typeof val === "string") {
      body.appendChild(el("div", "wf-result-text", val));
    } else if (typeof val === "number" || typeof val === "boolean") {
      body.appendChild(el("div", "wf-result-line", String(val)));
    } else if (Array.isArray(val)) {
      renderList(body, val.map(citationLabel));
    } else if (Array.isArray(val.artifacts)) {
      countLine(body, val.artifacts.length, "artifact");
      renderList(
        body,
        val.artifacts.map(function (a) {
          return a.file || a.path || a.name || a.id;
        }),
      );
    } else if (Array.isArray(val.events)) {
      countLine(body, val.events.length, "event");
      renderList(body, val.events.map(eventLabel));
    } else if (Array.isArray(val.segments)) {
      countLine(body, val.segments.length, "segment");
      renderList(
        body,
        val.segments.map(function (s) {
          return (s && (s.text || s.label)) || "";
        }),
      );
    } else if (Array.isArray(val.records)) {
      countLine(body, val.records.length, "item");
      if (val.path) body.appendChild(el("div", "wf-result-line", basename(val.path)));
    } else if (val.path) {
      body.appendChild(el("div", "wf-result-line", basename(val.path)));
    } else {
      body.appendChild(el("div", "wf-result-text", JSON.stringify(val)));
    }
  }

  function countLine(body, n, noun) {
    body.appendChild(
      el("div", "wf-result-line", n + " " + noun + (n === 1 ? "" : "s")),
    );
  }

  function renderList(body, items) {
    items.slice(0, 8).forEach(function (it) {
      body.appendChild(el("div", "wf-result-item", it == null ? "—" : String(it)));
    });
    if (items.length > 8) {
      body.appendChild(el("div", "wf-result-more", "+" + (items.length - 8) + " more"));
    }
  }

  function fmtClock(sec) {
    var s = Math.max(0, Math.round(Number(sec) || 0));
    var m = Math.floor(s / 60);
    var r = s % 60;
    return m + ":" + (r < 10 ? "0" : "") + r;
  }

  function eventLabel(ev) {
    if (!ev || typeof ev !== "object") return String(ev);
    var t = ev.time_in != null ? fmtClock(ev.time_in) + "  " : "";
    return t + (ev.event_type || ev.detector || "event");
  }

  function citationLabel(item) {
    if (item == null || typeof item !== "object") return item;
    return item.claim || item.text || item.label || item.quote || JSON.stringify(item);
  }

  function buildRunCard(run, expanded) {
    var card = el("div", "wf-run-card");
    card.dataset.runId = run.id;

    var head = el("div", "wf-run-head");
    head.appendChild(el("span", "wf-run-status wf-run-status-" + run.status, run.status));
    // Auto-launched by the watch-dir trigger (P6) — a bolt chip distinguishes it
    // from a manual run (margin-right:auto keeps it hugging the status label).
    if (run.triggered) {
      head.appendChild(el("span", "wf-run-triggered", "triggered"));
    }
    var counts = statusCounts(run);
    var total = Object.keys(run.nodeStates || {}).length;
    var done = (counts.completed || 0) + (counts.skipped || 0);
    head.appendChild(el("span", "wf-run-meta", done + "/" + total + " nodes"));
    // Re-run a terminal run of the active blueprint — relaunches the same graph
    // (no partial/memoized re-run; the engine just re-executes). Disabled while
    // a run/batch is in flight, mirroring the toolbar Run gate.
    if (isTerminal(run.status) && run.blueprintId === state.activeBlueprintId) {
      var rerun = el("button", "wf-run-rerun", "Re-run");
      rerun.type = "button";
      rerun.disabled = activeRunInFlight() || activeBatchInFlight();
      rerun.addEventListener("click", function (e) {
        e.stopPropagation();
        startRun();
      });
      head.appendChild(rerun);
    }
    card.appendChild(head);

    if (expanded) card.appendChild(buildNodeDetail(run));
    return card;
  }

  // ---- Batch rendering ------------------------------------------------------

  function batchCounts(batch) {
    return batch.counts || {};
  }

  function buildBatchCard(batch, expanded) {
    var card = el("div", "wf-batch-card");
    card.dataset.batchId = batch.id;

    var head = el("div", "wf-run-head");
    head.appendChild(
      el("span", "wf-run-status wf-run-status-" + batch.status, batch.status),
    );
    var counts = batchCounts(batch);
    var total = (batch.children || []).length;
    var done = counts.completed || 0;
    var parts = [done + "/" + total + " done"];
    if (counts.failed) parts.push(counts.failed + " failed");
    if (counts.cancelled) parts.push(counts.cancelled + " cancelled");
    head.appendChild(el("span", "wf-run-meta", "All participants · " + parts.join(" · ")));
    card.appendChild(head);

    if (expanded) {
      var rows = el("div", "wf-batch-children");
      (batch.children || []).forEach(function (child) {
        var focused = child.runId === state.activeRunId;
        var row = el(
          "div",
          "wf-batch-child wf-run-node-" + child.status + (focused ? " focused" : ""),
        );
        row.appendChild(
          el("span", "wf-batch-child-label", child.participant || child.runId),
        );
        row.appendChild(el("span", "wf-run-node-detail", child.status));
        row.addEventListener("click", function () {
          focusChild(child.runId);
        });
        rows.appendChild(row);
        // The focused child expands inline with its per-node detail (if its full
        // snapshot has streamed in via focusChild → subscribeRun).
        var run = focused ? findRun(child.runId) : null;
        if (run && run.nodeStates) rows.appendChild(buildNodeDetail(run));
      });
      card.appendChild(rows);
    }
    return card;
  }

  function renderRuns() {
    var container = qs("#wfRuns");
    if (!container) return;
    container.innerHTML = "";
    // Batch children are surfaced inside their batch card, not as loose runs.
    var looseRuns = (state.runs || []).filter(function (r) {
      return !r.batchId;
    });
    if (!state.batches.length && !looseRuns.length) {
      container.appendChild(
        el(
          "p",
          "wf-empty-hint",
          "Press Run (or Run all) to execute this workflow. Per-node progress and results appear here.",
        ),
      );
      return;
    }
    var frag = document.createDocumentFragment();
    // Batches first (the active one expanded), then loose single runs. Keyed by id
    // so a newer run can't steal the expansion from an older in-flight one.
    state.batches.forEach(function (batch) {
      frag.appendChild(buildBatchCard(batch, batch.id === state.activeBatchId));
    });
    looseRuns.forEach(function (run) {
      frag.appendChild(buildRunCard(run, run.id === state.activeRunId));
    });
    container.appendChild(frag);
  }

  // ---- Running UI -----------------------------------------------------------

  var _running = false;

  // Re-gate the Run button from the three inputs that can change independently:
  // an in-flight run, the load gate, and validation errors (P5). Called by both
  // setRunningUI and the validation satellite (after every recompute).
  function syncRunButton() {
    var v = state.validation;
    var hasErrors = !!(v && v.errors && v.errors.length);
    var blocked = _running || !state.ready || hasErrors;
    var runBtn = qs("#wfRunBtn");
    if (runBtn) {
      runBtn.disabled = blocked;
      runBtn.title = hasErrors
        ? "Fix the errors in the Issues panel to run"
        : "Run this workflow (set a Video Source to “All participants” to fan out)";
    }
    // "Run to here" needs exactly one selected node (its target).
    var runToBtn = qs("#wfRunToBtn");
    if (runToBtn) {
      var one = state.selection && state.selection.length === 1;
      runToBtn.disabled = blocked || !one;
    }
  }

  function setRunningUI(running) {
    _running = running;
    var stopBtn = qs("#wfStopBtn");
    if (stopBtn) stopBtn.classList.toggle("hidden", !running);
    syncRunButton();
  }

  // ---- Discover externally-started runs (P6 watch-dir triggers) -------------
  // A run can appear without this client starting it — the directory watcher
  // auto-launches one when a new video lands. The run panel otherwise only
  // refreshes on blueprint-open, so such runs would never surface live. A low-
  // frequency poll picks them up; refreshRuns() then reattaches + streams the
  // live one. Gated to idle so it never tears down a stream we're already on.

  // True if the run list for the active blueprint differs from what we hold
  // (a new run id, or a status change) — only then is a full refresh worth it.
  function runsChanged(latest) {
    var cur = state.runs || [];
    if (latest.length !== cur.length) return true;
    var byId = {};
    cur.forEach(function (r) {
      byId[r.id] = r.status;
    });
    for (var i = 0; i < latest.length; i++) {
      if (byId[latest[i].id] !== latest[i].status) return true;
    }
    return false;
  }

  function discoverTick() {
    if (document.hidden || !state.activeBlueprintId) return;
    // An active stream already keeps us current; don't disrupt it.
    if (activeRunInFlight() || activeBatchInFlight()) return;
    if (_stream || _poller || _batchStream || _batchPoller) return;
    apiGet("api/runs?blueprintId=" + encodeURIComponent(state.activeBlueprintId))
      .then(function (res) {
        var latest = (res && res.runs) || [];
        if (runsChanged(latest)) refreshRuns(); // reattaches to any live run
      })
      .catch(function () {});
  }

  function startDiscover() {
    if (_discoverPoller) return;
    _discoverPoller = createPoller(discoverTick, 5000);
    _discoverPoller.start();
  }

  function stopDiscover() {
    if (_discoverPoller) {
      _discoverPoller.stop();
      _discoverPoller = null;
    }
  }

  // Pause/resume the live streams when the tab is hidden (the poller already
  // self-pauses; the EventSource is reopened on return if work is in flight).
  function onVisibility() {
    if (document.hidden) {
      stopStream();
      stopBatchStream();
      return;
    }
    if (activeBatchInFlight() && !_batchStream && !_batchPoller) {
      subscribeBatch(state.activeBatchId);
    }
    if (activeRunInFlight() && !_stream && !_poller) {
      subscribeRun(state.activeRunId);
    }
  }

  function initRuns() {
    document.addEventListener("visibilitychange", onVisibility);
    setRunningUI(false);
    startDiscover();
  }

  // ---- Satellite interface ----
  WF.initRuns = initRuns;
  WF.startRun = startRun; // also fans out to a batch when a source is "All"
  WF.stopRun = stopRun;
  WF.refreshRuns = refreshRuns;
  WF.renderRuns = renderRuns;
  WF.syncRunButton = syncRunButton; // re-gated by the validation satellite
})();
