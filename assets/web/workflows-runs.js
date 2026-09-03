/* Workflows — run/results panel (satellite of workflows.js).
 *
 * Owns the Run lifecycle on the client: POST a run, subscribe to its per-run SSE
 * stream (with a polling fallback), render the run-history list + per-node status
 * rows in #wfRuns, and tint the canvas cards by status with a progress bar.
 * Mirrors the screenspace-tasks SSE+poller pattern.
 */

(function () {
  "use strict";

  var WF = window.ClipgenWorkflows;
  var state = WF.state;

  // Transport handles stay module-local; run and batch pairs stream concurrently during batches.
  var _stream = null; // EventSource for the active/focused run
  var _poller = null; // createPoller fallback when run SSE drops
  var _batchStream = null; // EventSource for the active batch
  var _batchPoller = null; // createPoller fallback when batch SSE drops
  var _discoverPoller = null; // low-freq poll surfacing runs THIS client didn't start
  var _reconnecting = false; // an SSE stream dropped; polling is covering the gap

  var TERMINAL = { completed: 1, degraded: 1, failed: 1, cancelled: 1 };

  function isTerminal(status) {
    return !!TERMINAL[status];
  }

  // "All participants" or a subset of ≥2 ids launches a batch; else one run.
  function blueprintWantsBatch() {
    var nodes = state.nodes || [];
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      if (n.type !== "video_source" || !n.params) continue;
      var p = n.params.participant;
      if (p === WF.ALL_PARTICIPANTS) return true;
      if (Array.isArray(p) && p.length >= 2) return true;
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
    _reconnecting = false; // fresh subscription — clear any stale gap state
    _stream = createSSEStream("api/runs/" + encodeURIComponent(runId) + "/stream", {
      onUnsupported: function () { startPolling(runId); },
      onMessage: function (data) {
        if (data && data.run) handleRunData(data.run);
      },
      onError: function () {
        // SSE dropped: show "Reconnecting…", poll instead; the next poll clears the flag.
        _reconnecting = true;
        _stream = null;
        startPolling(runId);
        renderRuns();
      },
    });
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
      { runImmediately: true, label: "workflows.run" },
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
    _reconnecting = false; // fresh subscription — clear any stale gap state
    _batchStream = createSSEStream(
      "api/batches/" + encodeURIComponent(batchId) + "/stream",
      {
        onUnsupported: function () { startBatchPolling(batchId); },
        onMessage: function (data) {
          if (data && data.batch) handleBatchData(data.batch);
        },
        onError: function () {
          _reconnecting = true;
          _batchStream = null;
          startBatchPolling(batchId);
          renderRuns();
        },
      },
    );
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
      { runImmediately: true, label: "workflows.batch" },
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

  function startRun(targetNodeId, resumeFromRunId, sampleWindowSeconds) {
    if (!state.ready || !state.activeBlueprintId) return;
    if (activeRunInFlight() || activeBatchInFlight()) {
      showToast("A run is already in flight");
      return;
    }
    // Guards the programmatic path; the button is already disabled. Warnings never block.
    if (state.validation && state.validation.errors.length) {
      showToast("Fix the errors in the Issues panel to run");
      return;
    }
    // Partial runs and resumes stay single; a resumed child keeps its participant server-side.
    if (!targetNodeId && !resumeFromRunId && blueprintWantsBatch()) {
      startBatch();
      return;
    }
    setRunningUI(true);
    // Flush pending canvas edits so the server runs the latest blueprint.
    Promise.resolve(WF.flushSave ? WF.flushSave() : null)
      .then(function () {
        var body = { blueprintId: state.activeBlueprintId };
        if (targetNodeId) body.targetNodeId = targetNodeId;
        if (resumeFromRunId) body.resumeFromRunId = resumeFromRunId;
        if (sampleWindowSeconds) body.sampleWindowSeconds = sampleWindowSeconds;
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
    // Stop cancels the in-flight batch (current child + remaining), else the single run.
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

  // Explicit subset from the first Video Source, or null for "all" (field omitted).
  function batchParticipants() {
    var nodes = state.nodes || [];
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      if (n.type !== "video_source" || !n.params) continue;
      var p = n.params.participant;
      if (p === WF.ALL_PARTICIPANTS) return null;
      if (Array.isArray(p) && p.length >= 2) return p.slice();
    }
    return null;
  }

  // One sequential run per selected participant, grouped under one batch card.
  function startBatch() {
    if (!state.ready || !state.activeBlueprintId) return;
    if (activeRunInFlight() || activeBatchInFlight()) {
      showToast("A run is already in flight");
      return;
    }
    setRunningUI(true);
    Promise.resolve(WF.flushSave ? WF.flushSave() : null)
      .then(function () {
        var body = { blueprintId: state.activeBlueprintId };
        var subset = batchParticipants();
        if (subset) body.participants = subset;
        return apiPost("api/batches", body);
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
    // The discover poller is blueprint-scoped, so refresh the "All" list here too.
    if (state.runScope === "all") fetchAllRuns();
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

        // A live batch owns Run/Stop; reattach first and focus its running child.
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
        consumePendingFocus();
      })
      .catch(function () {});
  }

  // History row click-through: drill into pendingFocusRunId once this blueprint's runs load.
  function consumePendingFocus() {
    var pf = state.pendingFocusRunId;
    if (!pf) return;
    state.pendingFocusRunId = null;
    if (findRun(pf)) focusChild(pf);
  }

  function firstNonTerminal(list) {
    for (var i = 0; i < (list || []).length; i++) {
      if (!isTerminal(list[i].status)) return list[i];
    }
    return null;
  }

  // ---- Data handling --------------------------------------------------------

  // Dirty-check gate: SSE/poll ticks re-render only on change. renderRuns() refreshes it.
  var _lastRunsFp = "";

  function runsFingerprint() {
    return JSON.stringify({
      activeRunId: state.activeRunId,
      activeBatchId: state.activeBatchId,
      runFilter: state.runFilter,
      runScope: state.runScope,
      allRuns: (state.allRuns || []).map(function (r) {
        return r.id + ":" + r.status;
      }),
      reconnecting: _reconnecting,
      runs: (state.runs || []).map(function (r) {
        var ns = r.nodeStates || {};
        var nodes = Object.keys(ns).map(function (id) {
          return id + ":" + ns[id].status + ":" + (ns[id].progress || 0);
        });
        return [
          r.id,
          r.status,
          r.startedAt,
          r.completedAt,
          r.triggered,
          r.blueprintId,
          nodes,
        ];
      }),
      batches: (state.batches || []).map(function (b) {
        var children = (b.children || []).map(function (c) {
          return c.status + ":" + c.runId;
        });
        return [b.id, b.status, b.counts, children];
      }),
    });
  }

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
    _reconnecting = false; // data flowed (live SSE or a successful poll)
    upsertRun(run);
    if (run.id === state.activeRunId && isTerminal(run.status)) {
      stopTransport();
      // The batch summary owns Run/Stop; a finished child must not reset them.
      if (!activeBatchInFlight()) setRunningUI(false);
    }
    if (runsFingerprint() !== _lastRunsFp) renderRuns();
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
    _reconnecting = false; // data flowed (live SSE or a successful poll)
    upsertBatch(batch);
    if (batch.id === state.activeBatchId && isTerminal(batch.status)) {
      stopBatchTransport();
      setRunningUI(false);
    }
    if (runsFingerprint() !== _lastRunsFp) renderRuns();
  }

  // Stream one child's per-node detail and tint the canvas; Run/Stop stay batch-owned.
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
    "run-degraded",
    "run-failed",
    "run-skipped",
  ];

  // Tint cards by node status; runs from other blueprints clear instead of painting.
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
      // Orthogonal accent on a completed-with-note node (degraded outcome).
      card.classList.toggle("run-note", !!ns.note);
      setNodeProgress(card, ns.status === "running" ? ns.progress || 0 : 0);
    }
  }

  function clearNodeRun(card) {
    card.classList.remove.apply(card.classList, NODE_RUN_CLASSES);
    card.classList.remove("run-note");
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

  // Rename, else catalog label, else the id; renames disambiguate duplicate node types.
  function nodeLabel(nodeId) {
    var nodes = state.nodes || [];
    for (var i = 0; i < nodes.length; i++) {
      if (nodes[i].id === nodeId) {
        if (nodes[i].name) return nodes[i].name;
        var type = state.catalogById[nodes[i].type];
        return (type && type.label) || nodes[i].type;
      }
    }
    return nodeId;
  }

  // Status glyph for detail rows; CSS keyed on data-status sets icon and colour.
  function statusIcon(status) {
    var icon = el("span", "wf-run-node-icon");
    icon.setAttribute("data-status", status || "queued");
    return icon;
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
    // Terminal chips: counts, plus path chips linking into the output dir via studio media.
    var results = run.results || {};
    var chips = el("div", "wf-run-results");
    var any = false;
    Object.keys(results).forEach(function (nodeId) {
      var ports = results[nodeId] || {};
      Object.keys(ports).forEach(function (port) {
        var val = ports[port];
        if (val == null) return;
        if (typeof val === "object" && val.path) {
          var name = basename(val.path);
          var link = el("a", "wf-run-chip wf-run-chip-link", port + ": " + name);
          link.href = "../studio/media/" + encodeURIComponent(name);
          link.target = "_blank";
          link.rel = "noopener";
          chips.appendChild(link);
          any = true;
          return;
        }
        var text = null;
        if (typeof val === "object") {
          if (typeof val.count === "number") text = port + ": " + val.count;
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

  // Per-node rows + result chips; `hasResult` rows lazily fetch the stored sidecar.
  function buildNodeDetail(run) {
    var wrap = document.createDocumentFragment();
    var rows = el("div", "wf-run-nodes");
    var states = run.nodeStates || {};
    Object.keys(states).forEach(function (nodeId) {
      var ns = states[nodeId];
      var row = el("div", "wf-run-node wf-run-node-" + ns.status);
      row.appendChild(statusIcon(ns.status));
      row.appendChild(el("span", "wf-run-node-label", nodeLabel(nodeId)));
      var detail =
        ns.status === "running"
          ? Math.round((ns.progress || 0) * 100) + "%"
          : ns.status;
      row.appendChild(el("span", "wf-run-node-detail", detail));
      if (ns.error) row.title = ns.error;
      else if (ns.note) row.title = ns.note;
      rows.appendChild(row);
      // A note means the node completed but produced nothing useful; show why.
      if (ns.note) rows.appendChild(el("div", "wf-run-node-note", ns.note));
      if (ns.hasResult) {
        row.classList.add("wf-run-node-expandable");
        var panel = el("div", "wf-result-panel hidden");
        row.addEventListener("click", function (e) {
          e.stopPropagation();
          toggleNodeResult(run, nodeId, panel);
        });
        rows.appendChild(panel);
        if (_expandedResults[run.id + ":" + nodeId]) {
          // Re-open across re-renders (the run._nodeResults cache makes this
          // instant; a not-yet-cached payload just re-fetches once).
          delete _expandedResults[run.id + ":" + nodeId];
          toggleNodeResult(run, nodeId, panel);
        }
      }
    });
    wrap.appendChild(rows);
    var chips = buildResultChips(run);
    if (chips) wrap.appendChild(chips);
    return wrap;
  }

  // ---- Lazy per-node result ------------------------------------------------

  // Expanded result panels, keyed "runId:nodeId"; kept off the DOM since ticks re-render.
  var _expandedResults = {};

  // Toggle a result panel; the payload is fetched once and cached on the run.
  function toggleNodeResult(run, nodeId, panel) {
    var key = run.id + ":" + nodeId;
    if (!panel.classList.contains("hidden")) {
      panel.classList.add("hidden");
      delete _expandedResults[key];
      return;
    }
    _expandedResults[key] = true;
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
    panel.appendChild(el("div", "wf-result-line cg-shimmer", "Loading…"));
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

  // Render a {port: value} result map, branching on each value's shape.
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

  function resultItem(it) {
    return el("div", "wf-result-item", it == null ? "—" : String(it));
  }

  function renderList(body, items) {
    var LIMIT = 8;
    items.slice(0, LIMIT).forEach(function (it) {
      body.appendChild(resultItem(it));
    });
    if (items.length > LIMIT) {
      // "+N more" reveals the rest in place; stopPropagation avoids toggling the row.
      var more = el("button", "wf-result-more", "+" + (items.length - LIMIT) + " more");
      more.type = "button";
      more.addEventListener("click", function (e) {
        e.stopPropagation();
        var frag = document.createDocumentFragment();
        items.slice(LIMIT).forEach(function (it) {
          frag.appendChild(resultItem(it));
        });
        body.insertBefore(frag, more);
        more.parentNode.removeChild(more);
      });
      body.appendChild(more);
    }
  }

  function fmtClock(sec) {
    var s = Math.max(0, Math.round(Number(sec) || 0));
    var m = Math.floor(s / 60);
    var r = s % 60;
    return m + ":" + (r < 10 ? "0" : "") + r;
  }

  // Wall-clock start time of a run/batch (ISO → local HH:MM); "" if unparseable.
  function fmtStartTime(iso) {
    if (!iso) return "";
    var d = new Date(iso);
    if (isNaN(d.getTime())) return "";
    return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  }

  // Elapsed wall-clock between two ISO stamps as a compact "12s" / "3m 4s".
  function fmtDuration(startIso, endIso) {
    if (!startIso || !endIso) return "";
    var ms = new Date(endIso).getTime() - new Date(startIso).getTime();
    if (!(ms >= 0)) return "";
    var s = Math.round(ms / 1000);
    if (s < 60) return s + "s";
    return Math.floor(s / 60) + "m " + (s % 60) + "s";
  }

  // Relative time ("5m ago") for history rows; clock times don't scan well there.
  function fmtRelTime(iso) {
    if (!iso) return "";
    var then = new Date(iso).getTime();
    if (isNaN(then)) return "";
    var s = Math.round((Date.now() - then) / 1000);
    if (s < 60) return "just now";
    if (s < 3600) return Math.floor(s / 60) + "m ago";
    if (s < 86400) return Math.floor(s / 3600) + "h ago";
    return Math.floor(s / 86400) + "d ago";
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
    // Bolt chip marks a trigger-launched run; margin-right:auto hugs the status label.
    if (run.triggered) {
      var trig = el("span", "wf-run-triggered", "triggered");
      if (run.triggerType) trig.title = "Auto-run trigger: " + run.triggerType;
      head.appendChild(trig);
    }
    if (run.sampleWindow) {
      var sw = el("span", "wf-run-triggered", "test " + run.sampleWindow + "s");
      sw.title =
        "Sample-window test: detectors scanned only the first " +
        run.sampleWindow +
        " seconds";
      head.appendChild(sw);
    }
    // Live stream dropped for the active run — polling is covering the gap.
    if (_reconnecting && run.id === state.activeRunId && !isTerminal(run.status)) {
      head.appendChild(el("span", "wf-run-reconnect cg-shimmer", "Reconnecting…"));
    }
    var counts = statusCounts(run);
    var total = Object.keys(run.nodeStates || {}).length;
    var done = (counts.completed || 0) + (counts.skipped || 0);
    // Meta: progress · start · duration in one span so space-between stays stable.
    var metaParts = [done + "/" + total + " nodes"];
    var started = fmtStartTime(run.startedAt);
    if (started) metaParts.push(started);
    var dur = fmtDuration(run.startedAt, run.completedAt);
    if (dur) metaParts.push(dur);
    head.appendChild(el("span", "wf-run-meta", metaParts.join(" · ")));
    // Re-run relaunches the same graph; disabled while anything is in flight.
    if (isTerminal(run.status) && run.blueprintId === state.activeBlueprintId) {
      // Resume reuses completed sidecar results and re-executes only failed nodes plus downstream.
      if (run.status === "failed" || run.status === "cancelled") {
        var resume = el("button", "wf-run-rerun wf-run-resume", "Resume");
        resume.type = "button";
        resume.title =
          "Re-run only the failed part, reusing this run's completed results";
        resume.disabled = activeRunInFlight() || activeBatchInFlight();
        resume.addEventListener("click", function (e) {
          e.stopPropagation();
          startRun(null, run.id);
        });
        head.appendChild(resume);
      }
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
    if (
      _reconnecting &&
      batch.id === state.activeBatchId &&
      !isTerminal(batch.status)
    ) {
      head.appendChild(el("span", "wf-run-reconnect cg-shimmer", "Reconnecting…"));
    }
    var counts = batchCounts(batch);
    var total = (batch.children || []).length;
    var done = counts.completed || 0;
    var parts = [done + "/" + total + " done"];
    if (counts.failed) parts.push(counts.failed + " failed");
    if (counts.cancelled) parts.push(counts.cancelled + " cancelled");
    var bstart = fmtStartTime(batch.createdAt);
    if (bstart) parts.push(bstart);
    var pCount = (batch.participants || []).length || total;
    var pLabel = pCount === 1 ? "1 participant" : pCount + " participants";
    head.appendChild(el("span", "wf-run-meta", pLabel + " · " + parts.join(" · ")));
    card.appendChild(head);

    if (expanded) {
      var rows = el("div", "wf-batch-children");
      (batch.children || []).forEach(function (child) {
        var focused = child.runId === state.activeRunId;
        var row = el(
          "div",
          "wf-batch-child wf-run-node-" + child.status + (focused ? " focused" : ""),
        );
        row.appendChild(statusIcon(child.status));
        row.appendChild(
          el("span", "wf-batch-child-label", child.participant || child.runId),
        );
        row.appendChild(el("span", "wf-run-node-detail", child.status));
        row.addEventListener("click", function () {
          focusChild(child.runId);
        });
        rows.appendChild(row);
        // The focused child expands inline once its full snapshot has streamed in.
        var run = focused ? findRun(child.runId) : null;
        if (run && run.nodeStates) rows.appendChild(buildNodeDetail(run));
      });
      card.appendChild(rows);
    }
    return card;
  }

  // "running" spans every non-terminal status; "failed" also folds in cancelled.
  function runMatchesFilter(status) {
    var f = state.runFilter || "all";
    if (f === "all") return true;
    if (f === "running") return !isTerminal(status);
    // "degraded" ran to the end, so it files under completed, not failed.
    if (f === "completed")
      return status === "completed" || status === "degraded";
    if (f === "failed") return status === "failed" || status === "cancelled";
    return true;
  }

  // Cross-blueprint history row; click opens the blueprint and drills in. Deleted blueprints render inert.
  function buildHistoryRow(run) {
    var bp = null;
    var bps = state.blueprints || [];
    for (var i = 0; i < bps.length; i++) {
      if (bps[i].id === run.blueprintId) {
        bp = bps[i];
        break;
      }
    }
    var row = el("div", "wf-history-row");
    row.appendChild(
      el("span", "wf-run-status wf-run-status-" + run.status, run.status),
    );
    var body = el("div", "wf-history-body");
    body.appendChild(
      el("span", "wf-history-bp", bp ? bp.name || "Untitled" : run.blueprintId),
    );
    var metaParts = [];
    if (run.participant) metaParts.push(run.participant);
    if (run.triggered) metaParts.push("triggered");
    var rel = fmtRelTime(run.startedAt);
    if (rel) metaParts.push(rel);
    if (metaParts.length) {
      body.appendChild(el("span", "wf-history-meta", metaParts.join(" · ")));
    }
    row.appendChild(body);
    if (bp) {
      row.classList.add("clickable");
      row.title = "Open this blueprint and its run";
      row.addEventListener("click", function () {
        state.pendingFocusRunId = run.id;
        setRunScope("blueprint");
        if (WF.openBlueprint) WF.openBlueprint(bp);
      });
    } else {
      row.title = "This run's blueprint no longer exists";
    }
    return row;
  }

  function renderRuns() {
    var container = qs("#wfRuns");
    if (!container) return;
    // Refresh the dirty-check baseline so all callers share one source of truth.
    _lastRunsFp = runsFingerprint();
    var prevScrollTop = container.scrollTop;
    container.innerHTML = "";
    // "All" scope renders state.allRuns flat, never state.runs, so canvas tinting stays untouched.
    if (state.runScope === "all") {
      var all = (state.allRuns || []).filter(function (r) {
        return runMatchesFilter(r.status);
      });
      if (!all.length) {
        container.appendChild(
          el("p", "wf-empty-hint", "No runs in the history yet."),
        );
      } else {
        var histFrag = document.createDocumentFragment();
        all.forEach(function (run) {
          histFrag.appendChild(buildHistoryRow(run));
        });
        container.appendChild(histFrag);
        container.scrollTop = prevScrollTop;
      }
      return;
    }
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
    var batches = state.batches.filter(function (b) {
      return runMatchesFilter(b.status);
    });
    var runs = looseRuns.filter(function (r) {
      return runMatchesFilter(r.status);
    });
    if (!batches.length && !runs.length) {
      container.appendChild(
        el("p", "wf-empty-hint", "No " + (state.runFilter || "") + " runs."),
      );
      return;
    }
    var frag = document.createDocumentFragment();
    // Batches first, then loose runs; keyed by id so expansion follows the right card.
    batches.forEach(function (batch) {
      frag.appendChild(buildBatchCard(batch, batch.id === state.activeBatchId));
    });
    runs.forEach(function (run) {
      frag.appendChild(buildRunCard(run, run.id === state.activeRunId));
    });
    container.appendChild(frag);
    container.scrollTop = prevScrollTop;
    applyLastRunBadges();
  }

  // ---- Last-run badges on canvas cards ---------------------------------------

  // "last: 14 events" badge from the newest terminal run; DOM-applied, no card re-render.
  function lastRunBadgeText(runs, nodeId) {
    for (var i = 0; i < runs.length; i++) {
      var ports = (runs[i].results || {})[nodeId];
      if (!ports) continue;
      var parts = [];
      Object.keys(ports).forEach(function (port) {
        var val = ports[port];
        if (val && typeof val === "object" && typeof val.count === "number") {
          parts.push(val.count + " " + port);
        } else if (val && typeof val === "object" && val.path) {
          parts.push(basename(val.path));
        }
      });
      if (parts.length) return "last: " + parts.join(" · ");
    }
    return null;
  }

  function applyLastRunBadges() {
    var world = qs("#wfWorld");
    if (!world) return;
    // Only terminal runs of the active blueprint, so partial summaries never show.
    var runs = (state.runs || []).filter(function (r) {
      return isTerminal(r.status) && r.blueprintId === state.activeBlueprintId;
    });
    var cards = world.querySelectorAll(".wf-node");
    for (var i = 0; i < cards.length; i++) {
      var card = cards[i];
      if (card.getAttribute("data-node-type") === "note") continue;
      var text = lastRunBadgeText(runs, card.getAttribute("data-node-id"));
      var old = card.querySelector(".wf-node-lastrun");
      if (old && old.textContent === text) continue;
      if (old) old.remove();
      if (text) card.appendChild(el("div", "wf-node-lastrun", text));
    }
  }

  // ---- Cross-blueprint history scope ----------------------------------------

  // Unfiltered history; the server caps it newest-first, so no pagination.
  function fetchAllRuns() {
    apiGet("api/runs")
      .then(function (res) {
        state.allRuns = (res && res.runs) || [];
        if (runsFingerprint() !== _lastRunsFp) renderRuns();
      })
      .catch(function () {});
  }

  function setRunScope(scope) {
    state.runScope = scope === "all" ? "all" : "blueprint";
    var host = qs("#wfRunScope");
    if (host) {
      var chips = host.querySelectorAll(".wf-run-scope-btn");
      for (var i = 0; i < chips.length; i++) {
        chips[i].classList.toggle(
          "active",
          chips[i].getAttribute("data-scope") === state.runScope,
        );
      }
    }
    if (state.runScope === "all") fetchAllRuns();
    renderRuns();
  }

  function initRunScope() {
    var host = qs("#wfRunScope");
    if (!host) return;
    host.addEventListener("click", function (e) {
      var btn = e.target.closest(".wf-run-scope-btn");
      if (!btn) return;
      setRunScope(btn.getAttribute("data-scope"));
    });
  }

  // Status-filter chips above the run list; no-op without the markup.
  function initRunFilter() {
    var host = qs("#wfRunFilter");
    if (!host) return;
    if (!state.runFilter) state.runFilter = "all";
    host.addEventListener("click", function (e) {
      var btn = e.target.closest(".wf-run-filter-btn");
      if (!btn) return;
      state.runFilter = btn.getAttribute("data-filter") || "all";
      var chips = host.querySelectorAll(".wf-run-filter-btn");
      for (var i = 0; i < chips.length; i++) {
        chips[i].classList.toggle("active", chips[i] === btn);
      }
      renderRuns();
    });
  }

  // ---- Running UI -----------------------------------------------------------

  var _running = false;

  // Re-gate Run from three independent inputs; the validation satellite calls this too.
  function syncRunButton() {
    var v = state.validation;
    var hasErrors = !!(v && v.errors && v.errors.length);
    var blocked = _running || !state.ready || hasErrors;
    var runBtn = qs("#wfRunBtn");
    if (runBtn) {
      runBtn.disabled = blocked;
      // [data-tooltip], not native title, so the singleton tooltip doesn't double up.
      runBtn.setAttribute(
        "data-tooltip",
        hasErrors
          ? "Fix the errors in the Issues panel to run"
          : "Run this workflow (set a Video Source to “All participants” to fan out)",
      );
    }
    // The Run split-button caret opens the more-options menu (same gate as Run).
    var caret = qs("#wfRunMenuBtn");
    if (caret) caret.disabled = blocked;
    // "Run to here" needs exactly one selected node (its target).
    var one = state.selection && state.selection.length === 1;
    var runToItem = qs("#wfRunToItem");
    if (runToItem) {
      runToItem.disabled = blocked || !one;
    }
    // Sample-window test needs a detector: it bounds the detector's unwired timeRange input.
    var sampleItem = qs("#wfRunSampleItem");
    if (sampleItem) {
      var selNode =
        one && WF.findNode ? WF.findNode(state.selection[0]) : null;
      var detector =
        selNode &&
        (selNode.type === "detect" || String(selNode.type).indexOf("ss_") === 0);
      sampleItem.disabled = blocked || !detector;
    }
  }

  function setRunningUI(running) {
    _running = running;
    var stopBtn = qs("#wfStopBtn");
    if (stopBtn) stopBtn.classList.toggle("hidden", !running);
    syncRunButton();
  }

  // ---- Discover externally-started runs (trigger-launched) ------------------
  // Trigger-launched runs would never surface live without it.

  // True when the active blueprint's run list gained an id or changed a status.
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
    _discoverPoller = createPoller(discoverTick, 5000, { label: "workflows.discover" });
    _discoverPoller.start();
  }

  // Hidden tab closes streams; on return, reopen them if work is in flight.
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
    initRunFilter();
    initRunScope();
    setRunningUI(false);
    startDiscover();
  }

  // ---- Satellite interface ----
  WF.initRuns = initRuns;
  WF.startRun = startRun; // also fans out to a batch when a source is "All"
  WF.stopRun = stopRun;
  WF.refreshRuns = refreshRuns;
  WF.renderRuns = renderRuns;
  WF.applyLastRunBadges = applyLastRunBadges; // re-applied after renderAllNodes
  WF.syncRunButton = syncRunButton; // re-gated by the validation satellite
})();
