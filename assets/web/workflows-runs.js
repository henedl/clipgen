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
  var _reconnecting = false; // an SSE stream dropped; polling is covering the gap

  var TERMINAL = { completed: 1, failed: 1, cancelled: 1 };

  function isTerminal(status) {
    return !!TERMINAL[status];
  }

  // A blueprint fans out when any Video Source is set to "All participants" or to
  // a subset of ≥2 participants — the single Run button then launches a batch
  // instead of one run. A single id (string) or a 1-element array runs once.
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
        // SSE dropped — flag the gap (surfaces a "Reconnecting…" pill) and fall
        // back to polling so progress still flows; the next poll clears the flag.
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

  function startRun(targetNodeId, resumeFromRunId) {
    if (!state.ready || !state.activeBlueprintId) return;
    if (activeRunInFlight() || activeBatchInFlight()) return; // one at a time
    // Errors gate the run (the button is already disabled; this guards the
    // programmatic path). Warnings never block.
    if (state.validation && state.validation.errors.length) return;
    // A Video Source set to "All participants" makes Run fan out over the study —
    // but a partial "run to here" or a resume is always a single run (a resumed
    // batch child keeps its participant binding server-side).
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

  // The explicit participant subset to fan out over, or null for "all". Resolved
  // from the first Video Source whose value is an array (subset) or the ALL
  // sentinel — multiple Video Sources share one batch list (the server rebinds
  // them all per child run; a per-source subset is out of scope). An array of ≥2
  // ids is sent as-is; ALL (or no explicit selection) omits the field so the
  // server's "all participants" branch runs.
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

  // Fan the active blueprint out across the selected participants (P3). One run
  // per participant, sequential, grouped under one batch card. Reached from
  // startRun when a Video Source is set to "All participants" or a subset.
  function startBatch() {
    if (!state.ready || !state.activeBlueprintId) return;
    if (activeRunInFlight() || activeBatchInFlight()) return; // one at a time
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
    // Keep the cross-blueprint list current too while it's the visible scope
    // (the discover poller stays blueprint-scoped; "All" refreshes on toggle
    // and on blueprint switches like this one).
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
        consumePendingFocus();
      })
      .catch(function () {});
  }

  // The cross-blueprint history's click-through handshake: the row sets
  // pendingFocusRunId and opens the target blueprint; once refreshRuns has that
  // blueprint's runs loaded, drill into the requested one (batch children are
  // in state.runs too, so focusChild covers both). Cleared silently when the
  // run has since been evicted from history.
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

  // Dirty-check gate for the run panel (mirrors screenspace-tasks' fingerprint).
  // handleRunData/handleBatchData fire on every SSE push / poll tick; without a
  // gate renderRuns() wipes + rebuilds #wfRuns each time, churning the panel and
  // (before the scroll-preserve below) yanking it to the top. renderRuns() itself
  // refreshes _lastRunsFp at the end, so user-driven renders (filter/focus/lifecycle)
  // keep it current too.
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
      // During a batch the batch summary owns the Run/Stop buttons — a finished
      // child must not flip them back to idle while siblings are still running.
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

  // A leading status glyph for a run-detail / batch-child row. The Heroicon and
  // colour are set by CSS keyed on data-status (so the row reads at a glance,
  // not just by its left-border tint).
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
      // A non-fatal note (Ollama down, nothing wired, an adapter that couldn't
      // coerce): the node completed but produced nothing useful — surface why.
      if (ns.note) rows.appendChild(el("div", "wf-run-node-note", ns.note));
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

  function resultItem(it) {
    return el("div", "wf-result-item", it == null ? "—" : String(it));
  }

  function renderList(body, items) {
    var LIMIT = 8;
    items.slice(0, LIMIT).forEach(function (it) {
      body.appendChild(resultItem(it));
    });
    if (items.length > LIMIT) {
      // Clickable "+N more" reveals the rest in place (the payload is already
      // here — no further fetch). stopPropagation so it doesn't toggle the row.
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

  // Compact relative time ("just now" / "5m ago" / "2h ago" / "3d ago") for the
  // cross-blueprint history rows, where absolute clock times don't scan well.
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
    // Auto-launched by the watch-dir trigger (P6) — a bolt chip distinguishes it
    // from a manual run (margin-right:auto keeps it hugging the status label).
    if (run.triggered) {
      var trig = el("span", "wf-run-triggered", "triggered");
      if (run.triggerType) trig.title = "Auto-run trigger: " + run.triggerType;
      head.appendChild(trig);
    }
    // Live stream dropped for the active run — polling is covering the gap.
    if (_reconnecting && run.id === state.activeRunId && !isTerminal(run.status)) {
      head.appendChild(el("span", "wf-run-reconnect", "Reconnecting…"));
    }
    var counts = statusCounts(run);
    var total = Object.keys(run.nodeStates || {}).length;
    var done = (counts.completed || 0) + (counts.skipped || 0);
    // Meta line: node progress · start time · duration (the last two only once
    // the run has a startedAt / has finished). Folded into one span so the head's
    // space-between layout stays stable regardless of how many parts there are.
    var metaParts = [done + "/" + total + " nodes"];
    var started = fmtStartTime(run.startedAt);
    if (started) metaParts.push(started);
    var dur = fmtDuration(run.startedAt, run.completedAt);
    if (dur) metaParts.push(dur);
    head.appendChild(el("span", "wf-run-meta", metaParts.join(" · ")));
    // Re-run a terminal run of the active blueprint — relaunches the same graph
    // (no partial/memoized re-run; the engine just re-executes). Disabled while
    // a run/batch is in flight, mirroring the toolbar Run gate.
    if (isTerminal(run.status) && run.blueprintId === state.activeBlueprintId) {
      // Resume a failed/cancelled run: the server reloads this run's completed
      // node results from its sidecars and executes only what failed (plus
      // everything downstream). Falls back to a full run when nothing is
      // reusable (expired sidecars).
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
      head.appendChild(el("span", "wf-run-reconnect", "Reconnecting…"));
    }
    var counts = batchCounts(batch);
    var total = (batch.children || []).length;
    var done = counts.completed || 0;
    var parts = [done + "/" + total + " done"];
    if (counts.failed) parts.push(counts.failed + " failed");
    if (counts.cancelled) parts.push(counts.cancelled + " cancelled");
    var bstart = fmtStartTime(batch.createdAt);
    if (bstart) parts.push(bstart);
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
        row.appendChild(statusIcon(child.status));
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

  // Client-side run-history filter (state.runFilter). "running" spans every
  // non-terminal status; "failed" also folds in cancelled (both are red ends).
  function runMatchesFilter(status) {
    var f = state.runFilter || "all";
    if (f === "all") return true;
    if (f === "running") return !isTerminal(status);
    if (f === "completed") return status === "completed";
    if (f === "failed") return status === "failed" || status === "cancelled";
    return true;
  }

  // One compact row of the cross-blueprint ("All") history list. Clicking a
  // row opens its blueprint and drills into the run (pendingFocusRunId
  // handshake); a run whose blueprint was deleted renders inert.
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
    // Every render refreshes the dirty-check baseline so the SSE/poll handlers
    // (and user-driven callers) share one source of truth.
    _lastRunsFp = runsFingerprint();
    var prevScrollTop = container.scrollTop;
    container.innerHTML = "";
    // "All" scope: a flat cross-blueprint list (batch children included as
    // plain rows — they carry their participant). Rendered from state.allRuns,
    // never state.runs, so canvas tinting / reattachment are untouched.
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
    // Batches first (the active one expanded), then loose single runs. Keyed by id
    // so a newer run can't steal the expansion from an older in-flight one.
    batches.forEach(function (batch) {
      frag.appendChild(buildBatchCard(batch, batch.id === state.activeBatchId));
    });
    runs.forEach(function (run) {
      frag.appendChild(buildRunCard(run, run.id === state.activeRunId));
    });
    container.appendChild(frag);
    container.scrollTop = prevScrollTop;
  }

  // ---- Cross-blueprint history scope ----------------------------------------

  // Fetch the unfiltered run history (the server already lists every blueprint's
  // runs newest-first, capped at its history limit — no pagination needed).
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

  // Wire the status-filter chips above the run list (set state.runFilter, toggle
  // the active chip, re-render). No-op if the markup isn't present.
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
    // The Run split-button caret opens the more-options menu (same gate as Run).
    var caret = qs("#wfRunMenuBtn");
    if (caret) caret.disabled = blocked;
    // "Run to here" needs exactly one selected node (its target).
    var runToItem = qs("#wfRunToItem");
    if (runToItem) {
      var one = state.selection && state.selection.length === 1;
      runToItem.disabled = blocked || !one;
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
  WF.syncRunButton = syncRunButton; // re-gated by the validation satellite
})();
