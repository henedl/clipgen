/* Workflows hub — the node-canvas frontend (4th top-level surface).
 *
 * Establishes the window.ClipgenWorkflows (WF) namespace that satellite files
 * (workflows-nodes, workflows-canvas) share state and functions through —
 * mirroring screenspace.js + window.ClipgenScreenspace and transcripts.js +
 * window.ClipgenTranscripts.
 *
 * The hub owns: WF.state, boot, the node catalog fetch + palette, the
 * blueprint switcher (create/name/switch/delete), and the debounced autosave.
 * The canvas interaction layer and card rendering live in the satellites, which
 * attach their functions back onto WF. Vanilla JS, ES5-style (no build step).
 */

(function () {
  "use strict";

  // Shared mutable state. Routed through WF.state (not bare `var`s) so the
  // satellites read/write the same object without cross-file ReferenceErrors —
  // the carve gotcha that bit Screenspace and Transcripts.
  var state = {
    catalog: null, // ordered node-type array from /api/catalog
    catalogById: {}, // id -> node type (built once for card lookups)
    adapters: new Set(), // "src>dst" pairs the runner coerces (widens canConnect)
    adapterDescriptions: {}, // "src>dst" -> what the coercion does (wire tooltip)
    context: { sheet: false, videoDir: false }, // launch context for grey-out
    blueprints: [], // all saved blueprints (full objects)
    nodes: [], // placed cards of the active blueprint: {id, type, params, position}
    edges: [], // wires (M2): {from, fromPort, to, toPort}
    viewport: { x: 0, y: 0, zoom: 1 }, // pan/zoom of the active blueprint
    selection: [], // selected node ids (transient — not persisted)
    selectedEdge: null, // selected wire id (transient — not persisted)
    activeBlueprintId: null,
    ready: false, // false until a blueprint is active; gates canvas edits
    // ---- Run state (M4; owned by the workflows-runs satellite) ----
    runs: [], // recent run snapshots (newest first)
    activeRunId: null, // the run currently streamed/polled, or null
    nodeRunStatus: {}, // node id -> {status, progress} for canvas tinting
    runFilter: "all", // run-history status filter (all|running|completed|failed)
    // ---- Batch state (P3: whole-study fan-out; owned by workflows-runs) ----
    batches: [], // recent batch summaries (newest first)
    activeBatchId: null, // the batch currently streamed/polled, or null
    // ---- Stash state (M5; owned by the workflows-stashes satellite) ----
    stashes: [], // built-in recipes + saved sub-graph stashes (built-ins first)
    // ---- Validation (P5; owned by the workflows-validate satellite) ----
    validation: { errors: [], warnings: [] }, // recomputed on every edit
  };

  // ---- Catalog + palette ----------------------------------------------------

  // Note: loadCatalog / loadBlueprints deliberately let rejections propagate;
  // loadWorkspace turns a failure into the persistent error gate (a swallowed
  // toast would leave a canvas that looks editable but never saves).
  function loadCatalog() {
    return apiGet("api/catalog").then(function (res) {
      if (res && res.config) clipgenApplyConfig(res.config);
      state.catalog = (res && res.catalog) || [];
      state.context = (res && res.context) || { sheet: false, videoDir: false };
      // Adapter pairs the runner coerces across; canConnect consults this Set so
      // the UI accepts the same wires the runner runs (events→clipRecords, …).
      state.adapters = new Set();
      state.adapterDescriptions = {};
      ((res && res.adapters) || []).forEach(function (pair) {
        var key = pair[0] + ">" + pair[1];
        state.adapters.add(key);
        if (pair[2]) state.adapterDescriptions[key] = pair[2];
      });
      state.catalogById = {};
      state.catalog.forEach(function (n) {
        injectControlPort(n);
        state.catalogById[n.id] = n;
      });
      renderPalette();
    });
  }

  // Every node gets a universal optional `control` input (`__gate__`): a Gate's
  // `control`-typed `pass` output wires here (exact-match) to gate the node, so a
  // false gate skips the whole downstream branch. It carries no data — the runner
  // excludes control edges from a node's inputs (see workflows.py _gather_inputs).
  function injectControlPort(node) {
    node.inputs = node.inputs || [];
    var has = node.inputs.some(function (p) {
      return p.name === "__gate__";
    });
    if (!has) {
      node.inputs.unshift({
        name: "__gate__",
        type: "control",
        optional: true,
        control: true,
      });
    }
  }

  // True when every `requires` entry is satisfied by the launch context.
  function nodeContextMet(node) {
    var reqs = (node && node.requires) || [];
    for (var i = 0; i < reqs.length; i++) {
      if (!state.context || !state.context[reqs[i]]) return false;
    }
    return true;
  }

  function buildPaletteItem(node) {
    var item = el("div", "wf-palette-item", node.label);
    item.setAttribute("data-domain", node.domain || "");
    // Brief description tooltip for every palette row (mirrors the on-card `?`).
    // Use the [data-tooltip] singleton, not native title — title doesn't render
    // on draggable=true rows, and the singleton is styled/positioned in-viewport.
    if (node.description) item.setAttribute("data-tooltip", node.description);
    if (nodeContextMet(node)) {
      item.draggable = true;
      item.dataset.nodeType = node.id;
      item.addEventListener("dragstart", function (e) {
        if (!state.ready) {
          e.preventDefault(); // don't start a drag the canvas can't accept yet
          return;
        }
        e.dataTransfer.setData("application/x-wf-node-type", node.id);
        e.dataTransfer.setData("text/plain", node.id);
        e.dataTransfer.effectAllowed = "copy";
      });
    } else {
      item.classList.add("disabled");
      // Disabled rows still show the description, plus what context they need.
      var req = "Requires " + ((node.requires || []).join(", ") || "context");
      item.setAttribute(
        "data-tooltip",
        node.description ? node.description + ". " + req : req,
      );
    }
    return item;
  }

  // The ~23 collection-algebra nodes (filter_*/partition_*/merge_*/limit_*/
  // dedup_*) share one "Collection" category; group them by operation so the
  // palette shows 5 collapsible sub-groups instead of a long flat list.
  var _COLLECTION_OPS = [
    ["filter", "Filter"],
    ["partition", "Partition"],
    ["merge", "Merge"],
    ["limit", "Limit"],
    ["dedup", "Dedup"],
  ];

  // Operation prefix of a collection node id (e.g. "filter_events" → "filter").
  function paletteOp(node) {
    return String(node.id || "").split("_")[0];
  }

  // Emit the Collection category as <details> sub-groups by operation. While a
  // search is active, sub-groups auto-expand so matches aren't hidden.
  function appendCollectionGroups(frag, nodes, query) {
    var byOp = {};
    nodes.forEach(function (n) {
      var op = paletteOp(n);
      (byOp[op] || (byOp[op] = [])).push(n);
    });
    _COLLECTION_OPS.forEach(function (pair) {
      var items = byOp[pair[0]];
      if (!items || !items.length) return;
      var details = el("details", "wf-palette-subgroup");
      if (query) details.open = true;
      details.appendChild(el("summary", "wf-palette-subgroup-label", pair[1]));
      items.forEach(function (n) {
        details.appendChild(buildPaletteItem(n));
      });
      frag.appendChild(details);
    });
  }

  // A node matches the palette search when the query is empty or a substring of
  // its label, description, or category (all case-insensitive).
  function paletteNodeMatches(node, query) {
    if (!query) return true;
    var hay = (
      (node.label || "") +
      " " +
      (node.description || "") +
      " " +
      (node.category || "")
    ).toLowerCase();
    return hay.indexOf(query) !== -1;
  }

  function renderPalette() {
    var palette = qs("#wfPalette");
    if (!palette || !state.catalog) return;
    palette.innerHTML = "";
    var searchInput = qs("#wfPaletteSearch");
    var query = (searchInput ? searchInput.value : "").trim().toLowerCase();
    // Group by category, preserving catalog order (per the perf rule, build a
    // DocumentFragment and append once). A group label is only emitted when at
    // least one of its nodes survives the search filter.
    var order = [];
    var byCat = {};
    state.catalog.forEach(function (node) {
      if (node.hidden) return; // kept in the catalog for specs, not palette-facing
      if (!paletteNodeMatches(node, query)) return;
      var cat = node.category || "Other";
      if (!byCat[cat]) {
        byCat[cat] = [];
        order.push(cat);
      }
      byCat[cat].push(node);
    });
    var frag = document.createDocumentFragment();
    order.forEach(function (cat) {
      frag.appendChild(el("div", "wf-palette-group-label", cat));
      if (cat === "Collection") {
        appendCollectionGroups(frag, byCat[cat], query);
      } else {
        byCat[cat].forEach(function (node) {
          frag.appendChild(buildPaletteItem(node));
        });
      }
    });
    if (!order.length && query) {
      frag.appendChild(el("div", "wf-palette-empty", "No matching nodes"));
    }
    palette.appendChild(frag);
  }

  // ---- Blueprint switcher ---------------------------------------------------

  function findBlueprint(id) {
    for (var i = 0; i < state.blueprints.length; i++) {
      if (state.blueprints[i].id === id) return state.blueprints[i];
    }
    return null;
  }

  function populateSelect() {
    var sel = qs("#wfBlueprintSelect");
    if (!sel) return;
    sel.innerHTML = "";
    var frag = document.createDocumentFragment();
    state.blueprints.forEach(function (bp) {
      var opt = el("option");
      opt.value = bp.id;
      opt.textContent = bp.name || "Untitled";
      frag.appendChild(opt);
    });
    sel.appendChild(frag);
    if (state.activeBlueprintId) sel.value = state.activeBlueprintId;
  }

  function syncToolbar() {
    var sel = qs("#wfBlueprintSelect");
    var nameInput = qs("#wfBlueprintName");
    var bp = findBlueprint(state.activeBlueprintId);
    if (sel && state.activeBlueprintId) sel.value = state.activeBlueprintId;
    if (nameInput) nameInput.value = (bp && bp.name) || "";
  }

  // Make `bp` the active canvas. Flushes the outgoing blueprint's pending save
  // first so a debounced edit is never lost or mis-attributed on switch.
  function openBlueprint(bp) {
    if (!bp) return;
    // Drop any in-flight wire gesture before swapping context — otherwise its
    // armed listeners survive and the next click could persist an edge that
    // references a node from the old blueprint.
    if (WF.cancelConnect) WF.cancelConnect();
    flushSave();
    state.activeBlueprintId = bp.id;
    state.nodes = bp.nodes || (bp.nodes = []);
    state.edges = bp.edges || (bp.edges = []);
    state.viewport = bp.viewport || (bp.viewport = { x: 0, y: 0, zoom: 1 });
    state.selection = [];
    state.selectedEdge = null;
    resetHistory(); // history doesn't span blueprints
    syncToolbar();
    syncTriggerButton();
    if (WF.renderAllNodes) WF.renderAllNodes();
    if (WF.applyViewport) WF.applyViewport();
    // Refresh the run panel for the newly-active blueprint (reattaches to an
    // in-flight run if one survived a reload).
    if (WF.refreshRuns) WF.refreshRuns();
    // Validate the freshly-loaded graph (gates Run, populates the Issues panel).
    if (WF.refreshValidation) WF.refreshValidation();
  }

  function loadBlueprints() {
    return apiGet("api/blueprints").then(function (res) {
      var list = (res && res.blueprints) || [];
      if (!list.length) {
        // Fresh launch — auto-create one so the canvas is immediately usable.
        return apiPost("api/blueprints", { name: "Untitled" }).then(function (r) {
          state.blueprints = [r.blueprint];
          populateSelect();
          openBlueprint(r.blueprint);
        });
      }
      state.blueprints = list;
      populateSelect();
      openBlueprint(list[0]);
    });
  }

  function createBlueprint() {
    apiPost("api/blueprints", { name: "Untitled" })
      .then(function (res) {
        if (!res || !res.ok) {
          showToast("Failed to create blueprint");
          return;
        }
        state.blueprints.push(res.blueprint);
        populateSelect();
        openBlueprint(res.blueprint);
      })
      .catch(function () {
        showToast("Failed to create blueprint");
      });
  }

  function deleteBlueprint(id) {
    if (!id) return;
    cancelSave(); // don't resurrect the row we're about to delete
    // Clear the active id up front so any autosave PUT already in flight for it
    // is a no-op (flushSave bails on a falsy id), not after the DELETE resolves.
    state.activeBlueprintId = null;
    apiDelete("api/blueprints/" + encodeURIComponent(id))
      .then(function () {
        state.blueprints = state.blueprints.filter(function (b) {
          return b.id !== id;
        });
        if (state.blueprints.length) {
          populateSelect();
          openBlueprint(state.blueprints[0]);
        } else {
          createBlueprint();
        }
      })
      .catch(function () {
        showToast("Failed to delete blueprint");
      });
  }

  // ---- Import / export ------------------------------------------------------

  // Download the active blueprint as JSON (the same {name,nodes,edges,viewport}
  // shape the autosave PUT sends, so import round-trips losslessly bar the
  // server-assigned id/createdAt).
  function exportBlueprint() {
    var bp = findBlueprint(state.activeBlueprintId);
    if (!bp) return;
    var data = {
      name: bp.name || "Untitled",
      nodes: state.nodes || [],
      edges: state.edges || [],
      viewport: state.viewport || { x: 0, y: 0, zoom: 1 },
    };
    var blob = new Blob([JSON.stringify(data, null, 2)], {
      type: "application/json",
    });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = (bp.name || "blueprint").replace(/[^a-zA-Z0-9_-]/g, "_") + ".json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Parse a JSON file and create a new blueprint from it. Unknown node types are
  // accepted as-is (they fail gracefully at run time, like any removed type).
  function importBlueprint(file) {
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function () {
      var data;
      try {
        data = JSON.parse(String(reader.result || ""));
      } catch (e) {
        showToast("Import failed: not valid JSON");
        return;
      }
      if (!data || !Array.isArray(data.nodes) || !Array.isArray(data.edges)) {
        showToast("Import failed: missing nodes/edges");
        return;
      }
      apiPost("api/blueprints", {
        name: data.name || "Imported",
        nodes: data.nodes,
        edges: data.edges,
        viewport: data.viewport || { x: 0, y: 0, zoom: 1 },
      })
        .then(function (res) {
          if (!res || !res.ok) {
            showToast("Import failed");
            return;
          }
          state.blueprints.push(res.blueprint);
          populateSelect();
          openBlueprint(res.blueprint);
        })
        .catch(function () {
          showToast("Import failed");
        });
    };
    reader.readAsText(file);
  }

  function renameActive(name) {
    var bp = findBlueprint(state.activeBlueprintId);
    if (!bp) return;
    bp.name = name;
    var sel = qs("#wfBlueprintSelect");
    if (sel) {
      var opts = sel.options;
      for (var i = 0; i < opts.length; i++) {
        if (opts[i].value === bp.id) {
          opts[i].textContent = name || "Untitled";
          break;
        }
      }
    }
    scheduleSave();
  }

  // ---- Watch-dir trigger (P6) -----------------------------------------------
  // A single armed blueprint auto-runs when a new participant video lands in the
  // input dir. Arming is single-active: the server disarms every other blueprint,
  // mirrored client-side below so the toolbar hint is accurate without a refetch.

  function blueprintArmed(bp) {
    return !!(
      bp &&
      bp.trigger &&
      bp.trigger.type === "watch_dir" &&
      bp.trigger.enabled
    );
  }

  function activeTriggerArmed() {
    return blueprintArmed(findBlueprint(state.activeBlueprintId));
  }

  function armedBlueprint() {
    for (var i = 0; i < state.blueprints.length; i++) {
      if (blueprintArmed(state.blueprints[i])) return state.blueprints[i];
    }
    return null;
  }

  function hasVideoSource() {
    return (state.nodes || []).some(function (n) {
      return n.type === "video_source";
    });
  }

  function toggleTrigger() {
    var id = state.activeBlueprintId;
    if (!id) return;
    var enabling = !activeTriggerArmed();
    apiPut("api/blueprints/" + encodeURIComponent(id) + "/trigger", {
      enabled: enabling,
    })
      .then(function (res) {
        if (!res || !res.ok) {
          showToast((res && res.error) || "Couldn't update auto-run");
          return;
        }
        // Single-active: enabling here disarmed every other blueprint server-side.
        if (enabling) {
          state.blueprints.forEach(function (b) {
            if (b.id !== id && b.trigger && b.trigger.type === "watch_dir") {
              b.trigger.enabled = false;
            }
          });
        }
        var bp = findBlueprint(id);
        if (bp) bp.trigger = res.blueprint.trigger;
        syncTriggerButton();
      })
      .catch(function () {
        // apiPut throws on a 4xx (e.g. arming a broken graph); the client gate
        // below normally prevents this, so a generic message is enough.
        showToast("Couldn't update auto-run");
      });
  }

  function syncTriggerButton() {
    var btn = qs("#wfTriggerBtn");
    if (!btn) return;
    var armed = activeTriggerArmed();
    var v = state.validation;
    var hasErrors = !!(v && v.errors && v.errors.length);
    var armable = !hasErrors && hasVideoSource();
    // Disarming is always allowed; arming needs a ready, valid graph with a source.
    btn.disabled = !state.ready || (!armed && !armable);
    btn.classList.toggle("wf-trigger-armed", armed);
    btn.setAttribute("aria-pressed", armed ? "true" : "false");
    var label = btn.querySelector(".wf-trigger-label");
    if (label) label.textContent = armed ? "Auto-running" : "Auto-run on new video";
    var other = armedBlueprint();
    if (armed) {
      btn.title =
        "Watching the input folder. New videos auto-run this blueprint. Click to stop.";
    } else if (other) {
      btn.title =
        "Auto-run is armed on “" +
        (other.name || "Untitled") +
        "”. Click to move it here.";
    } else if (!hasVideoSource()) {
      btn.title = "Add a Video Source node to enable auto-run on new videos";
    } else if (hasErrors) {
      btn.title = "Fix the errors in the Issues panel to enable auto-run";
    } else {
      btn.title =
        "Auto-run this blueprint when a new video lands in the input folder";
    }
    // Persistent global cue: which blueprint (if any) is armed, even when it
    // isn't the active canvas (the button only reflects the active blueprint).
    var hint = qs("#wfArmedHint");
    if (hint) {
      var armedBp = armedBlueprint();
      if (armedBp) {
        hint.textContent = "⚡ Auto-run: " + (armedBp.name || "Untitled");
        hint.title =
          "“" + (armedBp.name || "Untitled") + "” runs automatically on new videos";
        hint.classList.remove("hidden");
      } else {
        hint.classList.add("hidden");
      }
    }
  }

  // ---- Undo / redo ----------------------------------------------------------
  //
  // History is a stack of {nodes, edges} snapshots (the autosave shape; viewport
  // and selection are excluded so a pan never lands on the undo stack). Capture
  // hangs off the scheduleSave chokepoint and is coalesced by the same 600 ms
  // debounce, so a burst of param keystrokes collapses to one undo step.

  var _undoStack = [];
  var _redoStack = [];
  var _baseline = null; // last settled graph; what an undo restores TO
  var _snapPending = false; // a burst is mid-flight (already captured once)
  var _UNDO_CAP = 50;

  function cloneGraph() {
    return {
      nodes: JSON.parse(JSON.stringify(state.nodes || [])),
      edges: JSON.parse(JSON.stringify(state.edges || [])),
    };
  }

  function resetHistory() {
    _undoStack = [];
    _redoStack = [];
    _snapPending = false;
    _baseline = cloneGraph();
    syncUndoButtons();
  }

  // Toolbar undo/redo buttons mirror the keyboard stack; disabled when the
  // canvas isn't ready or the respective stack is empty. Called from every
  // history mutation so the buttons never lie about what's available.
  function syncUndoButtons() {
    var u = qs("#wfUndo");
    var r = qs("#wfRedo");
    if (u) u.disabled = !state.ready || !_undoStack.length;
    if (r) r.disabled = !state.ready || !_redoStack.length;
  }

  // On the first mutation of a burst, push the pre-burst baseline; later
  // keystrokes in the same burst are absorbed (snapPending) into one step.
  function captureHistory() {
    if (!_baseline) {
      _baseline = cloneGraph();
      return;
    }
    if (_snapPending) return;
    _undoStack.push(_baseline);
    if (_undoStack.length > _UNDO_CAP) _undoStack.shift();
    _redoStack = []; // a fresh edit invalidates the redo branch
    _snapPending = true;
    syncUndoButtons();
  }

  // Restore a graph snapshot and persist it without re-capturing history.
  function applyGraph(graph) {
    if (WF.cancelConnect) WF.cancelConnect();
    state.nodes = JSON.parse(JSON.stringify(graph.nodes || []));
    state.edges = JSON.parse(JSON.stringify(graph.edges || []));
    state.selection = [];
    state.selectedEdge = null;
    _baseline = cloneGraph();
    _snapPending = false;
    if (WF.renderAllNodes) WF.renderAllNodes();
    if (WF.refreshValidation) WF.refreshValidation();
    cancelSave();
    flushSave();
  }

  function undo() {
    if (!_undoStack.length) return false;
    _redoStack.push(cloneGraph());
    applyGraph(_undoStack.pop());
    syncUndoButtons();
    return true;
  }

  function redo() {
    if (!_redoStack.length) return false;
    _undoStack.push(cloneGraph());
    applyGraph(_redoStack.pop());
    syncUndoButtons();
    return true;
  }

  // ---- Debounced autosave ---------------------------------------------------

  var _saveTimer = null;

  // Reflect autosave state in the toolbar: "saving" while the debounce timer is
  // armed / a PUT is in flight, "saved" once it resolves, "error" on failure.
  function setSaveStatus(mode) {
    var status = qs("#wfSaveStatus");
    if (!status) return;
    status.classList.toggle("wf-save-error", mode === "error");
    status.textContent =
      mode === "saving"
        ? "Saving…"
        : mode === "saved"
          ? "Saved"
          : mode === "error"
            ? "Save failed"
            : "";
  }

  function cancelSave() {
    if (_saveTimer) {
      clearTimeout(_saveTimer);
      _saveTimer = null;
    }
  }

  function scheduleSave() {
    captureHistory();
    cancelSave();
    setSaveStatus("saving");
    _saveTimer = setTimeout(function () {
      // Burst settled → this is the new baseline a future edit captures from.
      _baseline = cloneGraph();
      _snapPending = false;
      flushSave();
    }, 600);
    // Validation is immediate (not debounced) so the Issues panel + Run button
    // never lag an edit. Every graph mutation funnels through here.
    if (WF.refreshValidation) WF.refreshValidation();
  }

  // Persist the active blueprint now (also syncs working state back into the
  // in-memory blueprint so switching without a refetch stays correct). Returns
  // the PUT promise so callers that need the server up to date (e.g. starting a
  // run) can await it; a resolved promise when there's nothing to save.
  function flushSave() {
    cancelSave();
    var id = state.activeBlueprintId;
    if (!id) return Promise.resolve();
    var bp = findBlueprint(id);
    if (bp) {
      bp.nodes = state.nodes;
      bp.edges = state.edges;
      bp.viewport = state.viewport;
    }
    return apiPut("api/blueprints/" + encodeURIComponent(id), {
      name: bp ? bp.name : undefined,
      nodes: state.nodes,
      edges: state.edges,
      viewport: state.viewport,
    })
      .then(function () {
        setSaveStatus("saved");
      })
      .catch(function () {
        setSaveStatus("error");
        showToast("Failed to save blueprint");
      });
  }

  // ---- Load gate ------------------------------------------------------------

  // Load the catalog + blueprints, then mark the canvas ready. Until a blueprint
  // is active the canvas is gated (overlay + disabled toolbar + interaction
  // handlers that no-op on !state.ready), so edits can't land in a void and a
  // load failure shows a persistent, retryable error instead of a canvas that
  // looks editable but silently never saves.
  function loadWorkspace() {
    setCanvasState("loading");
    return loadCatalog()
      .then(loadBlueprints) // catalog first so cards can resolve labels
      .then(function () {
        return WF.loadStashes ? WF.loadStashes() : null; // built-ins + user stashes
      })
      .then(function () {
        setCanvasState("ready");
      })
      .catch(function () {
        setCanvasState("error");
      });
  }

  function setCanvasState(mode) {
    state.ready = mode === "ready";
    var overlay = qs("#wfCanvasOverlay");
    if (overlay) overlay.classList.toggle("hidden", state.ready);
    var msg = qs("#wfOverlayMsg");
    if (msg) {
      msg.textContent =
        mode === "error"
          ? "Couldn't load workflows. Check the server, then retry."
          : "Loading workflows…";
    }
    var retry = qs("#wfOverlayRetry");
    if (retry) retry.classList.toggle("hidden", mode !== "error");
    // Spinner only spins while loading — on error it's replaced by the retry CTA.
    var spinner = qs("#wfOverlaySpinner");
    if (spinner) spinner.classList.toggle("hidden", mode === "error");
    setToolbarDisabled(!state.ready);
    // setToolbarDisabled(false) re-enables every toolbar control, but "Stash
    // selection" must stay disabled until nodes are selected — re-apply its
    // selection gate after the blanket enable on entering the ready state.
    if (state.ready && WF.syncStashButton) WF.syncStashButton();
    // Blanket re-enable above ignores the trigger's nuanced gate (valid graph +
    // a Video Source); re-apply it after entering the ready state.
    if (state.ready) syncTriggerButton();
    // Undo/redo gate on stack contents, not just readiness — re-derive after the
    // blanket enable so they don't light up on a fresh, history-less canvas.
    syncUndoButtons();
  }

  function setToolbarDisabled(disabled) {
    [
      "#wfBlueprintSelect",
      "#wfBlueprintName",
      "#wfNewBlueprint",
      "#wfDeleteBlueprint",
      "#wfUndo",
      "#wfRedo",
      "#wfCleanUp",
      "#wfFitView",
      "#wfRunBtn",
      "#wfRunMenuBtn",
      "#wfSaveStash",
      "#wfTriggerBtn",
    ].forEach(function (sel) {
      var node = qs(sel);
      if (node) node.disabled = disabled;
    });
  }

  // Generic dropdown toggle: wires a trigger button to a menu element with
  // outside-click + Escape close (shared by the Run split-button and the
  // shortcuts legend). Returns {open, close} for callers that need them.
  function bindMenuToggle(btn, menu) {
    function onDocDown(e) {
      // Ignore mousedowns on the trigger itself so its click handler toggles (a
      // close-then-reopen race otherwise).
      if (btn.contains(e.target) || menu.contains(e.target)) return;
      close();
    }
    function onKey(e) {
      if (e.key === "Escape") close();
    }
    function open() {
      menu.classList.remove("hidden");
      btn.setAttribute("aria-expanded", "true");
      document.addEventListener("mousedown", onDocDown, true);
      document.addEventListener("keydown", onKey, true);
    }
    function close() {
      menu.classList.add("hidden");
      btn.setAttribute("aria-expanded", "false");
      document.removeEventListener("mousedown", onDocDown, true);
      document.removeEventListener("keydown", onKey, true);
    }
    btn.addEventListener("click", function () {
      if (menu.classList.contains("hidden")) open();
      else close();
    });
    return { open: open, close: close };
  }

  // Run split-button: the caret opens a small menu whose "Run to here" item runs
  // the selected node + its ancestors (a partial run). The primary button still
  // runs the whole graph.
  function initRunMenu() {
    var caret = qs("#wfRunMenuBtn");
    var menu = qs("#wfRunMenu");
    var runTo = qs("#wfRunToItem");
    if (!caret || !menu) return;
    var ctl = bindMenuToggle(caret, menu);
    if (runTo) {
      runTo.addEventListener("click", function () {
        ctl.close();
        var sel = state.selection || [];
        if (WF.startRun && sel.length === 1) WF.startRun(sel[0]);
      });
    }
  }

  // TopNav Quick Actions (mirrors Studio / Screenspace / Transcripts): blueprint
  // JSON import/export live here, off the toolbar. Export is gated on an active
  // blueprint; onBeforeOpen refreshes that state each time the menu opens.
  function buildQuickActions() {
    if (!window.ClipgenTopNav) return;
    var importFile = qs("#wfImportFile");
    function rebuild() {
      window.ClipgenTopNav.setQuickActions([
        {
          icon: "arrow-down-tray",
          label: "Export blueprint JSON",
          action: exportBlueprint,
          disabled: !state.activeBlueprintId,
          title: state.activeBlueprintId
            ? "Download the active blueprint as a JSON file"
            : "Open a blueprint first to export it.",
        },
        {
          icon: "arrow-up-tray",
          label: "Import blueprint JSON",
          action: function () {
            if (importFile) importFile.click();
          },
          title: "Create a new blueprint from a JSON file",
        },
      ]);
    }
    rebuild();
    window.ClipgenTopNav.onBeforeOpen(rebuild);
  }

  // Command palette (command-palette.js): additions beyond the auto-ingested
  // quick actions — the toolbar's run/stop/new/fit/undo/redo buttons.
  function initCommandPalette() {
    if (!window.ClipgenCommandPalette) return;
    window.ClipgenCommandPalette.setParticipants(function () {
      return (state.context && state.context.participants) || [];
    });
    function buttonCommand(id, title, icon, keywords, elId) {
      return {
        id: id,
        title: title,
        icon: icon,
        keywords: keywords,
        section: "Workflows",
        enabled: function () {
          var btn = qs("#" + elId);
          return !!btn && !btn.disabled;
        },
        run: function () { qs("#" + elId).click(); },
      };
    }
    window.ClipgenCommandPalette.register("workflows", function () {
      var stopBtn = qs("#wfStopBtn");
      return [
        buttonCommand("workflows:run", "Run blueprint", "play",
          "execute start batch", "wfRunBtn"),
        {
          id: "workflows:stop",
          title: "Stop run",
          icon: "stop",
          keywords: "cancel abort",
          section: "Workflows",
          visible: !!stopBtn && !stopBtn.classList.contains("hidden"),
          run: function () { qs("#wfStopBtn").click(); },
        },
        buttonCommand("workflows:new", "New blueprint", "squares-plus",
          "create canvas", "wfNewBlueprint"),
        buttonCommand("workflows:fit", "Fit to view", "arrows-pointing-out",
          "zoom center canvas", "wfFitView"),
        buttonCommand("workflows:undo", "Undo", "arrow-uturn-left",
          "revert history", "wfUndo"),
        buttonCommand("workflows:redo", "Redo", "arrow-uturn-right",
          "repeat history", "wfRedo"),
      ];
    });
  }

  // ---- Boot -----------------------------------------------------------------

  function boot() {
    // TopNav renders the theme toggle (#themeToggle) and Settings (#settingsBtn)
    // buttons synchronously before this hub loads, so wire them here as the
    // other surfaces do. M1 has no page-specific settings, so Settings just
    // opens the shared modal.
    if (typeof initThemeToggle === "function") {
      initThemeToggle();
    }
    var settingsBtn = qs("#settingsBtn");
    if (settingsBtn && typeof window.openSettingsModal === "function") {
      settingsBtn.addEventListener("click", function () {
        window.openSettingsModal({});
      });
    }

    // Blueprint switcher toolbar.
    var sel = qs("#wfBlueprintSelect");
    if (sel) {
      sel.addEventListener("change", function () {
        openBlueprint(findBlueprint(sel.value));
      });
    }
    var nameInput = qs("#wfBlueprintName");
    if (nameInput) {
      nameInput.addEventListener("input", function () {
        renameActive(nameInput.value);
      });
    }
    var paletteSearch = qs("#wfPaletteSearch");
    if (paletteSearch) {
      paletteSearch.addEventListener("input", renderPalette);
    }
    var newBtn = qs("#wfNewBlueprint");
    if (newBtn) newBtn.addEventListener("click", createBlueprint);
    var importFile = qs("#wfImportFile");
    if (importFile) {
      importFile.addEventListener("change", function () {
        importBlueprint(importFile.files && importFile.files[0]);
        importFile.value = ""; // allow re-importing the same file
      });
    }
    buildQuickActions();
    initCommandPalette();
    var delBtn = qs("#wfDeleteBlueprint");
    if (delBtn) {
      delBtn.addEventListener("click", function () {
        deleteBlueprint(state.activeBlueprintId);
      });
    }
    var undoBtn = qs("#wfUndo");
    if (undoBtn) {
      undoBtn.addEventListener("click", function () {
        undo();
      });
    }
    var redoBtn = qs("#wfRedo");
    if (redoBtn) {
      redoBtn.addEventListener("click", function () {
        redo();
      });
    }
    var cleanBtn = qs("#wfCleanUp");
    if (cleanBtn) {
      cleanBtn.addEventListener("click", function () {
        if (WF.autoArrange) WF.autoArrange();
      });
    }
    var fitBtn = qs("#wfFitView");
    if (fitBtn) {
      fitBtn.addEventListener("click", function () {
        if (WF.fitToView) WF.fitToView();
      });
    }
    // Minimap zoom controls (in/out about the canvas centre + fit-to-content).
    var zoomInBtn = qs("#wfZoomIn");
    if (zoomInBtn) {
      zoomInBtn.addEventListener("click", function () {
        if (WF.zoomAtCenter) WF.zoomAtCenter(1.25);
      });
    }
    var zoomOutBtn = qs("#wfZoomOut");
    if (zoomOutBtn) {
      zoomOutBtn.addEventListener("click", function () {
        if (WF.zoomAtCenter) WF.zoomAtCenter(1 / 1.25);
      });
    }
    var minimapFitBtn = qs("#wfMinimapFit");
    if (minimapFitBtn) {
      minimapFitBtn.addEventListener("click", function () {
        if (WF.fitToView) WF.fitToView();
      });
    }
    var runBtn = qs("#wfRunBtn");
    if (runBtn) {
      runBtn.addEventListener("click", function () {
        if (WF.startRun) WF.startRun();
      });
    }
    initRunMenu();
    var stopBtn = qs("#wfStopBtn");
    if (stopBtn) {
      stopBtn.addEventListener("click", function () {
        if (WF.stopRun) WF.stopRun();
      });
    }
    var triggerBtn = qs("#wfTriggerBtn");
    if (triggerBtn) triggerBtn.addEventListener("click", toggleTrigger);

    var retryBtn = qs("#wfOverlayRetry");
    if (retryBtn) retryBtn.addEventListener("click", loadWorkspace);

    // Bind canvas + wire interactions (handlers no-op until state.ready), then
    // load the workspace, which flips the gate once a blueprint is active.
    if (WF.initCanvas) WF.initCanvas();
    if (WF.initWires) WF.initWires();
    if (WF.initRuns) WF.initRuns();
    if (WF.initStashes) WF.initStashes();
    loadWorkspace();
  }

  // ---- Satellite interface (window.ClipgenWorkflows) ----
  // The hub publishes `state` + shared helpers onto this namespace; the
  // satellites (workflows-nodes, workflows-canvas) attach their own functions
  // (renderNode, renderAllNodes, initCanvas, applyViewport) back onto it.
  var WF = (window.ClipgenWorkflows = window.ClipgenWorkflows || {});
  WF.state = state;
  // Sentinel a Video Source's participant param holds for "fan out over every
  // participant"; the Run button turns such a blueprint into a batch. (Frontend-
  // only — the batch rebinds each video_source to a real participant per run.)
  WF.ALL_PARTICIPANTS = "__all__";
  WF.scheduleSave = scheduleSave;
  WF.undo = undo;
  WF.redo = redo;
  WF.flushSave = flushSave; // runs satellite awaits this before POSTing a run
  WF.renderPalette = renderPalette;
  WF.openBlueprint = openBlueprint;
  // Published for the nodes satellite (palette grey-out logic shared, not duped).
  WF.nodeContextMet = nodeContextMet;
  // Published so the validate satellite re-gates the trigger toggle on every edit
  // (you can't arm a graph with errors), alongside its WF.syncRunButton call.
  WF.syncTriggerButton = syncTriggerButton;
  // Published for the nodes satellite's participant multi-select popover (same
  // outside-click/Escape toggle the run + shortcuts menus use).
  WF.bindMenuToggle = bindMenuToggle;

  // Every workflows script loads with `defer` (see workflows.html), so this hub
  // runs at readyState "interactive" — after DOM parse but BEFORE DOMContentLoaded
  // and BEFORE the workflows-*.js satellites execute. boot() invokes the
  // satellites' WF.initCanvas/initWires/initRuns/initStashes, so it must wait for
  // DOMContentLoaded (dispatched only once every deferred script has run).
  // Booting synchronously here would silently skip those init fns (still
  // undefined) and leave the canvas with no pan/zoom/drag handlers. Mirrors the
  // sibling hubs (studio/screenspace/transcripts).
  document.addEventListener("DOMContentLoaded", boot);
})();
