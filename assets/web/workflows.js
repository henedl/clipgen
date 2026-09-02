/* Workflows hub — the node-canvas frontend.
 *
 * Establishes the window.ClipgenWorkflows (WF) namespace the satellite files
 * share state and functions through, mirroring screenspace.js and transcripts.js.
 *
 * The hub owns WF.state, boot, the node catalog fetch + palette, the blueprint
 * switcher, and the debounced autosave. Canvas interaction and card rendering
 * live in the satellites, which attach their functions back onto WF.
 */

(function () {
  "use strict";

  // Shared through WF.state so satellites never hit cross-file ReferenceErrors.
  var state = {
    catalog: null, // ordered node-type array from /api/catalog
    catalogById: {}, // id -> node type (built once for card lookups)
    adapters: new Set(), // "src>dst" pairs the runner coerces (widens canConnect)
    adapterDescriptions: {}, // "src>dst" -> what the coercion does (wire tooltip)
    context: { sheet: false, videoDir: false }, // launch context for grey-out
    blueprints: [], // all saved blueprints (full objects)
    nodes: [], // placed cards of the active blueprint: {id, type, params, position}
    edges: [], // wires: {from, fromPort, to, toPort}
    viewport: { x: 0, y: 0, zoom: 1 }, // pan/zoom of the active blueprint
    selection: [], // selected node ids (transient — not persisted)
    selectedEdge: null, // selected wire id (transient — not persisted)
    activeBlueprintId: null,
    ready: false, // false until a blueprint is active; gates canvas edits
    // ---- Run state (owned by the workflows-runs satellite) ----
    runs: [], // recent run snapshots (newest first)
    activeRunId: null, // the run currently streamed/polled, or null
    nodeRunStatus: {}, // node id -> {status, progress} for canvas tinting
    runFilter: "all", // run-history status filter (all|running|completed|failed)
    runScope: "blueprint", // history scope: this blueprint | all (owned by -runs)
    allRuns: [], // cross-blueprint history for the "All" scope (owned by -runs)
    pendingFocusRunId: null, // history click-through handshake (owned by -runs)
    // ---- Batch state (whole-study fan-out; owned by workflows-runs) ----
    batches: [], // recent batch summaries (newest first)
    activeBatchId: null, // the batch currently streamed/polled, or null
    // ---- Stash state (owned by the workflows-stashes satellite) ----
    stashes: [], // built-in recipes + saved sub-graph stashes (built-ins first)
    // ---- Validation (owned by the workflows-validate satellite) ----
    validation: { errors: [], warnings: [] }, // recomputed on every edit
  };

  // ---- Catalog + palette ----------------------------------------------------

  // Rejections propagate on purpose: loadWorkspace turns them into the error gate.
  function loadCatalog() {
    return apiGet("api/catalog").then(function (res) {
      if (res && res.config) clipgenApplyConfig(res.config);
      state.catalog = (res && res.catalog) || [];
      state.context = (res && res.context) || { sheet: false, videoDir: false };
      // canConnect consults this Set so the UI accepts the wires the runner coerces.
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

  // Universal __gate__ control input: a false Gate skips the branch. Carries no data (workflows_runner.py _gather_inputs).
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
    // data-tooltip, not title: native title doesn't render on draggable rows.
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

  // Collection nodes are sub-grouped by operation prefix instead of one flat list.
  var _COLLECTION_OPS = [
    ["filter", "Filter"],
    ["merge", "Merge"],
    ["limit", "Limit"],
    ["dedup", "Dedup"],
  ];

  // Operation prefix of a collection node id (e.g. "filter_events" → "filter").
  function paletteOp(node) {
    return String(node.id || "").split("_")[0];
  }

  // <details> sub-groups per operation; an active search expands them so matches show.
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

  function paletteNodeMatches(node, query) {
    if (!query) return true;
    var hay = (
      (node.id || "") +
      " " +
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
    palette.removeAttribute("aria-busy"); // boot skeletons are about to go
    palette.innerHTML = "";
    var searchInput = qs("#wfPaletteSearch");
    var query = (searchInput ? searchInput.value : "").trim().toLowerCase();
    // Catalog order; a category label appears only when one of its nodes survives the filter.
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
    // The selected <option> already shows the name; nothing else to sync.
    if (sel && state.activeBlueprintId) sel.value = state.activeBlueprintId;
  }

  // Flushes the outgoing blueprint's pending save so no edit is lost or mis-attributed.
  function openBlueprint(bp) {
    if (!bp) return;
    // An in-flight wire gesture would otherwise persist an edge into the old blueprint's node.
    if (WF.cancelConnect) WF.cancelConnect();
    flushSave();
    state.activeBlueprintId = bp.id;
    state.nodes = bp.nodes || (bp.nodes = []);
    state.edges = bp.edges || (bp.edges = []);
    state.viewport = bp.viewport || (bp.viewport = { x: 0, y: 0, zoom: 1 });
    state.selection = [];
    state.selectedEdge = null;
    if (WF.clearRunPreview) WF.clearRunPreview(); // stale preview classes
    resetHistory(); // history doesn't span blueprints
    syncToolbar();
    // A corrupt blueprint must not throw into loadWorkspace: that disables the whole toolbar.
    try {
      syncTriggerButton();
      if (WF.renderAllNodes) WF.renderAllNodes();
      if (WF.applyViewport) WF.applyViewport();
      // Reattaches to an in-flight run that survived a reload.
      if (WF.refreshRuns) WF.refreshRuns();
      // Validate the freshly-loaded graph (gates Run, populates the Issues panel).
      if (WF.refreshValidation) WF.refreshValidation();
    } catch (err) {
      if (window.console && console.error) {
        console.error("Failed to render blueprint " + bp.id, err);
      }
      showToast("This blueprint couldn't be loaded — delete it or pick another");
    }
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
    // Clear the id first so an in-flight autosave PUT becomes a no-op (flushSave bails).
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

  // Same shape as the autosave PUT, so import round-trips bar id/createdAt.
  function exportBlueprint() {
    var bp = findBlueprint(state.activeBlueprintId);
    if (!bp) return;
    var data = {
      name: bp.name || "Untitled",
      nodes: state.nodes || [],
      edges: state.edges || [],
      viewport: state.viewport || { x: 0, y: 0, zoom: 1 },
    };
    clipgenSaveFile(
      (bp.name || "blueprint").replace(/[^a-zA-Z0-9_-]/g, "_") + ".json",
      JSON.stringify(data, null, 2),
      "application/json"
    );
  }

  // Unknown node types are accepted as-is; they fail at run time like removed types.
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

  // Rename the active blueprint via the shared prompt modal (mirrors Stash).
  function openRenameDialog() {
    var bp = findBlueprint(state.activeBlueprintId);
    if (!bp) return;
    openPromptDialog({
      title: "Rename blueprint",
      initial: bp.name || "",
      confirmLabel: "Save",
      onConfirm: function (v) {
        renameActive((v || "").trim() || "Untitled");
      },
    });
  }

  // Confirm-then-delete the active blueprint (shared by the toolbar button and
  // the Mod+Shift+Backspace hotkey).
  function requestDeleteBlueprint() {
    var bp = findBlueprint(state.activeBlueprintId);
    if (!bp) return;
    openConfirmDialog({
      title: "Delete blueprint “" + (bp.name || "Untitled") + "”?",
      body: "This can't be undone.",
      danger: true,
      onConfirm: function () {
        deleteBlueprint(bp.id);
      },
    });
  }

  // ---- Auto-run triggers ----------------------------------------------------
  // Single-active per trigger type; server disarms the rest, client mirrors it.

  var _triggerMenuCtl = null; // bindMenuToggle handle for the type picker

  function triggerTypes() {
    return (state.context && state.context.triggerTypes) || [];
  }

  function triggerTypeLabel(id) {
    var types = triggerTypes();
    for (var i = 0; i < types.length; i++) {
      if (types[i].id === id) return types[i].label;
    }
    return id;
  }

  function blueprintArmed(bp) {
    return !!(bp && bp.trigger && bp.trigger.enabled);
  }

  // The active blueprint's armed trigger type, or "" when disarmed.
  function activeTriggerType() {
    var bp = findBlueprint(state.activeBlueprintId);
    return blueprintArmed(bp) ? String(bp.trigger.type || "") : "";
  }

  function armedBlueprints() {
    return (state.blueprints || []).filter(blueprintArmed);
  }

  function hasVideoSource() {
    return (state.nodes || []).some(function (n) {
      return n.type === "video_source";
    });
  }

  function setTrigger(type, enabled) {
    var id = state.activeBlueprintId;
    if (!id) return;
    apiPut("api/blueprints/" + encodeURIComponent(id) + "/trigger", {
      enabled: enabled,
      type: type,
    })
      .then(function (res) {
        if (!res || !res.ok) {
          showToast((res && res.error) || "Couldn't update auto-run");
          return;
        }
        // Mirror the server's disarm of same-type triggers elsewhere.
        if (enabled) {
          state.blueprints.forEach(function (b) {
            if (b.id !== id && b.trigger && b.trigger.type === type) {
              b.trigger.enabled = false;
            }
          });
        }
        var bp = findBlueprint(id);
        if (bp) bp.trigger = res.blueprint.trigger;
        syncTriggerButton();
      })
      .catch(function () {
        // 4xx (e.g. arming a broken graph) is normally prevented client-side; generic is enough.
        showToast("Couldn't update auto-run");
      });
  }

  // Called from syncTriggerButton so the active checkmark is current when the menu opens.
  function rebuildTriggerMenu() {
    var menu = qs("#wfTriggerMenu");
    if (!menu) return;
    menu.innerHTML = "";
    var active = activeTriggerType();
    triggerTypes().forEach(function (t) {
      var item = el("button", "wf-run-menu-item", t.label);
      item.type = "button";
      item.setAttribute("role", "menuitem");
      if (t.id === active) item.classList.add("wf-trigger-item-active");
      item.title =
        t.id === active
          ? "Turn auto-run off"
          : "Auto-run this blueprint when: " + t.label.toLowerCase();
      item.addEventListener("click", function () {
        if (_triggerMenuCtl) _triggerMenuCtl.close();
        // Clicking the armed type disarms it; any other type arms/moves it.
        setTrigger(t.id, t.id !== active);
      });
      menu.appendChild(item);
    });
  }

  function syncTriggerButton() {
    var btn = qs("#wfTriggerBtn");
    if (!btn) return;
    var activeType = activeTriggerType();
    var armed = !!activeType;
    var v = state.validation;
    var hasErrors = !!(v && v.errors && v.errors.length);
    var armable = !hasErrors && hasVideoSource();
    // Disarming is always allowed; arming needs a ready, valid graph with a source.
    btn.disabled = !state.ready || (!armed && !armable);
    btn.classList.toggle("wf-trigger-armed", armed);
    btn.setAttribute("aria-pressed", armed ? "true" : "false");
    var label = btn.querySelector(".wf-trigger-label");
    if (label) {
      label.textContent = armed
        ? "Auto-run: " + triggerTypeLabel(activeType)
        : "Auto-run";
    }
    // data-tooltip, not title, or it doubles up with the singleton tooltip.
    if (armed) {
      btn.setAttribute(
        "data-tooltip",
        "Auto-runs when: " +
          triggerTypeLabel(activeType).toLowerCase() +
          ". Open to change or turn off.",
      );
    } else if (!hasVideoSource()) {
      btn.setAttribute("data-tooltip", "Add a Video Source node to enable auto-run");
    } else if (hasErrors) {
      btn.setAttribute("data-tooltip", "Fix the errors in the Issues panel to enable auto-run");
    } else {
      btn.setAttribute(
        "data-tooltip",
        "Auto-run this blueprint when a video lands, a transcript completes, or a scan completes",
      );
    }
    rebuildTriggerMenu();
    // Global cue listing armed blueprints, even when none is the active canvas.
    var hint = qs("#wfArmedHint");
    if (hint) {
      var armedList = armedBlueprints();
      if (armedList.length) {
        hint.textContent =
          "⚡ Auto-run: " +
          armedList
            .map(function (b) {
              return b.name || "Untitled";
            })
            .join(", ");
        hint.title = armedList
          .map(function (b) {
            return (
              "“" +
              (b.name || "Untitled") +
              "” runs when: " +
              triggerTypeLabel(String(b.trigger.type || "")).toLowerCase()
            );
          })
          .join(" · ");
        hint.classList.remove("hidden");
      } else {
        hint.classList.add("hidden");
      }
    }
  }

  // ---- Dialogs (prompt / confirm) -------------------------------------------
  // Native-dialog replacements on openBlockingModal; one lazy singleton overlay.

  var _dialogOverlay = null;

  function openDialog(opts) {
    if (!_dialogOverlay) {
      _dialogOverlay = el("div", "wf-dialog-overlay cg-modal-overlay hidden");
      _dialogOverlay.appendChild(el("div", "wf-dialog cg-modal-card"));
      document.body.appendChild(_dialogOverlay);
    }
    var overlay = _dialogOverlay;
    var box = overlay.querySelector(".wf-dialog");
    box.innerHTML = "";
    box.classList.toggle("wf-dialog-danger", !!opts.danger);
    box.appendChild(el("div", "wf-dialog-title", opts.title || ""));
    if (opts.body) box.appendChild(el("p", "wf-dialog-body", opts.body));
    // Form-wrapped so Enter submits the prompt input natively.
    var form = document.createElement("form");
    var input = null;
    if (opts.prompt) {
      input = document.createElement("input");
      input.type = "text";
      input.className = "wf-dialog-input";
      input.autocomplete = "off";
      input.value = opts.initial || "";
      form.appendChild(input);
    }
    var actions = el("div", "wf-dialog-actions");
    var cancelBtn = el("button", "btn btn-small", "Cancel");
    cancelBtn.type = "button";
    cancelBtn.addEventListener("click", close);
    var confirmBtn = el(
      "button",
      "btn btn-small btn-primary",
      opts.confirmLabel || "OK"
    );
    confirmBtn.type = "submit";
    actions.appendChild(cancelBtn);
    actions.appendChild(confirmBtn);
    form.appendChild(actions);
    box.appendChild(form);
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      var value = input ? input.value : null;
      close();
      opts.onConfirm(value);
    });
    function close() {
      overlay.classList.add("hidden");
      closeBlockingModal(overlay);
    }
    overlay.classList.remove("hidden");
    openBlockingModal(overlay, {
      trapFocus: true, // the input (when present) is first, so it gets focus
      restoreFocus: true,
      onEscape: close,
      onBackdropClick: close,
    });
    if (input) input.select();
  }

  // openPromptDialog({title, initial, confirmLabel, onConfirm(value)})
  function openPromptDialog(opts) {
    openDialog({
      title: opts.title,
      prompt: true,
      initial: opts.initial,
      confirmLabel: opts.confirmLabel || "Save",
      onConfirm: opts.onConfirm,
    });
  }

  // openConfirmDialog({title, body, danger, confirmLabel, onConfirm()})
  function openConfirmDialog(opts) {
    openDialog({
      title: opts.title,
      body: opts.body,
      danger: opts.danger,
      confirmLabel: opts.confirmLabel || "Delete",
      onConfirm: function () {
        opts.onConfirm();
      },
    });
  }

  // ---- Undo / redo ----------------------------------------------------------
  // {nodes, edges} snapshots; save debounce coalesces bursts into one step.

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

  // Called from every history mutation so the buttons never lie.
  function syncUndoButtons() {
    var u = qs("#wfUndo");
    var r = qs("#wfRedo");
    if (u) u.disabled = !state.ready || !_undoStack.length;
    if (r) r.disabled = !state.ready || !_redoStack.length;
  }

  // First mutation of a burst pushes the baseline; the rest are absorbed (snapPending).
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

  function setSaveStatus(mode) {
    var status = qs("#wfSaveStatus");
    if (!status) return;
    status.classList.toggle("wf-save-error", mode === "error");
    status.classList.toggle("cg-shimmer", mode === "saving");
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
    // Validation is not debounced, so the Issues panel never lags an edit.
    if (WF.refreshValidation) WF.refreshValidation();
  }

  // scheduleSave minus history/validation: a pan must never push an undo step. Settles pending bursts.
  function scheduleViewportSave() {
    cancelSave();
    setSaveStatus("saving");
    _saveTimer = setTimeout(function () {
      if (_snapPending) {
        _baseline = cloneGraph();
        _snapPending = false;
      }
      flushSave();
    }, 600);
  }

  // Syncs working state into the blueprint, then PUTs; returns the promise for run starts.
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

  // Canvas stays gated until a blueprint is active; a load failure shows a retryable error.
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
    if (mode === "error") {
      // Clear the skeleton rows: a failed load never reaches the renders that replace them.
      ["#wfPalette", "#wfStashList"].forEach(function (sel) {
        var host = qs(sel);
        if (!host) return;
        host.innerHTML = "";
        host.removeAttribute("aria-busy");
      });
    }
    var overlay = qs("#wfCanvasOverlay");
    if (overlay) overlay.classList.toggle("hidden", state.ready);
    var msg = qs("#wfOverlayMsg");
    if (msg) {
      msg.classList.toggle("cg-shimmer", mode !== "error");
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
    // The blanket enable above ignores the stash button's selection gate; re-apply it.
    if (state.ready && WF.syncStashButton) WF.syncStashButton();
    // Same for the trigger's gate (valid graph + Video Source).
    if (state.ready) syncTriggerButton();
    // Same for undo/redo, which gate on stack contents.
    syncUndoButtons();
  }

  function setToolbarDisabled(disabled) {
    [
      "#wfBlueprintSelect",
      "#wfRenameBlueprint",
      "#wfNewBlueprint",
      "#wfDeleteBlueprint",
      "#wfUndo",
      "#wfRedo",
      "#wfCleanUp",
      "#wfAddNote",
      "#wfSnapBtn",
      "#wfRunBtn",
      "#wfRunMenuBtn",
      "#wfSaveStash",
      "#wfTriggerBtn",
    ].forEach(function (sel) {
      var node = qs(sel);
      if (node) node.disabled = disabled;
    });
  }

  // onOpen fires after unhide (menu is measurable); wrapping {open, close} misses the click path.
  function bindMenuToggle(btn, menu, opts) {
    opts = opts || {};
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
      if (opts.onOpen) opts.onOpen();
      document.addEventListener("mousedown", onDocDown, true);
      document.addEventListener("keydown", onKey, true);
    }
    function close() {
      menu.classList.add("hidden");
      btn.setAttribute("aria-expanded", "false");
      document.removeEventListener("mousedown", onDocDown, true);
      document.removeEventListener("keydown", onKey, true);
      if (opts.onClose) opts.onClose();
    }
    btn.addEventListener("click", function () {
      if (menu.classList.contains("hidden")) open();
      else close();
    });
    return { open: open, close: close };
  }

  // Caret menu: "Run to here" runs the selected node plus its ancestors.
  function initRunMenu() {
    var caret = qs("#wfRunMenuBtn");
    var menu = qs("#wfRunMenu");
    var runTo = qs("#wfRunToItem");
    var runSample = qs("#wfRunSampleItem");
    if (!caret || !menu) return;
    var ctl = bindMenuToggle(caret, menu);
    if (runTo) {
      runTo.addEventListener("click", function () {
        ctl.close();
        var sel = state.selection || [];
        if (WF.startRun && sel.length === 1) WF.startRun(sel[0]);
      });
    }
    if (runSample) {
      runSample.addEventListener("click", function () {
        ctl.close();
        var sel = state.selection || [];
        if (WF.startRun && sel.length === 1) {
          WF.startRun(sel[0], null, 30);
        }
      });
    }
  }

  // Import/export live in TopNav Quick Actions; onBeforeOpen refreshes the export gate.
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
    // Click the matching data-filter chip so initRunFilter owns the state change.
    function runFilterCommand(filter, title, icon) {
      return {
        id: "workflows:runs-" + filter,
        title: title,
        icon: icon,
        keywords: "run history filter status " + filter,
        section: "Workflows",
        visible: function () { return !!qs("#wfRunFilter"); },
        run: function () {
          var btn = qs('#wfRunFilter .wf-run-filter-btn[data-filter="' + filter + '"]');
          if (btn) btn.click();
        },
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
          "zoom center canvas", "wfMinimapFit"),
        buttonCommand("workflows:undo", "Undo", "arrow-uturn-left",
          "revert history", "wfUndo"),
        buttonCommand("workflows:redo", "Redo", "arrow-uturn-right",
          "repeat history", "wfRedo"),
        buttonCommand("workflows:cleanup", "Clean up canvas", "squares-2x2",
          "arrange tidy layout auto", "wfCleanUp"),
        buttonCommand("workflows:add-note", "Add sticky note", "pencil-square",
          "annotate comment canvas", "wfAddNote"),
        buttonCommand("workflows:toggle-trigger", "Toggle auto-run on new video", "bolt",
          "watch dir trigger arm", "wfTriggerBtn"),
        runFilterCommand("all", "Show all runs", "bars-3"),
        runFilterCommand("failed", "Show failed runs", "funnel"),
      ];
    });
  }

  // ---- Boot -----------------------------------------------------------------

  function boot() {
    // TopNav renders these buttons before the hub loads; wire them here. No page-specific settings.
    if (typeof initThemeToggle === "function") {
      initThemeToggle();
    }
    if (window.wireSettingsButton) window.wireSettingsButton({});

    // Blueprint switcher toolbar.
    var sel = qs("#wfBlueprintSelect");
    if (sel) {
      sel.addEventListener("change", function () {
        openBlueprint(findBlueprint(sel.value));
      });
    }
    var renameBtn = qs("#wfRenameBlueprint");
    if (renameBtn) renameBtn.addEventListener("click", openRenameDialog);
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
    if (delBtn) delBtn.addEventListener("click", requestDeleteBlueprint);
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
    var snapBtn = qs("#wfSnapBtn");
    if (snapBtn) {
      snapBtn.addEventListener("click", function () {
        if (WF.toggleSnap) WF.toggleSnap();
      });
    }
    var noteBtn = qs("#wfAddNote");
    if (noteBtn) {
      noteBtn.addEventListener("click", function () {
        if (WF.addNote) WF.addNote();
      });
    }

    // Blueprint action hotkeys; canvas-editing ones register in workflows-canvas.js. All gate on readiness.
    function _blueprintReady() {
      return !!state.ready;
    }
    window.ClipgenHotkeys.register([
      { id: "workflows.cleanUp", when: _blueprintReady, handler: function () { if (WF.autoArrange) WF.autoArrange(); } },
      { id: "workflows.toggleSnap", when: _blueprintReady, handler: function () { if (WF.toggleSnap) WF.toggleSnap(); } },
      {
        id: "workflows.stash",
        when: function () { return _blueprintReady() && !!(state.selection && state.selection.length); },
        handler: function () { if (WF.saveSelectionAsStash) WF.saveSelectionAsStash(); },
      },
      { id: "workflows.newBlueprint", when: _blueprintReady, handler: function () { createBlueprint(); } },
      { id: "workflows.renameBlueprint", when: _blueprintReady, handler: function () { openRenameDialog(); } },
      {
        id: "workflows.focusSelector",
        when: _blueprintReady,
        handler: function () { var s = qs("#wfBlueprintSelect"); if (s) s.focus(); },
      },
      {
        id: "workflows.deleteBlueprint",
        when: function () { return _blueprintReady() && !!state.activeBlueprintId; },
        handler: function () { requestDeleteBlueprint(); },
      },
    ]);
    // Zoom readout doubles as reset-to-100%; writeViewport keeps its text current.
    var zoomLevelBtn = qs("#wfZoomLevel");
    if (zoomLevelBtn) {
      zoomLevelBtn.addEventListener("click", function () {
        if (WF.zoomAtCenter) {
          WF.zoomAtCenter(1 / (state.viewport.zoom || 1));
        }
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
        if (WF.clearRunPreview) WF.clearRunPreview();
        if (WF.startRun) WF.startRun();
      });
    }
    // Hover preview of what would run. On the container: a disabled button swallows mouse events.
    var runSplit = qs(".wf-run-split");
    if (runSplit) {
      runSplit.addEventListener("mouseenter", function () {
        if (WF.showRunPreview) WF.showRunPreview(null);
      });
      runSplit.addEventListener("mouseleave", function () {
        if (WF.clearRunPreview) WF.clearRunPreview();
      });
    }
    var runToPreview = qs("#wfRunToItem");
    if (runToPreview) {
      runToPreview.addEventListener("mouseenter", function () {
        var sel = state.selection || [];
        if (WF.showRunPreview && sel.length === 1) WF.showRunPreview(sel[0]);
      });
      runToPreview.addEventListener("mouseleave", function () {
        if (WF.clearRunPreview) WF.clearRunPreview();
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
    var triggerMenu = qs("#wfTriggerMenu");
    if (triggerBtn && triggerMenu) {
      _triggerMenuCtl = bindMenuToggle(triggerBtn, triggerMenu);
    }

    var retryBtn = qs("#wfOverlayRetry");
    if (retryBtn) retryBtn.addEventListener("click", loadWorkspace);

    // Handlers no-op until state.ready; loadWorkspace flips the gate.
    if (WF.initCanvas) WF.initCanvas();
    if (WF.initWires) WF.initWires();
    if (WF.initRuns) WF.initRuns();
    if (WF.initStashes) WF.initStashes();
    loadWorkspace();
  }

  // ---- Satellite interface (window.ClipgenWorkflows) ----
  // Hub publishes state + helpers; satellites attach their functions back.
  var WF = (window.ClipgenWorkflows = window.ClipgenWorkflows || {});
  WF.state = state;
  // Video Source participant sentinel meaning "every participant"; Run turns it into a batch.
  WF.ALL_PARTICIPANTS = "__all__";
  WF.scheduleSave = scheduleSave;
  WF.scheduleViewportSave = scheduleViewportSave;
  WF.undo = undo;
  WF.redo = redo;
  WF.flushSave = flushSave; // runs satellite awaits this before POSTing a run
  WF.renderPalette = renderPalette;
  WF.openBlueprint = openBlueprint;
  // Published for the nodes satellite (palette grey-out logic shared, not duped).
  WF.nodeContextMet = nodeContextMet;
  // The validate satellite re-gates the trigger on every edit: errors block arming.
  WF.syncTriggerButton = syncTriggerButton;
  // For the nodes satellite's participant popover.
  WF.bindMenuToggle = bindMenuToggle;
  // In-page prompt/confirm dialogs (workflows-stashes.js consumes them).
  WF.openPromptDialog = openPromptDialog;
  WF.openConfirmDialog = openConfirmDialog;

  // Deferred satellites run after this hub; boot() needs their init fns, so wait for DOMContentLoaded.
  document.addEventListener("DOMContentLoaded", boot);
})();
