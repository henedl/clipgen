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
  };

  // ---- Catalog + palette ----------------------------------------------------

  // Note: loadCatalog / loadBlueprints deliberately let rejections propagate;
  // loadWorkspace turns a failure into the persistent error gate (a swallowed
  // toast would leave a canvas that looks editable but never saves).
  function loadCatalog() {
    return apiGet("api/catalog").then(function (res) {
      state.catalog = (res && res.catalog) || [];
      state.context = (res && res.context) || { sheet: false, videoDir: false };
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
      item.title = "Requires " + ((node.requires || []).join(", ") || "context");
    }
    return item;
  }

  function renderPalette() {
    var palette = qs("#wfPalette");
    if (!palette || !state.catalog) return;
    palette.innerHTML = "";
    // Group by category, preserving catalog order (per the perf rule, build a
    // DocumentFragment and append once).
    var order = [];
    var byCat = {};
    state.catalog.forEach(function (node) {
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
      byCat[cat].forEach(function (node) {
        frag.appendChild(buildPaletteItem(node));
      });
    });
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
    syncToolbar();
    if (WF.renderAllNodes) WF.renderAllNodes();
    if (WF.applyViewport) WF.applyViewport();
    // Refresh the run panel for the newly-active blueprint (reattaches to an
    // in-flight run if one survived a reload).
    if (WF.refreshRuns) WF.refreshRuns();
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

  // ---- Debounced autosave ---------------------------------------------------

  var _saveTimer = null;

  function cancelSave() {
    if (_saveTimer) {
      clearTimeout(_saveTimer);
      _saveTimer = null;
    }
  }

  function scheduleSave() {
    cancelSave();
    _saveTimer = setTimeout(flushSave, 600);
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
    }).catch(function () {
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
    setToolbarDisabled(!state.ready);
  }

  function setToolbarDisabled(disabled) {
    [
      "#wfBlueprintSelect",
      "#wfBlueprintName",
      "#wfNewBlueprint",
      "#wfDeleteBlueprint",
      "#wfCleanUp",
      "#wfRunBtn",
    ].forEach(function (sel) {
      var node = qs(sel);
      if (node) node.disabled = disabled;
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
    var newBtn = qs("#wfNewBlueprint");
    if (newBtn) newBtn.addEventListener("click", createBlueprint);
    var delBtn = qs("#wfDeleteBlueprint");
    if (delBtn) {
      delBtn.addEventListener("click", function () {
        deleteBlueprint(state.activeBlueprintId);
      });
    }
    var cleanBtn = qs("#wfCleanUp");
    if (cleanBtn) {
      cleanBtn.addEventListener("click", function () {
        if (WF.autoArrange) WF.autoArrange();
      });
    }
    var runBtn = qs("#wfRunBtn");
    if (runBtn) {
      runBtn.addEventListener("click", function () {
        if (WF.startRun) WF.startRun();
      });
    }
    var stopBtn = qs("#wfStopBtn");
    if (stopBtn) {
      stopBtn.addEventListener("click", function () {
        if (WF.stopRun) WF.stopRun();
      });
    }

    var retryBtn = qs("#wfOverlayRetry");
    if (retryBtn) retryBtn.addEventListener("click", loadWorkspace);

    // Bind canvas + wire interactions (handlers no-op until state.ready), then
    // load the workspace, which flips the gate once a blueprint is active.
    if (WF.initCanvas) WF.initCanvas();
    if (WF.initWires) WF.initWires();
    if (WF.initRuns) WF.initRuns();
    loadWorkspace();
  }

  // ---- Satellite interface (window.ClipgenWorkflows) ----
  // The hub publishes `state` + shared helpers onto this namespace; the
  // satellites (workflows-nodes, workflows-canvas) attach their own functions
  // (renderNode, renderAllNodes, initCanvas, applyViewport) back onto it.
  var WF = (window.ClipgenWorkflows = window.ClipgenWorkflows || {});
  WF.state = state;
  WF.boot = boot;
  WF.scheduleSave = scheduleSave;
  WF.flushSave = flushSave; // runs satellite awaits this before POSTing a run
  WF.renderPalette = renderPalette;
  WF.openBlueprint = openBlueprint;
  // Published for the nodes satellite (palette grey-out logic shared, not duped).
  WF.nodeContextMet = nodeContextMet;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
