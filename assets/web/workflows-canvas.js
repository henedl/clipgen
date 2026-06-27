/* Workflows — canvas interaction layer (satellite of workflows.js).
 *
 * Pan/zoom (transformed #wfWorld), palette drag-drop, node move, click /
 * shift-click / marquee selection, and Delete. Mirrors screenspace.js patterns:
 * clientToWorld coord mapping (cf. canvasCoords), RAF-throttled redraws, and
 * zoom-about-cursor (cf. the timeline wheel handler). Reads shared state through
 * WF.state and calls satellite/hub fns late-bound via WF.* (load-order-safe).
 */

(function () {
  "use strict";

  var WF = window.ClipgenWorkflows;
  var state = WF.state;

  var ZOOM_MIN = 0.25;
  var ZOOM_MAX = 2.5;
  var _gridBase = 24; // px at zoom 1; synced from --wf-grid-size in initCanvas
  var _bound = false;
  var _vpRaf = 0;
  var _moveRaf = 0;

  function clamp(v, lo, hi) {
    return v < lo ? lo : v > hi ? hi : v;
  }

  function randomId() {
    return Math.random().toString(36).slice(2, 10);
  }

  function findNode(id) {
    for (var i = 0; i < state.nodes.length; i++) {
      if (state.nodes[i].id === id) return state.nodes[i];
    }
    return null;
  }

  function typesHas(types, t) {
    if (!types) return false;
    for (var i = 0; i < types.length; i++) {
      if (types[i] === t) return true;
    }
    return false;
  }

  // ---- Viewport ----

  // Write the world transform + grid offset, RAF-throttled (cf. screenspace's
  // scheduleOverlayRender).
  function applyViewport() {
    if (_vpRaf) return;
    _vpRaf = requestAnimationFrame(function () {
      _vpRaf = 0;
      var world = qs("#wfWorld");
      var canvas = qs("#wfCanvas");
      var vp = state.viewport;
      // One transform on #wfWorld pans/zooms both the cards and the nested wire
      // layer in lockstep (no per-path recompute on pan/zoom — only on node move).
      var t = "translate(" + vp.x + "px," + vp.y + "px) scale(" + vp.zoom + ")";
      if (world) world.style.transform = t;
      if (canvas) {
        var g = _gridBase * vp.zoom;
        canvas.style.backgroundSize = g + "px " + g + "px";
        canvas.style.backgroundPosition = vp.x + "px " + vp.y + "px";
      }
      // The wire-delete button is screen-positioned, so reposition it when the
      // viewport changes (wires themselves move with the SVG transform).
      if (WF.refreshWireDelete) WF.refreshWireDelete();
    });
  }

  function clientToWorld(cx, cy) {
    var canvas = qs("#wfCanvas");
    var rect = canvas.getBoundingClientRect();
    var vp = state.viewport;
    return {
      x: (cx - rect.left - vp.x) / vp.zoom,
      y: (cy - rect.top - vp.y) / vp.zoom,
    };
  }

  // ---- Palette → canvas (HTML5 drag-drop) ----

  function onDragOver(e) {
    if (!state.ready) return;
    var types = e.dataTransfer && e.dataTransfer.types;
    if (
      typesHas(types, "application/x-wf-node-type") ||
      typesHas(types, "application/x-wf-stash")
    ) {
      e.preventDefault();
      e.dataTransfer.dropEffect = "copy";
    }
  }

  function onDrop(e) {
    if (!state.ready) return;
    // A stash drop instantiates a whole sub-graph (the stashes satellite remaps
    // ids); a node-type drop creates one card. Try the stash MIME first.
    var stashId = e.dataTransfer.getData("application/x-wf-stash");
    if (stashId) {
      e.preventDefault();
      if (WF.instantiateStash) {
        WF.instantiateStash(stashId, clientToWorld(e.clientX, e.clientY));
      }
      return;
    }
    var type = e.dataTransfer.getData("application/x-wf-node-type");
    if (!type) return;
    e.preventDefault();
    var w = clientToWorld(e.clientX, e.clientY);
    var node = {
      id: "n_" + randomId(),
      type: type,
      params: defaultParams(type),
      position: { x: Math.round(w.x), y: Math.round(w.y) },
    };
    state.nodes.push(node);
    state.selection = [node.id];
    if (WF.renderAllNodes) WF.renderAllNodes();
    WF.scheduleSave();
  }

  function defaultParams(type) {
    var nt = state.catalogById[type];
    var out = {};
    if (nt && nt.params) {
      nt.params.forEach(function (p) {
        out[p.name] = p.default;
      });
    }
    return out;
  }

  // ---- Pointer gestures on the canvas ----

  function onCanvasMouseDown(e) {
    if (!state.ready) return;
    // Middle button → pan (grab to navigate), anywhere on the canvas — a
    // navigation gesture shouldn't depend on hitting empty space.
    if (e.button === 1) {
      e.preventDefault(); // suppress middle-click autoscroll
      startPan(e);
      return;
    }
    if (e.button !== 0) return;
    var t = e.target;
    if (!t || !t.closest) {
      startMarquee(e);
      return;
    }
    // The floating wire-delete button handles its own click — don't let the
    // mousedown fall through to a gesture handler (startNodeDrag/startMarquee),
    // whose e.preventDefault() would swallow the button's click.
    if (t.closest("#wfWireDelete")) return;
    // 1. Port dot → start a typed wire drag (wires satellite owns the gesture).
    var dot = t.closest(".wf-port-dot");
    if (dot) {
      if (WF.startWireDrag) WF.startWireDrag(e, dot);
      return;
    }
    // 2. Wire hit-target → select that edge (Delete-key / × button can remove it).
    var wireGroup = t.closest(".wf-wire-group");
    if (wireGroup) {
      if (WF.selectEdge) WF.selectEdge(wireGroup.getAttribute("data-edge-id"));
      return;
    }
    // 3. A param control → let it focus natively; do not drag/select/re-render,
    //    so typing into a card input isn't interrupted by a card rebuild.
    if (
      t.tagName === "INPUT" ||
      t.tagName === "SELECT" ||
      t.tagName === "TEXTAREA" ||
      t.closest(".wf-node-params")
    ) {
      return;
    }
    var card = t.closest(".wf-node");
    if (card) {
      startNodeDrag(e, card);
    } else {
      // Empty canvas → drag-select (marquee). Shift makes it additive; a bare
      // click is a zero-area non-additive marquee, which clears the selection.
      startMarquee(e);
    }
  }

  function startPan(e) {
    var canvas = qs("#wfCanvas");
    canvas.classList.add("panning");
    var startX = e.clientX;
    var startY = e.clientY;
    var origX = state.viewport.x;
    var origY = state.viewport.y;
    var moved = false;
    function move(ev) {
      state.viewport.x = origX + (ev.clientX - startX);
      state.viewport.y = origY + (ev.clientY - startY);
      moved = true;
      applyViewport();
    }
    function up() {
      document.removeEventListener("mousemove", move);
      document.removeEventListener("mouseup", up);
      canvas.classList.remove("panning");
      // Pan is a middle-button navigation gesture — it never changes the
      // selection (deselect-on-empty-click is handled by left-click → marquee).
      if (moved) WF.scheduleSave();
    }
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
  }

  function startNodeDrag(e, card) {
    var id = card.getAttribute("data-node-id");
    // Selecting a node clears any wire selection (mirror selectEdge clearing the
    // node selection), so Delete targets the node, not the previously-picked wire.
    state.selectedEdge = null;
    // Update selection on mousedown so a plain drag moves what you grabbed.
    if (e.shiftKey) {
      var at = state.selection.indexOf(id);
      if (at >= 0) state.selection.splice(at, 1);
      else state.selection.push(id);
    } else if (state.selection.indexOf(id) < 0) {
      state.selection = [id];
    }
    if (WF.renderAllNodes) WF.renderAllNodes();

    var startX = e.clientX;
    var startY = e.clientY;
    var zoom = state.viewport.zoom;
    var orig = {};
    state.selection.forEach(function (sid) {
      var n = findNode(sid);
      if (n) orig[sid] = { x: n.position.x || 0, y: n.position.y || 0 };
    });
    var moved = false;

    function move(ev) {
      var dx = (ev.clientX - startX) / zoom;
      var dy = (ev.clientY - startY) / zoom;
      state.selection.forEach(function (sid) {
        var n = findNode(sid);
        var o = orig[sid];
        if (n && o) {
          n.position.x = o.x + dx;
          n.position.y = o.y + dy;
        }
      });
      moved = true;
      scheduleNodePositionFlush();
    }
    function up() {
      document.removeEventListener("mousemove", move);
      document.removeEventListener("mouseup", up);
      if (moved) WF.scheduleSave();
    }
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
    e.preventDefault(); // suppress text selection / native drag
  }

  // RAF-throttled per-card style update (cf. the code-review canvas perf rule).
  function scheduleNodePositionFlush() {
    if (_moveRaf) return;
    _moveRaf = requestAnimationFrame(function () {
      _moveRaf = 0;
      state.selection.forEach(function (sid) {
        var n = findNode(sid);
        var card = qs('.wf-node[data-node-id="' + sid + '"]');
        if (n && card) {
          card.style.left = (n.position.x || 0) + "px";
          card.style.top = (n.position.y || 0) + "px";
        }
      });
      // Wires read live node.position (offsets stay cached — card internals are
      // unchanged), so a moved node's wires track it without a full re-render.
      if (WF.renderWires) WF.renderWires();
    });
  }

  function startMarquee(e) {
    var canvas = qs("#wfCanvas");
    var rect = canvas.getBoundingClientRect();
    var x0 = e.clientX - rect.left;
    var y0 = e.clientY - rect.top;
    var box = el("div", "wf-marquee");
    canvas.appendChild(box);

    function move(ev) {
      var x1 = ev.clientX - rect.left;
      var y1 = ev.clientY - rect.top;
      box.style.left = Math.min(x0, x1) + "px";
      box.style.top = Math.min(y0, y1) + "px";
      box.style.width = Math.abs(x1 - x0) + "px";
      box.style.height = Math.abs(y1 - y0) + "px";
    }
    function up(ev) {
      document.removeEventListener("mousemove", move);
      document.removeEventListener("mouseup", up);
      if (box.parentNode) box.parentNode.removeChild(box);
      selectInMarquee(rect, x0, y0, ev.clientX - rect.left, ev.clientY - rect.top, ev.shiftKey);
    }
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
    e.preventDefault();
  }

  // Select cards whose rendered rect intersects the marquee (client-coord test —
  // robust to zoom/pan without re-deriving world geometry).
  function selectInMarquee(rect, ax, ay, bx, by, additive) {
    var ml = Math.min(ax, bx);
    var mt = Math.min(ay, by);
    var mr = Math.max(ax, bx);
    var mb = Math.max(ay, by);
    var sel = additive ? state.selection.slice() : [];
    var cards = qsa(".wf-node");
    for (var i = 0; i < cards.length; i++) {
      var cr = cards[i].getBoundingClientRect();
      var cl = cr.left - rect.left;
      var ct = cr.top - rect.top;
      var crr = cr.right - rect.left;
      var crb = cr.bottom - rect.top;
      if (cl < mr && crr > ml && ct < mb && crb > mt) {
        var id = cards[i].getAttribute("data-node-id");
        if (sel.indexOf(id) < 0) sel.push(id);
      }
    }
    state.selectedEdge = null; // node selection wins over a wire selection
    state.selection = sel;
    if (WF.renderAllNodes) WF.renderAllNodes();
  }

  // ---- Zoom ----

  function onWheel(e) {
    if (!state.ready) return;
    e.preventDefault();
    var canvas = qs("#wfCanvas");
    var rect = canvas.getBoundingClientRect();
    var vp = state.viewport;
    var mx = e.clientX - rect.left;
    var my = e.clientY - rect.top;
    // World point under the cursor before the zoom.
    var wx = (mx - vp.x) / vp.zoom;
    var wy = (my - vp.y) / vp.zoom;
    var factor = e.deltaY < 0 ? 1.1 : 1 / 1.1;
    vp.zoom = clamp(vp.zoom * factor, ZOOM_MIN, ZOOM_MAX);
    // Re-pin (wx,wy) under the cursor: x = mx - wx*zoom.
    vp.x = mx - wx * vp.zoom;
    vp.y = my - wy * vp.zoom;
    applyViewport();
    WF.scheduleSave();
  }

  // ---- Keyboard ----

  // ---- Clipboard (copy / paste / duplicate) ----

  // In-memory clipboard ({nodes, edges}); not the system clipboard (vanilla JS,
  // no async-clipboard permission dance, and node graphs aren't text anyway).
  var _clipboard = null;

  // Capture the current selection + induced edges (both endpoints selected) as a
  // deep-cloned sub-graph. Returns false when nothing is selected.
  function copySelection() {
    var sel = state.selection || [];
    if (!sel.length) return false;
    var selSet = {};
    sel.forEach(function (id) {
      selSet[id] = true;
    });
    var nodes = (state.nodes || [])
      .filter(function (n) {
        return selSet[n.id];
      })
      .map(function (n) {
        return JSON.parse(JSON.stringify(n));
      });
    if (!nodes.length) return false;
    var edges = (state.edges || [])
      .filter(function (ed) {
        return selSet[ed.from] && selSet[ed.to];
      })
      .map(function (ed) {
        return JSON.parse(JSON.stringify(ed));
      });
    _clipboard = { nodes: nodes, edges: edges };
    return true;
  }

  // Stamp the clipboard onto the canvas (fresh ids, cascaded offset). The stashes
  // satellite owns the id-remap; call it late-bound so load order doesn't matter.
  function pasteClipboard() {
    if (!_clipboard || !_clipboard.nodes.length || !WF.instantiateSubgraph) {
      return false;
    }
    WF.instantiateSubgraph(_clipboard.nodes, _clipboard.edges, null);
    return true;
  }

  function duplicateSelection() {
    return copySelection() && pasteClipboard();
  }

  function onKeyDown(e) {
    if (!state.ready) return;
    var t = e.target;
    var inField =
      t &&
      (t.tagName === "INPUT" ||
        t.tagName === "TEXTAREA" ||
        t.tagName === "SELECT" ||
        t.isContentEditable);
    // Clipboard shortcuts — skipped while typing in a field so the browser's
    // native Ctrl/Cmd+C/V/D still work there. Cmd on macOS, Ctrl elsewhere.
    if ((e.metaKey || e.ctrlKey) && !inField) {
      var k = (e.key || "").toLowerCase();
      if (k === "c") {
        if (copySelection()) e.preventDefault();
        return;
      }
      if (k === "v") {
        if (pasteClipboard()) e.preventDefault();
        return;
      }
      if (k === "d") {
        if (duplicateSelection()) e.preventDefault();
        return;
      }
    }
    if (e.key !== "Delete" && e.key !== "Backspace") return;
    if (inField) return;
    // A selected wire takes priority over node selection (wires satellite owns
    // single-edge removal, shared with the floating × button).
    if (state.selectedEdge) {
      e.preventDefault();
      if (WF.removeEdge) WF.removeEdge(state.selectedEdge);
      return;
    }
    if (!state.selection.length) return;
    e.preventDefault();
    var drop = {};
    state.selection.forEach(function (id) {
      drop[id] = true;
    });
    state.nodes = state.nodes.filter(function (n) {
      return !drop[n.id];
    });
    // Drop any wire touching a deleted node so no dangling edge survives.
    state.edges = state.edges.filter(function (ed) {
      return !drop[ed.from] && !drop[ed.to];
    });
    state.selection = [];
    if (WF.renderAllNodes) WF.renderAllNodes();
    WF.scheduleSave();
  }

  // ---- Auto-arrange ("Clean up") ----

  // Lay the graph out left→right in dependency layers: a node sits one column
  // right of its deepest upstream node, stacked by current vertical order within
  // the column. Bounded relaxation keeps it cycle-safe (M2 doesn't reject cycles
  // yet). Resets the viewport so the tidied graph is visible from the origin.
  function autoArrange() {
    if (!state.ready) return;
    var nodes = state.nodes || [];
    if (!nodes.length) return;
    // Re-layout rebuilds every card and moves the source node, so cancel any
    // in-flight wire gesture rather than leave it armed with a stale highlight.
    if (WF.cancelConnect) WF.cancelConnect();
    var edges = state.edges || [];

    var layer = {};
    nodes.forEach(function (n) {
      layer[n.id] = 0;
    });
    // Longest-path layering; capped at node count so a cycle can't loop forever.
    for (var iter = 0; iter < nodes.length; iter++) {
      var changed = false;
      edges.forEach(function (e) {
        if (layer[e.from] === undefined || layer[e.to] === undefined) return;
        if (layer[e.to] < layer[e.from] + 1) {
          layer[e.to] = layer[e.from] + 1;
          changed = true;
        }
      });
      if (!changed) break;
    }

    var byLayer = {};
    nodes.forEach(function (n) {
      var L = layer[n.id] || 0;
      (byLayer[L] = byLayer[L] || []).push(n);
    });

    var COL = 280;
    var ROW = 180;
    Object.keys(byLayer).forEach(function (key) {
      var col = byLayer[key];
      // Preserve current top-to-bottom order so the tidy feels stable.
      col.sort(function (a, b) {
        return (a.position.y || 0) - (b.position.y || 0);
      });
      var L = parseInt(key, 10);
      col.forEach(function (n, idx) {
        n.position.x = L * COL;
        n.position.y = idx * ROW;
      });
    });

    state.viewport = { x: 40, y: 40, zoom: 1 };
    state.selection = [];
    state.selectedEdge = null;
    if (WF.renderAllNodes) WF.renderAllNodes();
    if (WF.applyViewport) WF.applyViewport();
    WF.scheduleSave();
  }

  // ---- Focus a node (from the validation panel) ----

  // Select a node and pan it to the canvas centre — the Issues-panel rows call
  // this so clicking a finding reveals the offending card.
  function focusNode(id) {
    var node = findNode(id);
    if (!node) return;
    state.selectedEdge = null;
    state.selection = [id];
    if (WF.renderAllNodes) WF.renderAllNodes();
    var canvas = qs("#wfCanvas");
    if (!canvas) return;
    var rect = canvas.getBoundingClientRect();
    var card = qs('.wf-node[data-node-id="' + id + '"]');
    var vp = state.viewport;
    var w = card ? card.offsetWidth : 200;
    var h = card ? card.offsetHeight : 120;
    var cx = (node.position.x || 0) + w / 2;
    var cy = (node.position.y || 0) + h / 2;
    vp.x = rect.width / 2 - cx * vp.zoom;
    vp.y = rect.height / 2 - cy * vp.zoom;
    applyViewport();
  }

  // ---- Boot ----

  function initCanvas() {
    if (_bound) return;
    var canvas = qs("#wfCanvas");
    if (!canvas) return;
    _bound = true;

    var gridVar = getComputedStyle(document.documentElement).getPropertyValue(
      "--wf-grid-size"
    );
    var parsed = parseFloat(gridVar);
    if (parsed) _gridBase = parsed;

    canvas.addEventListener("dragover", onDragOver);
    canvas.addEventListener("drop", onDrop);
    canvas.addEventListener("mousedown", onCanvasMouseDown);
    canvas.addEventListener("wheel", onWheel, { passive: false });
    document.addEventListener("keydown", onKeyDown);
  }

  WF.initCanvas = initCanvas;
  WF.applyViewport = applyViewport;
  WF.autoArrange = autoArrange;
  // Consumed by the wires satellite (cursor→world for the in-flight wire; node
  // lookup for port endpoints).
  WF.clientToWorld = clientToWorld;
  WF.findNode = findNode;
  // Consumed by the stashes satellite (fresh node ids on instantiate).
  WF.randomId = randomId;
  // Consumed by the validation satellite (reveal a node from an Issues row).
  WF.focusNode = focusNode;
})();
