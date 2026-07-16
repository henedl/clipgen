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
  var _minimapRaf = 0;
  // Node-drag / auto-pan state (set in startNodeDrag, cleared on mouseup).
  var _draggingNode = false;
  var _dragRect = null; // canvas getBoundingClientRect cached at drag start
  var _dragCursorX = 0;
  var _dragCursorY = 0;
  var _dragOffsets = null; // { nodeId: {x, y} } world offset from cursor grab point

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
      writeViewport();
    });
  }

  // Synchronously write the world transform + grid offset from the current
  // viewport. Split out of applyViewport so the node-drag flush (already in a
  // RAF) can apply an auto-pan transform in the SAME frame it positions the
  // cards — deferring it would paint cards one pan-step ahead of the world.
  function writeViewport() {
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
    // Pan/zoom moved the viewport rectangle — redraw the minimap to match.
    renderMinimap();
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
    // Likewise the floating minimap + its zoom controls: let them handle their
    // own clicks (the minimap canvas also stops propagation) rather than starting
    // a marquee that would clear the selection.
    if (t.closest("#wfMinimapWrap")) return;
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

    // Anchor the drag in world space: store each selected node's offset from the
    // cursor's world point at grab time, then on every frame set its position to
    // (cursorWorld + offset). World-anchoring (vs a screen-pixel delta) is what
    // lets the node keep trailing the cursor while auto-pan scrolls the viewport
    // — a pure screen delta would drift the node away from the cursor by the pan.
    var canvasEl = qs("#wfCanvas");
    _dragRect = canvasEl ? canvasEl.getBoundingClientRect() : null;
    _dragCursorX = e.clientX;
    _dragCursorY = e.clientY;
    var grab = clientToWorld(e.clientX, e.clientY);
    _dragOffsets = {};
    state.selection.forEach(function (sid) {
      var n = findNode(sid);
      if (n) {
        _dragOffsets[sid] = {
          x: (n.position.x || 0) - grab.x,
          y: (n.position.y || 0) - grab.y,
        };
      }
    });
    _draggingNode = true;
    var moved = false;

    function move(ev) {
      _dragCursorX = ev.clientX;
      _dragCursorY = ev.clientY;
      moved = true;
      scheduleNodePositionFlush();
    }
    function up() {
      document.removeEventListener("mousemove", move);
      document.removeEventListener("mouseup", up);
      _draggingNode = false;
      _dragRect = null;
      _dragOffsets = null;
      if (moved) WF.scheduleSave();
    }
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
    e.preventDefault(); // suppress text selection / native drag
  }

  // When a node drag reaches an edge band of the canvas, scroll the viewport
  // away from that edge so the drag can continue past the visible bounds.
  // Returns true if it nudged (the flush re-arms next frame while it does, so a
  // drag held still at the edge keeps scrolling). Uses the rect cached at drag
  // start per the canvas-perf rule. Only mutates the viewport; the caller writes
  // the transform synchronously (see scheduleNodePositionFlush) so it stays in
  // step with the card positions painted the same frame.
  var EDGE_BAND = 40; // px from the canvas edge that arms auto-pan
  var EDGE_STEP = 12; // px/frame the viewport scrolls
  function autoPanWhileDragging() {
    if (!_draggingNode || !_dragRect) return false;
    var vp = state.viewport;
    var nudged = false;
    if (_dragCursorX - _dragRect.left < EDGE_BAND) {
      vp.x += EDGE_STEP;
      nudged = true;
    } else if (_dragRect.right - _dragCursorX < EDGE_BAND) {
      vp.x -= EDGE_STEP;
      nudged = true;
    }
    if (_dragCursorY - _dragRect.top < EDGE_BAND) {
      vp.y += EDGE_STEP;
      nudged = true;
    } else if (_dragRect.bottom - _dragCursorY < EDGE_BAND) {
      vp.y -= EDGE_STEP;
      nudged = true;
    }
    return nudged;
  }

  // RAF-throttled per-card style update (cf. the code-review canvas perf rule).
  // Recomputes dragged-node positions from the live cursor world point so they
  // track the cursor across auto-pan; updates only the moved cards + wires.
  function scheduleNodePositionFlush() {
    if (_moveRaf) return;
    _moveRaf = requestAnimationFrame(function () {
      _moveRaf = 0;
      var panned = autoPanWhileDragging();
      // Apply the auto-panned transform now, in this same frame, so the world
      // transform and the card positions below are computed from one viewport
      // value. Deferring via applyViewport() would lag the world a frame behind
      // the cards, leaving dragged nodes a pan-step (EDGE_STEP) off the cursor.
      if (panned) writeViewport();
      if (_dragRect && _dragOffsets) {
        var vp = state.viewport;
        // Cursor world point (post-pan), using the cached rect to avoid a layout
        // read; mirrors clientToWorld's math.
        var curX = (_dragCursorX - _dragRect.left - vp.x) / vp.zoom;
        var curY = (_dragCursorY - _dragRect.top - vp.y) / vp.zoom;
        state.selection.forEach(function (sid) {
          var n = findNode(sid);
          var o = _dragOffsets[sid];
          var card = qs('.wf-node[data-node-id="' + sid + '"]');
          if (n && o) {
            n.position.x = curX + o.x;
            n.position.y = curY + o.y;
          }
          if (n && card) {
            card.style.left = (n.position.x || 0) + "px";
            card.style.top = (n.position.y || 0) + "px";
          }
        });
      }
      // Wires read live node.position (offsets stay cached — card internals are
      // unchanged), so a moved node's wires track it without a full re-render.
      if (WF.renderWires) WF.renderWires();
      renderMinimap();
      // Keep panning while a drag is held inside the edge band even if the
      // cursor isn't moving (move() won't re-fire). Bounded to the drag.
      if (panned && _draggingNode) scheduleNodePositionFlush();
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

  // Zoom by `factor` about the canvas centre (the minimap +/- buttons; mirrors
  // onWheel's re-pin math but anchored to the viewport centre rather than the
  // cursor). > 1 zooms in, < 1 out; clamped to [ZOOM_MIN, ZOOM_MAX].
  function zoomAtCenter(factor) {
    if (!state.ready) return;
    var canvas = qs("#wfCanvas");
    if (!canvas) return;
    var rect = canvas.getBoundingClientRect();
    var vp = state.viewport;
    var cx = rect.width / 2;
    var cy = rect.height / 2;
    var wx = (cx - vp.x) / vp.zoom;
    var wy = (cy - vp.y) / vp.zoom;
    vp.zoom = clamp(vp.zoom * factor, ZOOM_MIN, ZOOM_MAX);
    vp.x = cx - wx * vp.zoom;
    vp.y = cy - wy * vp.zoom;
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

  function _canvasReady() {
    return !!state.ready;
  }

  function deleteSelection() {
    // A selected wire takes priority over node selection (wires satellite owns
    // single-edge removal, shared with the floating × button).
    if (state.selectedEdge) {
      if (WF.removeEdge) WF.removeEdge(state.selectedEdge);
      return;
    }
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

  // Clipboard/undo handlers return their function's boolean so a false (empty
  // selection / empty clipboard / empty history) declines the event and native
  // copy/paste stays intact. The hub owns the history stack; call late-bound.
  function initCanvasHotkeys() {
    window.ClipgenHotkeys.register([
      { id: "workflows.copy", when: _canvasReady, handler: function () { return copySelection(); } },
      { id: "workflows.paste", when: _canvasReady, handler: function () { return pasteClipboard(); } },
      { id: "workflows.duplicate", when: _canvasReady, handler: function () { return duplicateSelection(); } },
      { id: "edit.undo", when: _canvasReady, handler: function () { return !!(WF.undo && WF.undo()); } },
      { id: "edit.redo", when: _canvasReady, handler: function () { return !!(WF.redo && WF.redo()); } },
      { id: "workflows.fitView", when: _canvasReady, handler: function () { fitToView(); } },
      {
        id: "workflows.deleteSelection",
        when: function () {
          return _canvasReady() && !!(state.selectedEdge || state.selection.length);
        },
        handler: function () { deleteSelection(); },
      },
      {
        id: "global.primary",
        when: function () {
          var btn = qs("#wfRunBtn");
          return _canvasReady() && !!(btn && !btn.disabled);
        },
        handler: function () { qs("#wfRunBtn").click(); },
      },
      { id: "workflows.note.pan" },
      { id: "workflows.note.zoom" },
      { id: "workflows.note.select" },
    ]);
  }

  // ---- Auto-arrange ("Clean up") ----

  // Lay the graph out left→right in dependency layers: a node sits one column
  // right of its deepest upstream node, stacked by current vertical order within
  // the column. Bounded relaxation keeps it cycle-safe (validation blocks Run on
  // cycles, but Clean up can fire on an unvalidated graph). Resets the viewport
  // so the tidied graph is visible from the origin.
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

  // ---- Fit to view ----

  // World-space bounding box of every node card. Card pixel size is read the way
  // focusNode does (offsetWidth/Height, default 200×120 for an unrendered card).
  // Shared by fitToView and the minimap. Returns null when there are no nodes.
  function nodesBoundingBox() {
    var nodes = state.nodes || [];
    if (!nodes.length) return null;
    var minX = Infinity,
      minY = Infinity,
      maxX = -Infinity,
      maxY = -Infinity;
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      var x = (n.position && n.position.x) || 0;
      var y = (n.position && n.position.y) || 0;
      var card = qs('.wf-node[data-node-id="' + n.id + '"]');
      var w = card ? card.offsetWidth : 200;
      var h = card ? card.offsetHeight : 120;
      if (x < minX) minX = x;
      if (y < minY) minY = y;
      if (x + w > maxX) maxX = x + w;
      if (y + h > maxY) maxY = y + h;
    }
    return {
      minX: minX,
      minY: minY,
      maxX: maxX,
      maxY: maxY,
      w: maxX - minX,
      h: maxY - minY,
    };
  }

  // Frame all nodes in the viewport: zoom to fit the bounding box (with padding,
  // clamped) and centre it. Mirrors focusNode's centring math. No-op when not
  // ready or empty (cf. autoArrange's guards).
  function fitToView() {
    if (!state.ready) return;
    var box = nodesBoundingBox();
    if (!box) return;
    var canvas = qs("#wfCanvas");
    if (!canvas) return;
    var rect = canvas.getBoundingClientRect();
    var pad = 40;
    var vw = rect.width - pad * 2;
    var vh = rect.height - pad * 2;
    // Guard the degenerate single-node (zero-size) box and tiny viewports.
    var zoomX = box.w > 0 ? vw / box.w : ZOOM_MAX;
    var zoomY = box.h > 0 ? vh / box.h : ZOOM_MAX;
    var zoom = clamp(Math.min(zoomX, zoomY), ZOOM_MIN, ZOOM_MAX);
    var cx = box.minX + box.w / 2;
    var cy = box.minY + box.h / 2;
    var vp = state.viewport;
    vp.zoom = zoom;
    vp.x = rect.width / 2 - cx * zoom;
    vp.y = rect.height / 2 - cy * zoom;
    applyViewport();
    WF.scheduleSave();
  }

  // ---- Minimap ----

  // Draw the corner minimap: a scaled-down, zoomed-out mirror of the canvas view.
  // Node rects scale with vp.zoom (contents zoom with the canvas), but the minimap
  // shows OVERVIEW_FACTOR× the visible area so nodes just outside the frame stay
  // visible, and the view frame is free to drift toward the graph edges — the
  // camera clamps to the graph bounds, so when the whole graph fits the frame
  // roams freely and when zoomed in it follows the viewport with edge drift rather
  // than staying pinned centre. RAF-throttled (mirrors applyViewport's _vpRaf).
  // Hidden when there are no nodes or the tab is backgrounded (canvas perf rule).
  var OVERVIEW_FACTOR = 2.5; // how many viewports the minimap spans (zoom-out)
  var GRAPH_MARGIN = 60; // world px of breathing room kept around the graph
  function renderMinimap() {
    if (_minimapRaf) return;
    _minimapRaf = requestAnimationFrame(function () {
      _minimapRaf = 0;
      var mm = qs("#wfMinimap");
      if (!mm) return;
      // The wrap (minimap + zoom controls) owns visibility; fall back to the
      // minimap canvas if the wrap markup is absent.
      var hideEl = qs("#wfMinimapWrap") || mm;
      var canvas = qs("#wfCanvas");
      var nodes = state.nodes || [];
      var box = nodesBoundingBox();
      if (!canvas || !box || document.hidden) {
        hideEl.classList.add("hidden");
        return;
      }
      var rect = canvas.getBoundingClientRect();
      if (!rect.width || !rect.height) {
        hideEl.classList.add("hidden");
        return;
      }
      hideEl.classList.remove("hidden");
      var ctx = mm.getContext("2d");
      if (!ctx) return;
      var mmW = mm.width;
      var mmH = mm.height;
      ctx.clearRect(0, 0, mmW, mmH);

      var vp = state.viewport;
      var pad = 8;
      // world → minimap scale: mirror the canvas (∝ vp.zoom so node rects zoom
      // with it), then divide by OVERVIEW_FACTOR to zoom the whole minimap out.
      var mFit = Math.min((mmW - pad * 2) / rect.width, (mmH - pad * 2) / rect.height);
      var scale = (vp.zoom * mFit) / OVERVIEW_FACTOR;

      // Camera: follow the viewport centre, clamped in two passes.
      // 1. Soft — keep the minimap window over the graph (+ margin) for context
      //    and "play": clampCenter centres on the graph when the window is larger
      //    than it (frame roams freely), else follows with drift near the edges.
      // 2. Hard — keep the view frame fully inside the minimap: clampFrameInside
      //    pans the camera so the frame never clips the minimap edge, overriding
      //    the graph clamp at the outermost edges (showing a little empty space).
      var halfWx = mmW / (2 * scale);
      var halfWy = mmH / (2 * scale);
      var viewCx = (rect.width / 2 - vp.x) / vp.zoom;
      var viewCy = (rect.height / 2 - vp.y) / vp.zoom;
      var frameWW = rect.width / vp.zoom; // view frame world size
      var frameWH = rect.height / vp.zoom;
      var framePad = pad / scale; // keep the frame `pad` minimap-px from the edge
      var camX = clampCenter(viewCx, box.minX - GRAPH_MARGIN, box.maxX + GRAPH_MARGIN, halfWx);
      var camY = clampCenter(viewCy, box.minY - GRAPH_MARGIN, box.maxY + GRAPH_MARGIN, halfWy);
      camX = clampFrameInside(camX, viewCx, frameWW, halfWx, framePad);
      camY = clampFrameInside(camY, viewCy, frameWH, halfWy, framePad);
      var offX = mmW / 2 - camX * scale;
      var offY = mmH / 2 - camY * scale;
      // Stash the transform so click/drag can invert it back to world coords.
      _mmTransform = { scale: scale, offX: offX, offY: offY };

      var styles = getComputedStyle(document.documentElement);
      var nodeFill = (styles.getPropertyValue("--color-text-muted") || "#888").trim();
      var vpStroke = (styles.getPropertyValue("--color-accent") || "#4a9").trim();

      // Node rects (any falling outside the minimap are clipped by the canvas).
      ctx.fillStyle = nodeFill;
      for (var i = 0; i < nodes.length; i++) {
        var n = nodes[i];
        var nx = ((n.position && n.position.x) || 0) * scale + offX;
        var ny = ((n.position && n.position.y) || 0) * scale + offY;
        var card = qs('.wf-node[data-node-id="' + n.id + '"]');
        var nw = (card ? card.offsetWidth : 200) * scale;
        var nh = (card ? card.offsetHeight : 120) * scale;
        ctx.fillRect(nx, ny, Math.max(2, nw), Math.max(2, nh));
      }

      // View frame: the visible world rect mapped through the minimap transform.
      // Kept fully on-screen (never clipping the edge) by the camera clamp above.
      var wx = -vp.x / vp.zoom;
      var wy = -vp.y / vp.zoom;
      ctx.strokeStyle = vpStroke;
      ctx.lineWidth = 1.5;
      ctx.strokeRect(wx * scale + offX, wy * scale + offY, frameWW * scale, frameWH * scale);
    });
  }

  // Centre a minimap camera on `c`, clamped so a window of half-size `half` stays
  // within [lo, hi]; if that window is wider than the span, centre on the span.
  function clampCenter(c, lo, hi, half) {
    if (hi - lo <= 2 * half) return (lo + hi) / 2;
    return clamp(c, lo + half, hi - half);
  }

  // Constrain a minimap camera `cam` so the view frame ([viewC ± frameW/2], world)
  // stays fully inside the window ([cam ± half], minus a `padW` world margin). If
  // the frame is wider than the window, just centre it. This is the hard rule that
  // prevents the view frame from drifting off the edge of the minimap.
  function clampFrameInside(cam, viewC, frameW, half, padW) {
    var lo = viewC + frameW / 2 - half + padW;
    var hi = viewC - frameW / 2 + half - padW;
    if (lo > hi) return viewC;
    return clamp(cam, lo, hi);
  }

  // Recenter the viewport on the world point under a minimap click/drag.
  var _mmTransform = null;
  var _mmDragging = false;
  function minimapRecenter(ev) {
    var mm = qs("#wfMinimap");
    var canvas = qs("#wfCanvas");
    if (!mm || !canvas || !_mmTransform) return;
    var mmRect = mm.getBoundingClientRect();
    var t = _mmTransform;
    var worldX = (ev.clientX - mmRect.left - t.offX) / t.scale;
    var worldY = (ev.clientY - mmRect.top - t.offY) / t.scale;
    var rect = canvas.getBoundingClientRect();
    var vp = state.viewport;
    vp.x = rect.width / 2 - worldX * vp.zoom;
    vp.y = rect.height / 2 - worldY * vp.zoom;
    applyViewport();
  }

  function onMinimapMouseDown(e) {
    if (!state.ready) return;
    e.preventDefault();
    // The minimap sits inside #wfCanvas — stop the mousedown bubbling to
    // onCanvasMouseDown, which would otherwise start a marquee selection.
    e.stopPropagation();
    _mmDragging = true;
    minimapRecenter(e);
    function move(ev) {
      if (_mmDragging) minimapRecenter(ev);
    }
    function up() {
      document.removeEventListener("mousemove", move);
      document.removeEventListener("mouseup", up);
      _mmDragging = false;
      WF.scheduleSave();
    }
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
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
    initCanvasHotkeys();
    var minimap = qs("#wfMinimap");
    if (minimap) minimap.addEventListener("mousedown", onMinimapMouseDown);
    // Redraw the minimap when the tab returns to the foreground (it skips
    // drawing while hidden).
    document.addEventListener("visibilitychange", function () {
      if (!document.hidden) renderMinimap();
    });
  }

  WF.initCanvas = initCanvas;
  WF.applyViewport = applyViewport;
  WF.autoArrange = autoArrange;
  // Consumed by the hub toolbar (button + "F" shortcut) and the minimap sync
  // hook in the nodes satellite (renderAllNodes → structural changes).
  WF.fitToView = fitToView;
  // Consumed by the hub's minimap zoom +/- buttons.
  WF.zoomAtCenter = zoomAtCenter;
  WF.renderMinimap = renderMinimap;
  // Consumed by the wires satellite (cursor→world for the in-flight wire; node
  // lookup for port endpoints).
  WF.clientToWorld = clientToWorld;
  WF.findNode = findNode;
  // Consumed by the stashes satellite (fresh node ids on instantiate).
  WF.randomId = randomId;
  // Consumed by the validation satellite (reveal a node from an Issues row).
  WF.focusNode = focusNode;
})();
