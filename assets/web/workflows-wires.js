/* Workflows — SVG wire layer (satellite of workflows.js).
 *
 * Owns the typed connector edges: connect from a port either by dragging or by
 * clicking once to arm (the wire then trails the cursor until the next click —
 * on a card it connects, anywhere else it cancels); dropping onto a card snaps
 * to its nearest compatible port. While connecting, compatible cards + palette
 * items glow and the rest dim. Type validation is exact-match (the canConnect()
 * seam widens to the ADAPTERS table in M3). Plus bezier rendering in the
 * transformed #wfWires layer, edge selection, and removal (Delete key via the
 * canvas + a floating × button).
 *
 * Edge shape: { id, from, fromPort, to, toPort } — always normalized so `from`
 * is the output side and `to` the input side. Endpoints are computed from the
 * live node.position (state) plus a cached per-port offset measured once from
 * the card DOM, so wires track a node during a drag without re-reading layout.
 *
 * Reads shared state through WF.state; reaches the canvas/hub late-bound via
 * WF.* (load-order-safe — this file loads after workflows-canvas.js).
 */

(function () {
  "use strict";

  var WF = window.ClipgenWorkflows;
  var state = WF.state;

  var SVG_NS = "http://www.w3.org/2000/svg";

  var MOVE_THRESHOLD = 4; // px before a port mousedown counts as a drag, not a click

  // nodeId|port|dir -> {x,y} dot-centre offset within its card (zoom-1 px).
  // Invalidated by clearPortCache() whenever the port DOM is rebuilt.
  var _portOffsets = {};
  // The in-flight connection (drag or click-armed), or null. Shape:
  // { src, temp(<path>), moved(bool), mode("drag"|"armed"), raf, lastX, lastY }.
  var _connect = null;
  var _hoverEdge = null; // edge id currently hovered (drives the × button)
  var _hideTimer = 0; // grace timer so the pointer can reach the × button
  var _bound = false;

  function svgEl(tag, cls) {
    var node = document.createElementNS(SVG_NS, tag);
    if (cls) node.setAttribute("class", cls);
    return node;
  }

  function findEdge(id) {
    var edges = state.edges || [];
    for (var i = 0; i < edges.length; i++) {
      if (edges[i].id === id) return edges[i];
    }
    return null;
  }

  // ---- Geometry ----

  function clearPortCache() {
    _portOffsets = {};
  }

  // World position of a port dot's centre = node.position + cached dot offset.
  function portWorldPos(nodeId, port, dir) {
    var node = WF.findNode(nodeId);
    if (!node || !node.position) return null;
    var key = nodeId + "|" + port + "|" + dir;
    var off = _portOffsets[key];
    if (!off) {
      var dot = qs(
        '.wf-node[data-node-id="' +
          nodeId +
          '"] .wf-port-dot[data-port="' +
          port +
          '"][data-port-dir="' +
          dir +
          '"]'
      );
      if (!dot) return null;
      var card = dot.closest(".wf-node");
      if (!card) return null;
      // Measure relative to the card and divide by zoom → zoom-1 layout offset,
      // exact across borders/padding. Cached, so this layout read happens once
      // per renderAllNodes, not per drag frame.
      var dotRect = dot.getBoundingClientRect();
      var cardRect = card.getBoundingClientRect();
      var zoom = state.viewport.zoom || 1;
      off = _portOffsets[key] = {
        x: (dotRect.left + dotRect.width / 2 - cardRect.left) / zoom,
        y: (dotRect.top + dotRect.height / 2 - cardRect.top) / zoom,
      };
    }
    return { x: (node.position.x || 0) + off.x, y: (node.position.y || 0) + off.y };
  }

  // Horizontal S-curve from an output (a, bulges right) to an input (b, left).
  function bezierPath(a, b) {
    var dx = Math.max(40, Math.abs(b.x - a.x) * 0.4);
    return (
      "M " +
      a.x +
      " " +
      a.y +
      " C " +
      (a.x + dx) +
      " " +
      a.y +
      ", " +
      (b.x - dx) +
      " " +
      b.y +
      ", " +
      b.x +
      " " +
      b.y
    );
  }

  // ---- Rendering ----

  function renderWires() {
    var svg = qs("#wfWires");
    if (!svg) return;
    while (svg.firstChild) svg.removeChild(svg.firstChild);
    var frag = document.createDocumentFragment();
    (state.edges || []).forEach(function (edge) {
      var a = portWorldPos(edge.from, edge.fromPort, "out");
      var b = portWorldPos(edge.to, edge.toPort, "in");
      if (!a || !b) return;
      var d = bezierPath(a, b);
      var group = svgEl("g", "wf-wire-group");
      group.setAttribute("data-edge-id", edge.id);
      var hit = svgEl("path", "wf-wire-hit");
      hit.setAttribute("d", d);
      var wire = svgEl(
        "path",
        "wf-wire" + (state.selectedEdge === edge.id ? " selected" : "")
      );
      wire.setAttribute("d", d);
      group.appendChild(hit);
      group.appendChild(wire);
      frag.appendChild(group);
    });
    svg.appendChild(frag);
    // A connection is in flight — re-attach its temp wire (the clear above
    // dropped it). renderWires isn't normally called mid-connect, but be safe.
    if (_connect && _connect.temp) svg.appendChild(_connect.temp);
    refreshWireDelete();
  }

  // ---- Type validation ----

  function canConnect(outType, inType) {
    // M2: exact type match. M3 widens this to also accept a registered adapter
    // (events → clipRecords, segments → timeRange, …) once the backend serves
    // the ADAPTERS table to the frontend.
    return outType === inType;
  }

  // ---- Connect: highlight ----

  // Does node type `typeId` expose a port that could receive/feed `source`?
  function typeHasCompatiblePort(typeId, source) {
    var nt = state.catalogById[typeId];
    if (!nt) return false;
    if (source.dir === "out") {
      return (nt.inputs || []).some(function (p) {
        return canConnect(source.type, p.type);
      });
    }
    return (nt.outputs || []).some(function (p) {
      return canConnect(p.type, source.type);
    });
  }

  // Glow the cards + palette items that can accept this connection; dim the rest
  // (the source card stays neutral). Cleared by clearConnectHighlight().
  function applyConnectHighlight(source) {
    var cards = qsa(".wf-node");
    for (var i = 0; i < cards.length; i++) {
      var id = cards[i].getAttribute("data-node-id");
      if (id === source.nodeId) {
        cards[i].classList.remove("wf-compatible", "wf-dim");
        continue;
      }
      var node = WF.findNode(id);
      var ok = !!(node && typeHasCompatiblePort(node.type, source));
      cards[i].classList.toggle("wf-compatible", ok);
      cards[i].classList.toggle("wf-dim", !ok);
    }
    var items = qsa(".wf-palette-item");
    for (var j = 0; j < items.length; j++) {
      var typeId = items[j].getAttribute("data-node-type");
      var pok = !!(typeId && typeHasCompatiblePort(typeId, source));
      items[j].classList.toggle("wf-compatible", pok);
      items[j].classList.toggle("wf-dim", !pok);
    }
  }

  function clearConnectHighlight() {
    var marked = qsa(".wf-compatible, .wf-dim");
    for (var i = 0; i < marked.length; i++) {
      marked[i].classList.remove("wf-compatible", "wf-dim");
    }
  }

  // ---- Connect: gesture (drag OR click-to-arm) ----

  function portInfo(dot) {
    var card = dot.closest(".wf-node");
    if (!card) return null;
    return {
      nodeId: card.getAttribute("data-node-id"),
      port: dot.getAttribute("data-port"),
      type: dot.getAttribute("data-port-type"),
      dir: dot.getAttribute("data-port-dir"),
    };
  }

  function drawTempWire(clientX, clientY) {
    if (!_connect || !_connect.temp) return;
    var srcPos = portWorldPos(_connect.src.nodeId, _connect.src.port, _connect.src.dir);
    if (!srcPos) return;
    var cur = WF.clientToWorld(clientX, clientY);
    // The curve always flows output→input; orient by which end the source is.
    var from = _connect.src.dir === "out" ? srcPos : cur;
    var to = _connect.src.dir === "out" ? cur : srcPos;
    _connect.temp.setAttribute("d", bezierPath(from, to));
  }

  function scheduleTempDraw(clientX, clientY) {
    if (!_connect) return;
    _connect.lastX = clientX;
    _connect.lastY = clientY;
    if (_connect.raf) return;
    _connect.raf = requestAnimationFrame(function () {
      if (!_connect) return;
      _connect.raf = 0;
      drawTempWire(_connect.lastX, _connect.lastY); // latest cursor, not frame-start
    });
  }

  function beginConnect(src) {
    var svg = qs("#wfWires");
    if (!svg) return false;
    var temp = svgEl("path", "wf-wire wf-wire-temp");
    svg.appendChild(temp);
    _connect = { src: src, temp: temp, moved: false, mode: "drag", raf: 0 };
    applyConnectHighlight(src);
    return true;
  }

  // Tear down the in-flight connection (temp wire, highlight, armed listeners).
  function endConnect() {
    if (_connect) {
      if (_connect.raf) cancelAnimationFrame(_connect.raf);
      if (_connect.temp && _connect.temp.parentNode) {
        _connect.temp.parentNode.removeChild(_connect.temp);
      }
    }
    document.removeEventListener("mousedown", onArmedMouseDown, true);
    document.removeEventListener("mousemove", onArmedMove);
    document.removeEventListener("keydown", onArmedKey, true);
    clearConnectHighlight();
    _connect = null;
  }

  // Resolve at a screen point: connect to the card there (snapping to its nearest
  // compatible port), or — if it's empty space / the source node — just cancel.
  function finishConnect(clientX, clientY) {
    if (!_connect) return;
    var src = _connect.src;
    var elAt = document.elementFromPoint(clientX, clientY);
    endConnect();
    if (!elAt || !elAt.closest) return;
    var card = elAt.closest(".wf-node");
    if (card && card.getAttribute("data-node-id") !== src.nodeId) {
      connectToCard(src, card, clientX, clientY);
    }
  }

  // Move to click-armed mode: the wire trails the cursor until the next mousedown
  // (handled in capture so it pre-empts pan / node-drag / palette handlers).
  function armConnect() {
    if (!_connect) return;
    _connect.mode = "armed";
    document.addEventListener("mousemove", onArmedMove);
    document.addEventListener("mousedown", onArmedMouseDown, true);
    document.addEventListener("keydown", onArmedKey, true);
  }

  function onArmedMove(ev) {
    scheduleTempDraw(ev.clientX, ev.clientY);
  }

  function onArmedMouseDown(ev) {
    if (ev.button !== 0) return;
    ev.preventDefault();
    ev.stopPropagation();
    finishConnect(ev.clientX, ev.clientY);
  }

  function onArmedKey(ev) {
    if (ev.key === "Escape") {
      ev.preventDefault();
      endConnect();
    }
  }

  // Entry point from the canvas mousedown router (a port dot was pressed).
  function startWireDrag(e, dot) {
    if (!state.ready) return;
    if (_connect) return; // armed — the armed mousedown handler resolves this press
    var src = portInfo(dot);
    if (!src) return;
    e.preventDefault();
    if (!beginConnect(src)) return;
    drawTempWire(e.clientX, e.clientY);

    var startX = e.clientX;
    var startY = e.clientY;
    function move(ev) {
      if (!_connect) return;
      if (
        !_connect.moved &&
        (Math.abs(ev.clientX - startX) > MOVE_THRESHOLD ||
          Math.abs(ev.clientY - startY) > MOVE_THRESHOLD)
      ) {
        _connect.moved = true;
      }
      scheduleTempDraw(ev.clientX, ev.clientY);
    }
    function up(ev) {
      document.removeEventListener("mousemove", move);
      document.removeEventListener("mouseup", up);
      if (!_connect) return;
      if (_connect.moved) {
        finishConnect(ev.clientX, ev.clientY); // a drag: resolve where released
      } else {
        armConnect(); // a click: keep the wire live until the next click
      }
    }
    document.addEventListener("mousemove", move);
    document.addEventListener("mouseup", up);
  }

  // ---- Connect: commit ----

  // Snap to the nearest type-compatible port (of the side we need) on `card`.
  function connectToCard(src, card, clientX, clientY) {
    var needDir = src.dir === "out" ? "in" : "out";
    var dots = card.querySelectorAll('.wf-port-dot[data-port-dir="' + needDir + '"]');
    var best = null;
    var bestDist = Infinity;
    for (var i = 0; i < dots.length; i++) {
      var pt = dots[i].getAttribute("data-port-type");
      var ok = needDir === "in" ? canConnect(src.type, pt) : canConnect(pt, src.type);
      if (!ok) continue;
      var r = dots[i].getBoundingClientRect();
      var dx = r.left + r.width / 2 - clientX;
      var dy = r.top + r.height / 2 - clientY;
      var dist = dx * dx + dy * dy;
      if (dist < bestDist) {
        bestDist = dist;
        best = dots[i];
      }
    }
    if (!best) {
      showToast("No compatible port on that node");
      return;
    }
    connectPorts(src, portInfo(best));
  }

  function connectPorts(src, dst) {
    if (!dst || src.dir === dst.dir) {
      if (dst) showToast("Connect an output to an input");
      return;
    }
    var out = src.dir === "out" ? src : dst;
    var inp = src.dir === "in" ? src : dst;
    if (out.nodeId === inp.nodeId) {
      showToast("Can't wire a node to itself");
      return;
    }
    if (!canConnect(out.type, inp.type)) {
      showToast("Can't connect " + out.type + " → " + inp.type);
      return;
    }
    // An input holds at most one wire — connecting to an occupied input replaces
    // it (this also drops an exact duplicate of the same source).
    var edges = (state.edges || []).filter(function (edge) {
      return !(edge.to === inp.nodeId && edge.toPort === inp.port);
    });
    edges.push({
      id: "e_" + Math.random().toString(36).slice(2, 10),
      from: out.nodeId,
      fromPort: out.port,
      to: inp.nodeId,
      toPort: inp.port,
    });
    state.edges = edges;
    state.selectedEdge = null;
    if (WF.renderAllNodes) WF.renderAllNodes(); // refresh validation + redraw wires
    WF.scheduleSave();
  }

  function isConnecting() {
    return !!(_connect && _connect.mode === "armed");
  }

  // ---- Selection + removal ----

  function selectEdge(id) {
    state.selectedEdge = id;
    if (state.selection.length) state.selection = []; // wire select clears nodes
    if (WF.renderAllNodes) WF.renderAllNodes(); // de-highlight cards + highlight wire
  }

  function removeEdge(id) {
    if (!id) return;
    state.edges = (state.edges || []).filter(function (edge) {
      return edge.id !== id;
    });
    if (state.selectedEdge === id) state.selectedEdge = null;
    if (_hoverEdge === id) _hoverEdge = null;
    if (WF.renderAllNodes) WF.renderAllNodes();
    WF.scheduleSave();
  }

  // ---- Floating × button (hover / selected) ----

  function refreshWireDelete() {
    var btn = qs("#wfWireDelete");
    if (!btn) return;
    var id = _hoverEdge || state.selectedEdge;
    var edge = id ? findEdge(id) : null;
    if (!edge) {
      btn.classList.add("hidden");
      return;
    }
    var a = portWorldPos(edge.from, edge.fromPort, "out");
    var b = portWorldPos(edge.to, edge.toPort, "in");
    if (!a || !b) {
      btn.classList.add("hidden");
      return;
    }
    // World midpoint → screen (inside #wfCanvas): screen = world*zoom + pan.
    var vp = state.viewport;
    btn.style.left = ((a.x + b.x) / 2) * vp.zoom + vp.x + "px";
    btn.style.top = ((a.y + b.y) / 2) * vp.zoom + vp.y + "px";
    btn.setAttribute("data-edge-id", id);
    btn.classList.remove("hidden");
  }

  function initWires() {
    if (_bound) return;
    var svg = qs("#wfWires");
    var btn = qs("#wfWireDelete");
    if (!svg || !btn) return;
    _bound = true;

    svg.addEventListener("pointerover", function (e) {
      var group = e.target && e.target.closest ? e.target.closest(".wf-wire-group") : null;
      if (!group) return;
      if (_hideTimer) {
        clearTimeout(_hideTimer);
        _hideTimer = 0;
      }
      _hoverEdge = group.getAttribute("data-edge-id");
      refreshWireDelete();
    });
    svg.addEventListener("pointerout", function (e) {
      var group = e.target && e.target.closest ? e.target.closest(".wf-wire-group") : null;
      if (!group) return;
      _hoverEdge = null;
      // Grace so the pointer can travel onto the button before it hides.
      if (_hideTimer) clearTimeout(_hideTimer);
      _hideTimer = setTimeout(refreshWireDelete, 120);
    });
    btn.addEventListener("pointerenter", function () {
      if (_hideTimer) {
        clearTimeout(_hideTimer);
        _hideTimer = 0;
      }
    });
    btn.addEventListener("pointerleave", function () {
      if (!state.selectedEdge) {
        _hoverEdge = null;
        refreshWireDelete();
      }
    });
    btn.addEventListener("click", function () {
      removeEdge(btn.getAttribute("data-edge-id"));
    });
  }

  WF.initWires = initWires;
  WF.renderWires = renderWires;
  WF.clearPortCache = clearPortCache;
  WF.startWireDrag = startWireDrag;
  WF.isConnecting = isConnecting;
  // Tear down an in-flight connection (temp wire, highlight, armed document
  // listeners). The hub calls this on blueprint switch / re-layout so an armed
  // wire can't outlive its source node. Safe to call when nothing is connecting.
  WF.cancelConnect = endConnect;
  WF.selectEdge = selectEdge;
  WF.removeEdge = removeEdge;
  WF.canConnect = canConnect;
  WF.refreshWireDelete = refreshWireDelete;
})();
