/* Workflows — node card rendering (satellite of workflows.js).
 *
 * Renders a placed node generically from its catalog NodeType: title, domain
 * accent, ParamSpec-driven param editors, the typed port markers the wires
 * satellite hooks drag-to-connect onto, and validation cues (greyed when the
 * launch context is unmet, warned when a required input is unwired).
 */

(function () {
  "use strict";

  var WF = window.ClipgenWorkflows;
  var state = WF.state;

  // One column of port rows (inputs left, outputs right — the outputs column is
  // CSS-flipped so its dot sits on the card edge). Ports already wired per
  // `state.edges` get `.wf-port-connected`, which fills the dot rather than
  // leaving it a hollow ring.
  function buildPortColumn(node, ports, isOutput) {
    var col = el("div", "wf-port-col " + (isOutput ? "outputs" : "inputs"));
    var edges = state.edges || [];
    (ports || []).forEach(function (port) {
      // The universal control input (`__gate__`) reads as a muted "gate" anchor,
      // not a literal port name — a Gate's `pass` output wires here to gate the node.
      var isControl = port.type === "control";
      var connected = edges.some(function (e) {
        return isOutput
          ? e.from === node.id && e.fromPort === port.name
          : e.to === node.id && e.toPort === port.name;
      });
      var row = el(
        "div",
        "wf-port" +
          (port.optional ? " optional" : "") +
          (isControl ? " wf-port-control" : "") +
          (connected ? " wf-port-connected" : ""),
      );
      var dot = el("span", "wf-port-dot");
      dot.setAttribute("data-port", port.name);
      dot.setAttribute("data-port-type", port.type);
      dot.setAttribute("data-port-dir", isOutput ? "out" : "in");
      // Hovering a port reveals its data type — clarifies adapter-coerced wires
      // (e.g. a `timeRange` output into a `clips`/clipRecords input).
      dot.title = isControl ? "gate" : port.type;
      // Assistive tech: the dot is an interactive connection point.
      dot.setAttribute("role", "button");
      dot.setAttribute(
        "aria-label",
        isControl
          ? "Gate control input"
          : (isOutput ? "Output" : "Input") + ": " + port.name + " (" + port.type + ")",
      );
      row.appendChild(dot);
      row.appendChild(
        el("span", "wf-port-label", isControl ? "gate" : port.name),
      );
      col.appendChild(row);
    });
    return col;
  }

  // Multitool step types: the per-frame (check_frame) detectors needing no
  // uploaded reference, derived from the catalog's `multitoolStep` flag — the
  // backend's _MULTITOOL_STEP_TOOLS is the single source, no hardcoded JS list.
  // Each step reuses its ss_<type> catalog params.
  function multitoolStepTypes() {
    var out = [];
    (state.catalog || []).forEach(function (n) {
      if (n.multitoolStep && n.id && n.id.indexOf("ss_") === 0) {
        out.push(n.id.slice(3));
      }
    });
    return out;
  }

  function stepParamSpecs(stepType) {
    var nt = state.catalogById && state.catalogById["ss_" + stepType];
    return (nt && nt.params) || [];
  }

  // Detector keys for the unified Detect node, derived from the (hidden) ss_<tool>
  // catalog nodes — no duplicated list in JS.
  function detectTypes() {
    var out = [];
    (state.catalog || []).forEach(function (n) {
      if (n.id && n.id.indexOf("ss_") === 0) out.push(n.id.slice(3));
    });
    return out;
  }

  // The unified Detect node: a detector dropdown plus that detector's ss_<tool>
  // param set, swapped in place on change. The Multitool step editor generalised
  // to a single node-level step.
  function buildDetectEditor(node) {
    if (!node.params) node.params = {};
    var types = detectTypes();
    if (!node.params.detector) node.params.detector = types[0] || "text";
    var wrap = el("div", "wf-node-params");

    var row = el("div", "wf-param");
    row.appendChild(el("label", "wf-param-label", "Detector"));
    var sel = el("select", "wf-param-input");
    types.forEach(function (t) {
      var o = el("option");
      o.value = t;
      o.textContent = t;
      if (t === node.params.detector) o.selected = true;
      sel.appendChild(o);
    });
    var body = el("div", "wf-detect-body");
    function renderBody() {
      body.innerHTML = "";
      stepParamSpecs(node.params.detector).forEach(function (ps) {
        // Seed the spec default so number fields show a value (and persist on the
        // next save); the server also defaults missing params defensively.
        if (node.params[ps.name] === undefined) node.params[ps.name] = ps.default;
        var prow = el("div", "wf-param");
        prow.appendChild(el("label", "wf-param-label", ps.label || ps.name));
        prow.appendChild(buildParamControl(node, ps));
        body.appendChild(prow);
      });
    }
    sel.addEventListener("change", function () {
      node.params.detector = sel.value;
      WF.scheduleSave();
      renderBody(); // the param set differs per detector
      if (WF.refreshValidation) WF.refreshValidation();
    });
    row.appendChild(sel);
    wrap.appendChild(row);
    wrap.appendChild(body);
    renderBody();
    return wrap;
  }

  // One ParamSpec editor (number / enum / bool / participant / string / step-list),
  // writing back to `store` (default node.params) on change and autosaving. Scalar
  // editors do NOT re-render on edit, so focus/caret survive typing — the mousedown
  // router leaves param controls alone for the same reason.
  function buildParamControl(node, spec, store) {
    if (spec.type === "step-list") return buildStepList(node, spec);
    store = store || node.params;
    var value = store ? store[spec.name] : spec.default;
    var input;
    if (spec.type === "number") {
      input = el("input", "wf-param-input");
      input.type = "number";
      if (spec.min !== undefined) input.min = spec.min;
      if (spec.max !== undefined) input.max = spec.max;
      input.value = value != null ? value : "";
      input.addEventListener("input", function () {
        var n = parseFloat(input.value);
        store[spec.name] = isNaN(n) ? null : n;
        WF.scheduleSave();
      });
    } else if (spec.type === "enum") {
      input = el("select", "wf-param-input");
      (spec.choices || []).forEach(function (choice) {
        var opt = el("option");
        opt.value = choice;
        opt.textContent = choice;
        if (choice === value) opt.selected = true;
        input.appendChild(opt);
      });
      input.addEventListener("change", function () {
        store[spec.name] = input.value;
        WF.scheduleSave();
      });
    } else if (spec.type === "bool") {
      input = el("input", "wf-param-input");
      input.type = "checkbox";
      input.checked = !!value;
      input.addEventListener("change", function () {
        store[spec.name] = input.checked;
        WF.scheduleSave();
      });
    } else if (spec.type === "participant") {
      input = buildParticipantSelect(spec, store);
    } else {
      input = el("input", "wf-param-input");
      input.type = "text";
      input.autocomplete = "off";
      input.value = value != null ? value : "";
      input.addEventListener("input", function () {
        store[spec.name] = input.value;
        WF.scheduleSave();
      });
    }
    return input;
  }

  // Multi-select participant picker: a summary button opening a checkbox popover.
  // Writes a normalized value back to store[spec.name]:
  //   • a single id string  → single run (server's scalar path),
  //   • the ALL sentinel     → batch over every participant,
  //   • an array of ≥2 ids   → batch over that subset,
  //   • an empty array       → nothing selected (flagged by validation).
  // Normalizing a single pick to a string, never a 1-element array, keeps the
  // server's single-run path untouched. Like the scalar editors it doesn't
  // re-render the card on change, so focus and the open popover survive.
  function buildParticipantSelect(spec, store) {
    var ALL = WF.ALL_PARTICIPANTS;
    var participants = (state.context && state.context.participants) || [];
    var current = store ? store[spec.name] : spec.default;
    var isAll = current === ALL;

    // Discovered ids, plus any stored id not currently discovered (launched
    // without it) so a saved selection round-trips.
    var ids = participants.slice();
    var initSel = {};
    if (Array.isArray(current)) {
      current.forEach(function (id) {
        initSel[id] = true;
      });
    } else if (current && current !== ALL) {
      initSel[current] = true;
    }
    Object.keys(initSel).forEach(function (id) {
      if (ids.indexOf(id) < 0) ids.push(id);
    });

    var wrap = el("div", "wf-participant-select");
    var btn = el("button", "wf-participant-btn");
    btn.type = "button";
    btn.setAttribute("aria-haspopup", "menu");
    btn.setAttribute("aria-expanded", "false");
    var menu = el("div", "wf-participant-menu hidden");
    menu.setAttribute("role", "menu");
    wrap.appendChild(btn);
    wrap.appendChild(menu);

    var allCb = null;
    var rowCbs = {};

    function pickedIds() {
      return ids.filter(function (id) {
        return rowCbs[id] && rowCbs[id].checked;
      });
    }
    function refreshSummary() {
      var picked = pickedIds();
      var txt;
      if (isAll) txt = "All participants";
      else if (!picked.length) txt = "Select…";
      else if (picked.length === 1) txt = picked[0];
      else txt = picked.length + " participants";
      btn.textContent = txt;
    }
    function persist() {
      var picked = pickedIds();
      var out;
      if (isAll) out = ALL;
      else if (!picked.length) out = [];
      else if (picked.length === 1) out = picked[0];
      else out = picked;
      if (store) store[spec.name] = out;
      refreshSummary();
      WF.scheduleSave();
    }

    // "All participants" shortcut — only offered when there are participants to
    // fan out over (matches the old select's gating).
    if (participants.length) {
      var allRow = el("label", "wf-participant-opt wf-participant-all");
      allCb = el("input");
      allCb.type = "checkbox";
      allCb.checked = isAll;
      allRow.appendChild(allCb);
      allRow.appendChild(el("span", null, "All participants"));
      menu.appendChild(allRow);
      allCb.addEventListener("change", function () {
        isAll = allCb.checked;
        ids.forEach(function (id) {
          if (rowCbs[id]) rowCbs[id].checked = isAll;
        });
        persist();
      });
    }

    ids.forEach(function (id) {
      var row = el("label", "wf-participant-opt");
      var cb = el("input");
      cb.type = "checkbox";
      cb.checked = isAll || !!initSel[id];
      rowCbs[id] = cb;
      row.appendChild(cb);
      row.appendChild(el("span", null, id));
      menu.appendChild(row);
      cb.addEventListener("change", function () {
        // Every box checked collapses to the ALL sentinel; otherwise it's an
        // explicit subset (or a single id, normalized in persist()).
        isAll =
          ids.length > 0 &&
          ids.every(function (x) {
            return rowCbs[x].checked;
          });
        if (allCb) allCb.checked = isAll;
        persist();
      });
    });

    if (WF.bindMenuToggle) WF.bindMenuToggle(btn, menu);
    refreshSummary();
    return wrap;
  }

  // Compound editor for the multitool `steps` param: an ordered list of
  // {type, logic, …per-type fields}. Structural changes (add/remove/reorder/type)
  // re-render the list container only; scalar edits write through
  // buildParamControl(step) with no re-render, preserving focus.
  function buildStepList(node, spec) {
    if (!Array.isArray(node.params[spec.name])) node.params[spec.name] = [];
    var steps = node.params[spec.name];
    var container = el("div", "wf-step-list");

    function rerender() {
      container.innerHTML = "";
      steps.forEach(function (step, idx) {
        container.appendChild(buildStepCard(node, spec, step, idx, rerender));
      });
      var add = el("button", "wf-step-add", "+ Add step");
      add.type = "button";
      add.addEventListener("click", function () {
        steps.push({ type: multitoolStepTypes()[0] || "color", logic: "AND" });
        WF.scheduleSave();
        rerender();
      });
      container.appendChild(add);
    }
    rerender();
    return container;
  }

  function buildStepCard(node, spec, step, idx, rerender) {
    var steps = node.params[spec.name];
    var card = el("div", "wf-step");
    var head = el("div", "wf-step-head");

    var typeSel = el("select", "wf-param-input");
    multitoolStepTypes().forEach(function (t) {
      var o = el("option");
      o.value = t;
      o.textContent = t;
      if (t === step.type) o.selected = true;
      typeSel.appendChild(o);
    });
    typeSel.addEventListener("change", function () {
      step.type = typeSel.value;
      WF.scheduleSave();
      rerender(); // body fields differ per type
    });
    head.appendChild(typeSel);

    // Steps after the first carry a chain logic (AND / NOT).
    if (idx > 0) {
      var logicSel = el("select", "wf-param-input wf-step-logic");
      ["AND", "NOT"].forEach(function (l) {
        var o = el("option");
        o.value = l;
        o.textContent = l;
        if (l === (step.logic || "AND")) o.selected = true;
        logicSel.appendChild(o);
      });
      logicSel.addEventListener("change", function () {
        step.logic = logicSel.value;
        WF.scheduleSave();
      });
      head.appendChild(logicSel);
    }

    function moveBtn(label, delta, disabled) {
      var b = el("button", "wf-step-btn", label);
      b.type = "button";
      b.disabled = disabled;
      b.addEventListener("click", function () {
        var j = idx + delta;
        if (j < 0 || j >= steps.length) return;
        steps.splice(j, 0, steps.splice(idx, 1)[0]);
        WF.scheduleSave();
        rerender();
      });
      return b;
    }
    head.appendChild(moveBtn("↑", -1, idx === 0));
    head.appendChild(moveBtn("↓", 1, idx === steps.length - 1));
    var rm = el("button", "wf-step-btn", "✕");
    rm.type = "button";
    rm.addEventListener("click", function () {
      steps.splice(idx, 1);
      WF.scheduleSave();
      rerender();
    });
    head.appendChild(rm);
    card.appendChild(head);

    var body = el("div", "wf-step-body");
    stepParamSpecs(step.type).forEach(function (ps) {
      var row = el("div", "wf-param");
      row.appendChild(el("label", "wf-param-label", ps.label || ps.name));
      row.appendChild(buildParamControl(node, ps, step));
      body.appendChild(row);
    });
    card.appendChild(body);
    return card;
  }

  function buildParamEditors(node, type) {
    if (node.type === "detect") return buildDetectEditor(node);
    var wrap = el("div", "wf-node-params");
    type.params.forEach(function (spec) {
      var row = el("div", "wf-param");
      row.appendChild(el("label", "wf-param-label", spec.label || spec.name));
      row.appendChild(buildParamControl(node, spec));
      wrap.appendChild(row);
    });
    return wrap;
  }

  // Sticky-note pseudo-node: a canvas annotation, not an executable card. It
  // keeps the .wf-node class + data-node-id so drag/marquee/delete/copy/minimap
  // all work untouched; the runner filters type "note" out server-side.
  function renderNoteCard(node) {
    var pos = node.position || { x: 0, y: 0 };
    var card = el("div", "wf-node wf-note");
    card.setAttribute("data-node-id", node.id);
    card.setAttribute("data-node-type", "note");
    card.style.left = (pos.x || 0) + "px";
    card.style.top = (pos.y || 0) + "px";
    if (state.selection && state.selection.indexOf(node.id) >= 0) {
      card.classList.add("selected");
    }
    // Slim header as the labeled grab surface (the textarea itself is exempt
    // from canvas drag via the param-control rule in onCanvasMouseDown).
    card.appendChild(el("div", "wf-note-header", "Note"));
    var ta = document.createElement("textarea");
    ta.className = "wf-note-text";
    ta.placeholder = "Write a note…";
    ta.value = (node.params && node.params.text) || "";
    ta.addEventListener("input", function () {
      if (!node.params) node.params = {};
      node.params.text = ta.value;
      WF.scheduleSave(); // no re-render — typing must not rebuild the card
    });
    card.appendChild(ta);
    return card;
  }

  function renderNode(node) {
    if (node.type === "note") return renderNoteCard(node);
    var type = state.catalogById[node.type] || {
      label: node.type,
      domain: "",
      inputs: [],
      outputs: [],
      params: [],
    };
    var pos = node.position || { x: 0, y: 0 };

    var card = el("div", "wf-node");
    card.setAttribute("data-node-id", node.id);
    card.setAttribute("data-node-type", node.type);
    card.setAttribute("data-domain", type.domain || "");
    // Busy nodes get extra width so their controls stay readable: the Detect
    // node's swappable param set, a compound step-list param (Multitool), or any
    // node carrying more than three params (e.g. Make Clips with titlecard knobs).
    if (
      node.type === "detect" ||
      (type.params || []).length > 3 ||
      (type.params || []).some(function (p) {
        return p.type === "step-list";
      })
    ) {
      card.classList.add("wf-node-wide");
    }
    card.style.left = (pos.x || 0) + "px";
    card.style.top = (pos.y || 0) + "px";
    if (state.selection && state.selection.indexOf(node.id) >= 0) {
      card.classList.add("selected");
    }
    // Muted nodes are dimmed; the runner skips them and their downstream subtree.
    if (node.disabled) card.classList.add("wf-node-muted");
    // Validation cue (shares WF.nodeIssues with the Issues panel): greyed when
    // the launch context can't satisfy `requires`; otherwise a dashed `.invalid`
    // border for any remaining error (unwired required input / empty required
    // param). Warnings surface only as a tooltip, never a blocking cue.
    if (WF.nodeContextMet && !WF.nodeContextMet(type)) {
      card.classList.add("disabled");
      card.title = "Requires " + ((type.requires || []).join(", ") || "context");
    } else {
      var issues = WF.nodeIssues
        ? WF.nodeIssues(node)
        : { errors: [], warnings: [] };
      if (issues.errors.length) {
        card.classList.add("invalid");
        card.title = issues.errors.join("; ");
      } else if (issues.warnings.length) {
        card.title = issues.warnings.join("; ");
      }
    }

    // Colour-coded title bar (domain background via CSS data-domain): the label
    // plus a `?` help glyph whose tooltip carries the catalog description. Uses
    // the [data-tooltip] singleton (styled/in-viewport), not native title.
    var titleBar = el("div", "wf-node-title");
    titleBar.appendChild(el("span", "wf-node-title-text", type.label || node.type));
    if (type.description) {
      var help = el("span", "wf-node-help");
      help.setAttribute("data-tooltip", type.description);
      titleBar.appendChild(help);
    }
    // Mute toggle: skip this node (and its downstream subtree) without deleting.
    var mute = el("span", "wf-node-mute");
    if (node.disabled) mute.classList.add("on");
    mute.setAttribute(
      "data-tooltip",
      node.disabled ? "Un-mute (run this node)" : "Mute (skip this node + downstream)",
    );
    mute.setAttribute("role", "button");
    // Stop the mousedown reaching the canvas's delegated drag/select handler.
    mute.addEventListener("mousedown", function (e) {
      e.stopPropagation();
    });
    mute.addEventListener("click", function (e) {
      e.stopPropagation();
      node.disabled = !node.disabled;
      if (WF.renderAllNodes) WF.renderAllNodes();
      if (WF.refreshValidation) WF.refreshValidation();
      WF.scheduleSave();
    });
    titleBar.appendChild(mute);
    card.appendChild(titleBar);
    card.appendChild(el("div", "wf-node-domain", type.domain || ""));

    if (type.params && type.params.length) {
      if (!node.params) node.params = {};
      card.appendChild(buildParamEditors(node, type));
    }

    var hasPorts =
      (type.inputs && type.inputs.length) || (type.outputs && type.outputs.length);
    if (hasPorts) {
      var ports = el("div", "wf-node-ports");
      ports.appendChild(buildPortColumn(node, type.inputs, false));
      ports.appendChild(buildPortColumn(node, type.outputs, true));
      card.appendChild(ports);
    }
    return card;
  }

  // Rebuild every card from state.nodes (one DocumentFragment append) and toggle
  // the canvas empty-state. Called on load, drop, delete, and selection change.
  function renderAllNodes() {
    var world = qs("#wfWorld");
    if (!world) return;
    // Clear the cards but keep the nested wire <svg> (it lives in #wfWorld so it
    // shares the cards' stacking context — see workflows.html). renderWires()
    // below repopulates its paths.
    var wires = world.querySelector("#wfWires");
    world.innerHTML = "";
    if (wires) world.appendChild(wires);
    var frag = document.createDocumentFragment();
    (state.nodes || []).forEach(function (node) {
      frag.appendChild(renderNode(node));
    });
    world.appendChild(frag);

    var empty = qs("#wfCanvasEmpty");
    if (empty) empty.classList.toggle("hidden", (state.nodes || []).length > 0);

    // Port DOM was rebuilt → drop the wires' cached port offsets, then redraw.
    if (WF.clearPortCache) WF.clearPortCache();
    if (WF.renderWires) WF.renderWires();

    // Selection may have changed (drop, marquee, delete) → re-gate the
    // selection-dependent toolbar buttons ("Stash selection", "Run to here").
    // One guarded site keeps them in sync without touching every gesture that
    // mutates state.selection.
    if (WF.syncStashButton) WF.syncStashButton();
    if (WF.syncRunButton) WF.syncRunButton();

    // Node set changed (add/delete/blueprint-load) → refresh the minimap. Pan/
    // zoom and drag are covered by their own hooks in the canvas satellite.
    if (WF.renderMinimap) WF.renderMinimap();
  }

  WF.renderNode = renderNode;
  WF.renderAllNodes = renderAllNodes;
})();
