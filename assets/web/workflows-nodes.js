/* Workflows — node card rendering (satellite of workflows.js).
 *
 * Renders a placed node generically from its catalog NodeType: title, domain
 * accent, ParamSpec-driven param editors, the typed input/output port markers
 * (the wires satellite hooks drag-to-connect onto them), and validation cues
 * (greyed when the launch context is unmet; warned when a required input is
 * unwired). Reads shared state through WF.state — never re-`var`s a divergent
 * `state` (the carve gotcha).
 */

(function () {
  "use strict";

  var WF = window.ClipgenWorkflows;
  var state = WF.state;

  // One column of port rows (inputs on the left, outputs on the right; the
  // outputs column is flipped via CSS so its dot sits on the card edge).
  function buildPortColumn(ports, isOutput) {
    var col = el("div", "wf-port-col " + (isOutput ? "outputs" : "inputs"));
    (ports || []).forEach(function (port) {
      // The universal control input (`__gate__`) reads as a muted "gate" anchor,
      // not a literal port name — a Gate's `pass` output wires here to gate the node.
      var isControl = port.type === "control";
      var row = el(
        "div",
        "wf-port" +
          (port.optional ? " optional" : "") +
          (isControl ? " wf-port-control" : ""),
      );
      var dot = el("span", "wf-port-dot");
      dot.setAttribute("data-port", port.name);
      dot.setAttribute("data-port-type", port.type);
      dot.setAttribute("data-port-dir", isOutput ? "out" : "in");
      row.appendChild(dot);
      row.appendChild(
        el("span", "wf-port-label", isControl ? "gate" : port.name),
      );
      col.appendChild(row);
    });
    return col;
  }

  // Multitool step types: only the per-frame (check_frame) detectors that need
  // no uploaded reference. Each step reuses its ss_<type> catalog params — no
  // duplicated field definitions (the param specs come from the catalog).
  var MT_STEP_TYPES = ["color", "change", "flow", "text", "numbers", "inactivity"];

  function stepParamSpecs(stepType) {
    var nt = state.catalogById && state.catalogById["ss_" + stepType];
    return (nt && nt.params) || [];
  }

  // One ParamSpec editor (number / enum / bool / participant / string / step-list),
  // writing back to `store` (defaults to node.params) on change and autosaving.
  // Scalar editors do NOT re-render on edit, so focus/caret survive typing (the
  // mousedown router also leaves param controls alone for the same reason).
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
      input = el("select", "wf-param-input");
      var participants = (state.context && state.context.participants) || [];
      var options = participants.slice();
      var ALL = WF.ALL_PARTICIPANTS;
      // Keep a stored id that isn't in the discovered list (launched without it),
      // but never the ALL sentinel (it gets its own option below).
      if (value && value !== ALL && options.indexOf(value) < 0)
        options.unshift(value);
      var blank = el("option");
      blank.value = "";
      blank.textContent = "—";
      input.appendChild(blank);
      // "All participants" turns a Run into a whole-study batch (one run each).
      // Only offered when there are participants with video to fan out over.
      if (participants.length) {
        var allOpt = el("option");
        allOpt.value = ALL;
        allOpt.textContent = "All participants";
        if (value === ALL) allOpt.selected = true;
        input.appendChild(allOpt);
      }
      options.forEach(function (pid) {
        var opt = el("option");
        opt.value = pid;
        opt.textContent = pid;
        if (pid === value) opt.selected = true;
        input.appendChild(opt);
      });
      input.addEventListener("change", function () {
        store[spec.name] = input.value;
        WF.scheduleSave();
      });
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

  // Compound editor for the multitool `steps` param: an ordered list of step
  // objects {type, logic, …per-type fields}. Structural changes (add/remove/
  // reorder/type) re-render the list container only; scalar field edits write
  // through buildParamControl(step) without a re-render (focus preserved).
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
        steps.push({ type: MT_STEP_TYPES[0], logic: "AND" });
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
    MT_STEP_TYPES.forEach(function (t) {
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
    var wrap = el("div", "wf-node-params");
    type.params.forEach(function (spec) {
      var row = el("div", "wf-param");
      row.appendChild(el("label", "wf-param-label", spec.label || spec.name));
      row.appendChild(buildParamControl(node, spec));
      wrap.appendChild(row);
    });
    return wrap;
  }

  // True unless some non-optional input port has no incoming edge.
  function requiredInputsSatisfied(node, type) {
    var inputs = type.inputs || [];
    for (var i = 0; i < inputs.length; i++) {
      var port = inputs[i];
      if (port.optional) continue;
      var wired = (state.edges || []).some(function (e) {
        return e.to === node.id && e.toPort === port.name;
      });
      if (!wired) return false;
    }
    return true;
  }

  function renderNode(node) {
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
    card.setAttribute("data-domain", type.domain || "");
    card.style.left = (pos.x || 0) + "px";
    card.style.top = (pos.y || 0) + "px";
    if (state.selection && state.selection.indexOf(node.id) >= 0) {
      card.classList.add("selected");
    }
    // Validation cue: greyed when the launch context can't satisfy `requires`;
    // otherwise warned when a required input is still unwired.
    if (WF.nodeContextMet && !WF.nodeContextMet(type)) {
      card.classList.add("disabled");
    } else if (!requiredInputsSatisfied(node, type)) {
      card.classList.add("invalid");
      card.title = "Connect required inputs";
    }

    card.appendChild(el("div", "wf-node-title", type.label || node.type));
    card.appendChild(el("div", "wf-node-domain", type.domain || ""));

    if (type.params && type.params.length) {
      if (!node.params) node.params = {};
      card.appendChild(buildParamEditors(node, type));
    }

    var hasPorts =
      (type.inputs && type.inputs.length) || (type.outputs && type.outputs.length);
    if (hasPorts) {
      var ports = el("div", "wf-node-ports");
      ports.appendChild(buildPortColumn(type.inputs, false));
      ports.appendChild(buildPortColumn(type.outputs, true));
      card.appendChild(ports);
    }
    return card;
  }

  // Rebuild every card from state.nodes (one DocumentFragment append) and toggle
  // the canvas empty-state. Called on load, drop, delete, and selection change.
  function renderAllNodes() {
    var world = qs("#wfWorld");
    if (!world) return;
    world.innerHTML = "";
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

    // Selection may have changed (drop, marquee, delete) → re-gate "Stash
    // selection". One guarded site keeps the button in sync without touching
    // every gesture that mutates state.selection.
    if (WF.syncStashButton) WF.syncStashButton();
  }

  WF.renderNode = renderNode;
  WF.renderAllNodes = renderAllNodes;
})();
