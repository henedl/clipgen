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

  // The participant menu portaled onto <body>, as {menu, close}; closeParticipantMenu() un-portals it.
  var _openParticipantMenu = null;

  function closeParticipantMenu() {
    if (_openParticipantMenu) _openParticipantMenu.close();
  }

  // One port column; outputs are CSS-flipped. Wired ports get `.wf-port-connected` (filled dot).
  function buildPortColumn(node, ports, isOutput) {
    var col = el("div", "wf-port-col " + (isOutput ? "outputs" : "inputs"));
    var edges = state.edges || [];
    (ports || []).forEach(function (port) {
      // The `__gate__` control input reads as a muted "gate" anchor, not a port name.
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
      // Hover reveals the data type, clarifying adapter-coerced wires.
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

  // Step types come from the catalog's `multitoolStep` flag; _MULTITOOL_STEP_TOOLS is the source.
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

  // Detector keys derive from the hidden ss_<tool> catalog nodes; no JS list.
  function detectTypes() {
    var out = [];
    (state.catalog || []).forEach(function (n) {
      if (n.id && n.id.indexOf("ss_") === 0) out.push(n.id.slice(3));
    });
    return out;
  }

  // Detect node: a detector dropdown plus that detector's ss_<tool> params, swapped in place.
  function buildDetectEditor(node) {
    if (!node.params) node.params = {};
    var types = detectTypes();
    var seeded = false;
    if (!node.params.detector) {
      node.params.detector = types[0] || "text";
      seeded = true;
    }
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
      var specs = stepParamSpecs(node.params.detector);
      specs.forEach(function (ps) {
        // Seed spec defaults so number fields show a value; seeded values must be saved.
        if (node.params[ps.name] === undefined) {
          node.params[ps.name] = ps.default;
          seeded = true;
        }
      });
      buildParamsInto(body, node, specs, node.params);
      if (seeded) {
        seeded = false;
        WF.scheduleSave();
      }
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

  // Datalists for free-text params, appended to <body> and shared by name; "llm-models" fetches ../api/models once.
  var _llmModelsRequested = false;
  function suggestionListId(spec) {
    if (Array.isArray(spec.suggestions) && spec.suggestions.length) {
      var id = "wfDatalist-" + spec.name;
      if (!document.getElementById(id)) {
        var dl = el("datalist");
        dl.id = id;
        spec.suggestions.forEach(function (s) {
          var o = el("option");
          o.value = s;
          dl.appendChild(o);
        });
        document.body.appendChild(dl);
      }
      return id;
    }
    if (spec.datalist === "llm-models") {
      var mid = "wfDatalistLlmModels";
      var mdl = document.getElementById(mid);
      if (!mdl) {
        mdl = el("datalist");
        mdl.id = mid;
        document.body.appendChild(mdl);
      }
      if (!_llmModelsRequested) {
        _llmModelsRequested = true;
        fetch("../api/models")
          .then(function (r) {
            return r.json();
          })
          .then(function (res) {
            var models = (res && res.llm && res.llm.models) || [];
            models.forEach(function (m) {
              var o = el("option");
              o.value = m.name;
              mdl.appendChild(o);
            });
          })
          .catch(function () {});
      }
      return mid;
    }
    return null;
  }

  // One ParamSpec editor writing to `store` and autosaving. Scalar editors never re-render, so focus survives.
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
      if (spec.default != null) input.placeholder = String(spec.default);
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
      // Unset means "server default"; select it so the display matches the run.
      if (value == null && spec.default != null) input.value = spec.default;
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
    } else if (
      spec.type === "region" &&
      ((state.context && state.context.regions) || []).length
    ) {
      // Saved regions become a picker so typos can't full-frame the scan; missing names stay selectable.
      input = el("select", "wf-param-input");
      var regions = state.context.regions;
      var names = [""].concat(regions);
      if (value && names.indexOf(value) < 0) names.push(value);
      names.forEach(function (name) {
        var opt = el("option");
        opt.value = name;
        if (name === "") opt.textContent = "(none)";
        else if (regions.indexOf(name) < 0) opt.textContent = name + " (missing)";
        else opt.textContent = name;
        if (name === value) opt.selected = true;
        input.appendChild(opt);
      });
      input.addEventListener("change", function () {
        store[spec.name] = input.value;
        WF.scheduleSave();
        if (WF.refreshValidation) WF.refreshValidation();
      });
    } else {
      input = el("input", "wf-param-input");
      input.type = "text";
      input.autocomplete = "off";
      if (spec.default != null && spec.default !== "") {
        input.placeholder = String(spec.default);
      }
      var listId = suggestionListId(spec);
      if (listId) input.setAttribute("list", listId);
      input.value = value != null ? value : "";
      input.addEventListener("input", function () {
        store[spec.name] = input.value;
        WF.scheduleSave();
      });
    }
    return input;
  }

  // Param row: label, reset chip (shown when value differs from default), control.
  function buildParamRow(node, spec, store) {
    store = store || node.params;
    var row = el("div", "wf-param");
    var head = el("div", "wf-param-head");
    head.appendChild(el("label", "wf-param-label", spec.label || spec.name));
    var control = buildParamControl(node, spec, store);
    var reset = null;
    var resettable =
      spec.type !== "step-list" &&
      spec.type !== "participant" &&
      spec.default !== undefined;
    function syncReset() {
      if (!reset) return;
      var cur = store[spec.name];
      // An unset value defers to the server default, so it counts as equal.
      var differs =
        cur != null &&
        String(cur) !== String(spec.default == null ? "" : spec.default);
      reset.classList.toggle("hidden", !differs);
    }
    if (resettable) {
      reset = el("button", "wf-param-reset hidden");
      reset.type = "button";
      reset.appendChild(el("span", "wf-btn-icon wf-reset-icon"));
      reset.setAttribute("data-tooltip", "Reset to default");
      reset.setAttribute("aria-label", "Reset to default");
      reset.addEventListener("mousedown", function (e) {
        // Keep the canvas's delegated drag/select handler out of it.
        e.stopPropagation();
      });
      reset.addEventListener("click", function () {
        store[spec.name] = spec.default;
        var next = buildParamControl(node, spec, store);
        row.replaceChild(next, control);
        control = next;
        WF.scheduleSave();
        if (WF.refreshValidation) WF.refreshValidation();
        syncReset();
        // Bubbles to the container so dependent showIf rows re-evaluate.
        row.dispatchEvent(new Event("change", { bubbles: true }));
      });
      head.appendChild(reset);
    }
    row.appendChild(head);
    row.appendChild(control);
    syncReset();
    return {
      row: row,
      spec: spec,
      syncReset: syncReset,
      getControl: function () {
        return control;
      },
    };
  }

  // Build rows for `specs` and keep showIf, numeric-aware values, and reset chips live.
  function buildParamsInto(container, node, specs, store) {
    var entries = specs.map(function (spec) {
      var entry = buildParamRow(node, spec, store);
      container.appendChild(entry.row);
      return entry;
    });
    function bySpecName(name) {
      for (var i = 0; i < entries.length; i++) {
        if (entries[i].spec.name === name) return entries[i];
      }
      return null;
    }
    function sync() {
      entries.forEach(function (en) {
        var spec = en.spec;
        if (spec.showIf && spec.showIf.param) {
          var v = store[spec.showIf.param];
          if (v == null) {
            // Unset → the server will use the controlling param's default.
            var ctrl = bySpecName(spec.showIf.param);
            if (ctrl) v = ctrl.spec.default;
          }
          var show = true;
          if (Object.prototype.hasOwnProperty.call(spec.showIf, "equals")) {
            show = String(v) === String(spec.showIf.equals);
          } else if (Object.prototype.hasOwnProperty.call(spec.showIf, "not")) {
            show = String(v) !== String(spec.showIf.not);
          }
          en.row.classList.toggle("hidden", !show);
        }
        if (spec.numericFor) {
          var fieldEntry = bySpecName(spec.numericFor);
          var fieldVal = store[spec.numericFor];
          if (fieldVal == null && fieldEntry) fieldVal = fieldEntry.spec.default;
          var numeric =
            fieldEntry &&
            (fieldEntry.spec.numericChoices || []).indexOf(fieldVal) >= 0;
          var input = en.getControl();
          if (input && input.tagName === "INPUT") {
            input.type = numeric ? "number" : "text";
          }
        }
        en.syncReset();
      });
    }
    container.addEventListener("input", sync);
    container.addEventListener("change", sync);
    sync();
  }

  // Participant picker. A single pick persists as a string, never a 1-element array.
  function buildParticipantSelect(spec, store) {
    var ALL = WF.ALL_PARTICIPANTS;
    var participants = (state.context && state.context.participants) || [];
    var current = store ? store[spec.name] : spec.default;
    var isAll = current === ALL;

    // Discovered ids plus any stored id not discovered, so saved selections round-trip.
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
    var menu = el("div", "wf-participant-menu cg-menu hidden");
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

    // "All participants" is offered only when there is something to fan out over.
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
        // All boxes checked collapses to ALL; otherwise an explicit subset (normalized in persist()).
        isAll =
          ids.length > 0 &&
          ids.every(function (x) {
            return rowCbs[x].checked;
          });
        if (allCb) allCb.checked = isAll;
        persist();
      });
    });

    // Portaled onto <body>: the canvas clips it and #wfWorld's transform defeats position:fixed. See transcripts-pills.js.
    if (WF.bindMenuToggle) {
      // `toggle` is assigned before any click can fire, so onOpen can close over it.
      var toggle = WF.bindMenuToggle(btn, menu, {
        onOpen: function () {
          closeParticipantMenu(); // only one open at a time
          document.body.appendChild(menu);
          positionPopoverAnchored(menu, btn.getBoundingClientRect());
          _openParticipantMenu = { menu: menu, close: toggle.close };
        },
        onClose: function () {
          // Back into the card, or drop it when the card is already gone.
          if (wrap.isConnected) wrap.appendChild(menu);
          else if (menu.parentNode) menu.parentNode.removeChild(menu);
          if (_openParticipantMenu && _openParticipantMenu.menu === menu) {
            _openParticipantMenu = null;
          }
        },
      });
    }
    refreshSummary();
    return wrap;
  }

  // Multitool `steps` editor. Structural changes re-render the list; scalar edits don't, preserving focus.
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
    buildParamsInto(body, node, stepParamSpecs(step.type), step);
    card.appendChild(body);
    return card;
  }

  function buildParamEditors(node, type) {
    if (node.type === "detect") return buildDetectEditor(node);
    if (!node.params) node.params = {};
    var wrap = el("div", "wf-node-params");
    buildParamsInto(wrap, node, type.params, node.params);
    return wrap;
  }

  // Sticky note: keeps .wf-node + data-node-id so canvas gestures work; runner drops type "note".
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
    // Header is the grab surface; the textarea is exempt from canvas drag.
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

  // Inline rename: commit on blur/Enter, Escape restores, empty clears the rename.
  function startRenameNode(node, titleText) {
    var type = state.catalogById[node.type] || {};
    var input = el("input", "wf-node-rename");
    input.type = "text";
    input.autocomplete = "off";
    input.value = node.name || "";
    input.placeholder = type.label || node.type;
    input.addEventListener("mousedown", function (e) {
      e.stopPropagation(); // keep the canvas drag handler out of it
    });
    input.addEventListener("keydown", function (e) {
      if (e.key === "Enter") input.blur();
      else if (e.key === "Escape") {
        input.value = node.name || "";
        input.blur();
      }
      e.stopPropagation();
    });
    input.addEventListener("blur", function () {
      var name = input.value.trim();
      if (name) node.name = name;
      else delete node.name;
      WF.scheduleSave();
      if (WF.renderAllNodes) WF.renderAllNodes();
      if (WF.refreshValidation) WF.refreshValidation();
    });
    titleText.textContent = "";
    titleText.appendChild(input);
    input.focus();
    input.select();
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
    // Extra width for Detect, step-list params, or more than three params.
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
    // Greyed when context is unmet, dashed `.invalid` on errors; warnings are tooltip-only.
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

    // Title bar: label plus a `?` help glyph using the [data-tooltip] singleton, not title.
    var titleBar = el("div", "wf-node-title");
    var titleText = el(
      "span",
      "wf-node-title-text",
      node.name || type.label || node.type,
    );
    // Double-click renames; the custom name disambiguates duplicate types. Tooltip keeps the type reachable.
    if (node.name) titleText.setAttribute("data-tooltip", type.label || node.type);
    titleText.addEventListener("dblclick", function (e) {
      e.stopPropagation();
      startRenameNode(node, titleText);
    });
    titleBar.appendChild(titleText);
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

  // Rebuild every card from state.nodes and toggle the canvas empty-state.
  function renderAllNodes() {
    return clipgenPerf.span("workflows.renderAllNodes", renderAllNodesImpl);
  }

  function renderAllNodesImpl() {
    var world = qs("#wfWorld");
    if (!world) return;
    // The anchor card is going away; close the <body>-portaled menu first.
    closeParticipantMenu();
    // Keep the wire <svg> (shares the cards' stacking context); renderWires() refills it.
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

    // Selection may have changed; re-gate the selection-dependent toolbar buttons here.
    if (WF.syncStashButton) WF.syncStashButton();
    if (WF.syncRunButton) WF.syncRunButton();
    // Card rebuild dropped the last-run badges — re-apply (runs satellite).
    if (WF.applyLastRunBadges) WF.applyLastRunBadges();

    // Node set changed; refresh the minimap. Pan/zoom/drag have their own hooks.
    if (WF.renderMinimap) WF.renderMinimap();
  }

  window.addEventListener("pagehide", closeParticipantMenu);

  WF.renderNode = renderNode;
  WF.renderAllNodes = renderAllNodes;
  WF.closeParticipantMenu = closeParticipantMenu;
})();
