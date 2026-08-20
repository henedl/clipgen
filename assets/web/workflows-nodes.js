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
        // Seed the spec default so number fields show a value; anything seeded
        // must also be saved, or the values exist only until the next reload.
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

  // Completion sources for free-text params. Static `suggestions` ride on the
  // catalog spec; the "ollama-models" source is fetched once from the combined
  // app's /api/models (page-relative ../api/models) and fills in when it lands.
  // Datalists are appended to <body> and shared by name across all cards.
  var _ollamaModelsRequested = false;
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
    if (spec.datalist === "ollama-models") {
      var mid = "wfDatalistOllamaModels";
      var mdl = document.getElementById(mid);
      if (!mdl) {
        mdl = el("datalist");
        mdl.id = mid;
        document.body.appendChild(mdl);
      }
      if (!_ollamaModelsRequested) {
        _ollamaModelsRequested = true;
        fetch("../api/models")
          .then(function (r) {
            return r.json();
          })
          .then(function (res) {
            var models = (res && res.ollama && res.ollama.models) || [];
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

  // One ParamSpec editor (number / enum / bool / participant / region / string /
  // step-list), writing back to `store` (default node.params) on change and
  // autosaving. Scalar editors do NOT re-render on edit, so focus/caret survive
  // typing — the mousedown router leaves param controls alone for the same reason.
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
      // An unset value means "server default"; select it rather than letting
      // the browser display option[0] while the run uses something else.
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
      // Saved Screenspace regions exist — offer them as a picker so a typo
      // can't silently full-frame the scan. A stored name that no longer
      // exists stays selectable, flagged "(missing)", so a saved blueprint
      // round-trips; with no saved regions the generic text branch below
      // applies (the server matches by name either way).
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

  // One param row: label + reset-to-default chip + control. The chip shows only
  // when the stored value differs from the spec default; resetting rebuilds the
  // control in place so the widget logic stays in buildParamControl.
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

  // Build every row for `specs` into `container` and keep the conditional bits
  // live: showIf visibility, numeric-aware free-text values (filter/partition
  // `value` follows the picked field's numericChoices), and the reset chips.
  // One delegated listener pair per container — scalar editors still never
  // re-render, so focus/caret survive typing.
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

  // Swap a card's title text for an inline rename input. Commit on blur/Enter;
  // Escape restores the previous name; an empty name clears the rename.
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
    var titleText = el(
      "span",
      "wf-node-title-text",
      node.name || type.label || node.type,
    );
    // Double-click renames the node (blank restores the catalog label) — the
    // custom name is what disambiguates duplicate types on the canvas and in
    // the run panel's rows. A renamed card keeps its type reachable via tooltip.
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

  // Rebuild every card from state.nodes (one DocumentFragment append) and toggle
  // the canvas empty-state. Called on load, drop, delete, and selection change.
  function renderAllNodes() {
    return clipgenPerf.span("workflows.renderAllNodes", renderAllNodesImpl);
  }

  function renderAllNodesImpl() {
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
    // Card rebuild dropped the last-run badges — re-apply (runs satellite).
    if (WF.applyLastRunBadges) WF.applyLastRunBadges();

    // Node set changed (add/delete/blueprint-load) → refresh the minimap. Pan/
    // zoom and drag are covered by their own hooks in the canvas satellite.
    if (WF.renderMinimap) WF.renderMinimap();
  }

  WF.renderNode = renderNode;
  WF.renderAllNodes = renderAllNodes;
})();
