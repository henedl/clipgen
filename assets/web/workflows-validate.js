/* Workflows — pre-run validation panel (satellite of workflows.js).
 *
 * Aggregates the graph's pre-run issues into #wfValidation and gates the Run
 * button on errors (warnings never block). Recomputed on every edit via
 * WF.refreshValidation, which the hub calls from scheduleSave + openBlueprint.
 * Owns the per-node issue computation (WF.nodeIssues), shared with the node-card
 * cue in workflows-nodes.js, plus a JS port of workflows.topo_order's cycle check
 * (the server 400 stays a backstop).
 */

(function () {
  "use strict";

  var WF = window.ClipgenWorkflows;
  var state = WF.state;

  function catalogType(node) {
    return (
      state.catalogById[node.type] || {
        inputs: [],
        outputs: [],
        params: [],
        requires: [],
      }
    );
  }

  function inputWired(nodeId, portName) {
    return (state.edges || []).some(function (e) {
      return e.to === nodeId && e.toPort === portName;
    });
  }

  // The node type wired into `nodeId`'s `portName` input, or null.
  function upstreamType(nodeId, portName) {
    var edges = state.edges || [];
    for (var i = 0; i < edges.length; i++) {
      var e = edges[i];
      if (e.to === nodeId && e.toPort === portName) {
        var src = WF.findNode ? WF.findNode(e.from) : null;
        return src ? src.type : null;
      }
    }
    return null;
  }

  // True unless some non-optional, non-control input port has no incoming edge.
  function requiredInputsSatisfied(node, type) {
    var inputs = type.inputs || [];
    for (var i = 0; i < inputs.length; i++) {
      var port = inputs[i];
      if (port.optional || port.type === "control") continue;
      if (!inputWired(node.id, port.name)) return false;
    }
    return true;
  }

  function paramEmpty(value) {
    return value === undefined || value === null || value === "";
  }

  // Per-node {errors, warnings}; shared by the node-card cue and panel rows.
  function nodeIssues(node) {
    // Muted nodes and sticky notes never run, so they report no issues.
    if (node.disabled || node.type === "note") {
      return { errors: [], warnings: [] };
    }
    var type = catalogType(node);
    var errors = [];
    var warnings = [];

    // A type missing from the catalog can never execute; catalogType's fallback hides that.
    if (!state.catalogById[node.type]) {
      errors.push("Unknown node type “" + node.type + "”");
    }
    // error — multitool needs two steps to chain; fewer yields empty events.
    if (node.type === "multitool") {
      var mtSteps = (node.params || {}).steps;
      if (!Array.isArray(mtSteps) || mtSteps.length < 2) {
        errors.push("Add at least 2 steps");
      }
    }
    // error — launch context can't satisfy `requires` (sheet/videoDir).
    if (WF.nodeContextMet && !WF.nodeContextMet(type)) {
      errors.push("Requires " + ((type.requires || []).join(", ") || "context"));
    }
    // error — a required input port is unwired.
    if (!requiredInputsSatisfied(node, type)) {
      errors.push("Connect required inputs");
    }
    // error — a required param is empty.
    (type.params || []).forEach(function (spec) {
      if (spec.required && paramEmpty((node.params || {})[spec.name])) {
        errors.push("Set “" + (spec.label || spec.name) + "”");
      }
    });
    // error — the Detect node's active detector params live on the hidden ss_<detector> spec.
    if (node.type === "detect") {
      var det = (node.params || {}).detector;
      var specNode = det && state.catalogById && state.catalogById["ss_" + det];
      ((specNode && specNode.params) || []).forEach(function (spec) {
        if (spec.required && paramEmpty((node.params || {})[spec.name])) {
          errors.push("Set “" + (spec.label || spec.name) + "”");
        }
      });
    }

    // warning — heatmap style needs a matching upstream detector.
    if (node.type === "heatmap") {
      // Style names map 1:1 to the ss_<style> detector producing the raw_results.
      var want = "ss_" + ((node.params || {}).style || "change");
      var up = upstreamType(node.id, "events");
      if (up && up !== want && up !== "multitool") {
        warnings.push("Heatmap style needs a " + want + " (or multitool) upstream");
      }
    }
    // warning — a gate with no scalar source can't evaluate.
    if (node.type === "gate" && !inputWired(node.id, "value")) {
      warnings.push("Gate has no scalar source");
    }
    // warning — [] means every participant box is unchecked; the run yields nothing.
    if (
      node.type === "video_source" &&
      Array.isArray((node.params || {}).participant) &&
      node.params.participant.length === 0
    ) {
      warnings.push("No participants selected");
    }
    // error — predicates that drop every item: ordering on text, non-numbers vs numeric fields.
    if (node.type.indexOf("filter_") === 0) {
      var fpSpecs = type.params || [];
      var fpParams = node.params || {};
      var specByName = function (n) {
        for (var si = 0; si < fpSpecs.length; si++) {
          if (fpSpecs[si].name === n) return fpSpecs[si];
        }
        return null;
      };
      var clauseError = function (fieldKey, opKey, valueKey, suffix) {
        var fs = specByName(fieldKey);
        var field =
          fpParams[fieldKey] != null ? fpParams[fieldKey] : fs && fs.default;
        var numeric = fs && (fs.numericChoices || []).indexOf(field) >= 0;
        var op = fpParams[opKey] != null ? fpParams[opKey] : ">=";
        var val = fpParams[valueKey];
        var ordering = op === ">=" || op === ">" || op === "<=" || op === "<";
        var isNum = /^\s*-?(\d+\.?\d*|\.\d+)\s*$/.test(String(val));
        if (ordering && !numeric) {
          errors.push("Ordering comparison" + suffix + " needs a numeric field");
        } else if (numeric && op !== "contains" && !paramEmpty(val) && !isNum) {
          errors.push("Value" + suffix + " must be a number");
        }
      };
      clauseError("field", "op", "value", "");
      if (fpParams.combine && fpParams.combine !== "off") {
        clauseError("field2", "op2", "value2", " 2");
        if (paramEmpty(fpParams.value2)) errors.push("Set “Value 2”");
      }
    }
    var connected = (state.edges || []).some(function (e) {
      return e.from === node.id || e.to === node.id;
    });
    var willShowOrphan = !connected && (state.nodes || []).length > 1;
    // warning — a merge with fewer than two wired inputs is a no-op passthrough.
    if (node.type.indexOf("merge_") === 0) {
      var wired = ["in1", "in2", "in3"].filter(function (p) {
        return inputWired(node.id, p);
      }).length;
      if (wired < 2) warnings.push("Merge needs 2+ inputs to combine");
    } else {
      // warning — data inputs exist but none wired; skipped when the orphan warning fires.
      var dataInputs = (type.inputs || []).filter(function (p) {
        return p.type !== "control";
      });
      var noneWired =
        dataInputs.length &&
        !dataInputs.some(function (p) {
          return inputWired(node.id, p.name);
        });
      if (noneWired && !willShowOrphan) warnings.push("Wire at least one input");
    }
    // warning — an orphan node (no incident edges) in a multi-node graph.
    if (willShowOrphan) {
      warnings.push("Not connected to the graph");
    }

    return { errors: errors, warnings: warnings };
  }

  // Kahn's algorithm, ported from workflows.topo_order; returns the ids it could not place.
  function cycleNodeIds() {
    var nodes = state.nodes || [];
    var ids = {};
    nodes.forEach(function (n) {
      ids[n.id] = true;
    });
    var indeg = {};
    var adj = {};
    nodes.forEach(function (n) {
      indeg[n.id] = 0;
      adj[n.id] = [];
    });
    (state.edges || []).forEach(function (e) {
      if (ids[e.from] && ids[e.to]) {
        adj[e.from].push(e.to);
        indeg[e.to] += 1;
      }
    });
    var ready = [];
    nodes.forEach(function (n) {
      if (indeg[n.id] === 0) ready.push(n.id);
    });
    var placed = {};
    while (ready.length) {
      var nid = ready.shift();
      placed[nid] = true;
      adj[nid].forEach(function (nxt) {
        indeg[nxt] -= 1;
        if (indeg[nxt] === 0) ready.push(nxt);
      });
    }
    return nodes
      .filter(function (n) {
        return !placed[n.id];
      })
      .map(function (n) {
        return n.id;
      });
  }

  function graphHasCycle() {
    return cycleNodeIds().length > 0;
  }

  // ---- Dry-run preview (what would execute) ----

  // Nodes a Run (or "Run to here") would execute; mirrors the runner's _should_skip minus gates.
  function computeWouldRun(targetNodeId) {
    var nodes = (state.nodes || []).filter(function (n) {
      return n.type !== "note";
    });
    var ids = {};
    nodes.forEach(function (n) {
      ids[n.id] = true;
    });
    var skip = {};
    nodes.forEach(function (n) {
      if (n.disabled) skip[n.id] = true;
    });
    var edges = state.edges || [];
    for (var iter = 0; iter < nodes.length; iter++) {
      var changed = false;
      for (var i = 0; i < nodes.length; i++) {
        var n = nodes[i];
        if (skip[n.id]) continue;
        var inputs = catalogType(n).inputs || [];
        for (var p = 0; p < inputs.length; p++) {
          var port = inputs[p];
          if (port.optional || port.type === "control") continue;
          var producers = edges.filter(function (e) {
            return e.to === n.id && e.toPort === port.name && ids[e.from];
          });
          var allDead =
            producers.length &&
            producers.every(function (e) {
              return skip[e.from];
            });
          if (allDead) {
            skip[n.id] = true;
            changed = true;
            break;
          }
        }
      }
      if (!changed) break;
    }
    var would = {};
    nodes.forEach(function (n) {
      if (!skip[n.id]) would[n.id] = true;
    });
    // "Run to here": keep only the target's ancestors (runner's _ancestors_inclusive).
    if (targetNodeId && ids[targetNodeId]) {
      var keep = {};
      var stack = [targetNodeId];
      while (stack.length) {
        var nid = stack.pop();
        if (keep[nid]) continue;
        keep[nid] = true;
        for (var e2 = 0; e2 < edges.length; e2++) {
          if (edges[e2].to === nid && ids[edges[e2].from]) {
            stack.push(edges[e2].from);
          }
        }
      }
      Object.keys(would).forEach(function (wid) {
        if (!keep[wid]) delete would[wid];
      });
    }
    return { ids: would, count: Object.keys(would).length, total: nodes.length };
  }

  // Hover cue: preview classes on cards plus the steps chip. Never touches run-* classes.
  function showRunPreview(targetNodeId) {
    if (!state.ready) return;
    var plan = computeWouldRun(targetNodeId);
    var cards = qsa("#wfWorld .wf-node");
    for (var i = 0; i < cards.length; i++) {
      var id = cards[i].getAttribute("data-node-id");
      var isNote = cards[i].getAttribute("data-node-type") === "note";
      cards[i].classList.toggle("wf-preview-run", !isNote && !!plan.ids[id]);
      cards[i].classList.toggle("wf-preview-skip", !isNote && !plan.ids[id]);
    }
    var chip = qs("#wfPreviewChip");
    if (chip) {
      var steps = plan.count === 1 ? " step" : " steps";
      var text =
        plan.count === plan.total
          ? plan.count + steps
          : plan.count + " of " + plan.total + " steps";
      // Count of video-duration-bound steps; a cost hint, not an ETA.
      var heavy = 0;
      (state.nodes || []).forEach(function (n) {
        if (plan.ids[n.id] && isHeavyNodeType(n.type)) heavy += 1;
      });
      if (heavy) text += " · " + heavy + " heavy";
      chip.textContent = text;
      chip.classList.toggle("hidden", !plan.total);
    }
  }

  // Video-duration-bound node types: whole-recording decode (detectors,
  // multitool, timelapse), transcription, or a whole-file rewrite/copy.
  function isHeavyNodeType(type) {
    if (String(type).indexOf("ss_") === 0) return true;
    return (
      type === "detect" ||
      type === "multitool" ||
      type === "timelapse" ||
      type === "transcribe" ||
      type === "post_process"
    );
  }

  function clearRunPreview() {
    var cards = qsa("#wfWorld .wf-node");
    for (var i = 0; i < cards.length; i++) {
      cards[i].classList.remove("wf-preview-run");
      cards[i].classList.remove("wf-preview-skip");
    }
    var chip = qs("#wfPreviewChip");
    if (chip) chip.classList.add("hidden");
  }

  // ---- Aggregate + render ----

  function nodeLabel(node) {
    if (node.name) return node.name;
    var type = state.catalogById[node.type];
    return (type && type.label) || node.type;
  }

  function compute() {
    var errors = [];
    var warnings = [];
    // A cycle blocks scheduling; anchor the row on the first unplaced node.
    var cyc = cycleNodeIds();
    if (cyc.length) {
      var cycLabels = cyc.slice(0, 3).map(function (id) {
        var n = WF.findNode ? WF.findNode(id) : null;
        return n ? nodeLabel(n) : id;
      });
      if (cyc.length > 3) cycLabels.push("+" + (cyc.length - 3) + " more");
      errors.push({
        message: "Graph has a cycle",
        nodeId: cyc[0],
        label: cycLabels.join(", "),
      });
    }
    (state.nodes || []).forEach(function (node) {
      var issues = nodeIssues(node);
      var label = nodeLabel(node);
      issues.errors.forEach(function (m) {
        errors.push({ message: m, nodeId: node.id, label: label });
      });
      issues.warnings.forEach(function (m) {
        warnings.push({ message: m, nodeId: node.id, label: label });
      });
    });
    return { errors: errors, warnings: warnings };
  }

  function buildRow(issue, severity) {
    var row = el("div", "wf-issue wf-issue-" + severity);
    row.appendChild(el("span", "wf-issue-dot"));
    var body = el("div", "wf-issue-body");
    body.appendChild(el("span", "wf-issue-msg", issue.message));
    if (issue.label) body.appendChild(el("span", "wf-issue-node", issue.label));
    row.appendChild(body);
    if (issue.nodeId) {
      row.classList.add("clickable");
      row.addEventListener("click", function () {
        if (WF.focusNode) WF.focusNode(issue.nodeId);
      });
    }
    return row;
  }

  function render() {
    var panel = qs("#wfValidation");
    if (!panel) return;
    var v = state.validation || { errors: [], warnings: [] };
    var total = v.errors.length + v.warnings.length;
    panel.classList.toggle("hidden", total === 0);
    panel.innerHTML = "";
    if (!total) return;

    var head = el("div", "wf-validation-head");
    head.appendChild(el("span", "wf-validation-title", "Issues"));
    var parts = [];
    if (v.errors.length) {
      parts.push(v.errors.length + (v.errors.length > 1 ? " errors" : " error"));
    }
    if (v.warnings.length) {
      parts.push(
        v.warnings.length + (v.warnings.length > 1 ? " warnings" : " warning"),
      );
    }
    head.appendChild(el("span", "wf-validation-count", parts.join(" · ")));
    panel.appendChild(head);

    var list = el("div", "wf-issue-list");
    var frag = document.createDocumentFragment();
    v.errors.forEach(function (e) {
      frag.appendChild(buildRow(e, "error"));
    });
    v.warnings.forEach(function (w) {
      frag.appendChild(buildRow(w, "warning"));
    });
    list.appendChild(frag);
    panel.appendChild(list);
  }

  function refreshValidation() {
    state.validation = compute();
    render();
    if (WF.syncRunButton) WF.syncRunButton();
    // The trigger toggle shares the Run gate (can't arm a graph with errors).
    if (WF.syncTriggerButton) WF.syncTriggerButton();
  }

  // ---- Satellite interface ----
  WF.nodeIssues = nodeIssues;
  WF.graphHasCycle = graphHasCycle;
  WF.refreshValidation = refreshValidation;
  // Dry-run preview (hub wires the Run split-button hover to these).
  WF.computeWouldRun = computeWouldRun;
  WF.showRunPreview = showRunPreview;
  WF.clearRunPreview = clearRunPreview;
})();
