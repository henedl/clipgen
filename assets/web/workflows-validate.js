/* Workflows — pre-run validation panel (satellite of workflows.js).
 *
 * Aggregates the graph's pre-run issues into one panel (#wfValidation) and gates
 * the Run button on errors (warnings never block). Recomputed on every edit via
 * WF.refreshValidation, which the hub calls from scheduleSave + openBlueprint.
 * Owns the per-node issue computation (WF.nodeIssues) — shared with the node-card
 * cue in workflows-nodes.js — plus a JS port of workflows.topo_order's cycle
 * check (the server 400 stays a backstop). Reads shared state through WF.state —
 * never re-`var`s a divergent `state` (the carve gotcha).
 */

(function () {
  "use strict";

  var WF = window.ClipgenWorkflows;
  var state = WF.state;

  // Heatmap style → the detector node that produces the matching raw_results
  // upstream. A mismatch is a warning, not an error (the scan still runs).
  var HEATMAP_STYLE_SOURCE = {
    template: "ss_template",
    flow: "ss_flow",
    change: "ss_change",
  };

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

  // Per-node issues: {errors:[msg…], warnings:[msg…]}. Single source of truth for
  // both the node-card cue (workflows-nodes.js) and the panel rows.
  function nodeIssues(node) {
    var type = catalogType(node);
    var errors = [];
    var warnings = [];

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

    // warning — heatmap style needs a matching upstream detector.
    if (node.type === "heatmap") {
      var want = HEATMAP_STYLE_SOURCE[(node.params || {}).style || "change"];
      var up = upstreamType(node.id, "events");
      if (want && up && up !== want && up !== "multitool") {
        warnings.push("Heatmap style needs a " + want + " (or multitool) upstream");
      }
    }
    // warning — a gate with no scalar source can't evaluate.
    if (node.type === "gate" && !inputWired(node.id, "value")) {
      warnings.push("Gate has no scalar source");
    }
    // warning — a merge with fewer than two wired inputs is a no-op passthrough.
    if (node.type.indexOf("merge_") === 0) {
      var wired = ["in1", "in2", "in3"].filter(function (p) {
        return inputWired(node.id, p);
      }).length;
      if (wired < 2) warnings.push("Merge needs 2+ inputs to combine");
    }
    // warning — an orphan node (no incident edges) in a multi-node graph.
    var connected = (state.edges || []).some(function (e) {
      return e.from === node.id || e.to === node.id;
    });
    if (!connected && (state.nodes || []).length > 1) {
      warnings.push("Not connected to the graph");
    }

    return { errors: errors, warnings: warnings };
  }

  // Kahn cycle check, ported from workflows.topo_order. Control edges are real
  // dependencies, so they count; edges to unknown nodes are ignored (a stale wire
  // never blocks). Returns true iff the graph contains a cycle.
  function graphHasCycle() {
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
    var seen = 0;
    while (ready.length) {
      var nid = ready.shift();
      seen += 1;
      adj[nid].forEach(function (nxt) {
        indeg[nxt] -= 1;
        if (indeg[nxt] === 0) ready.push(nxt);
      });
    }
    return seen !== nodes.length;
  }

  // ---- Aggregate + render ----

  function nodeLabel(node) {
    var type = state.catalogById[node.type];
    return (type && type.label) || node.type;
  }

  function compute() {
    var errors = [];
    var warnings = [];
    // Graph-level: a cycle makes the run unschedulable (Kahn never drains).
    if (graphHasCycle()) {
      errors.push({ message: "Graph has a cycle", nodeId: null });
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

  // Recompute from the live graph, re-render the panel, and re-gate Run. Called
  // by the hub on every edit (scheduleSave) and on blueprint load (openBlueprint).
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
})();
