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
    // A muted node never runs, so it can't block the run — report no issues.
    // Sticky notes are annotations, not executable nodes: same deal (kills the
    // orphan/no-input warnings a port-less card would otherwise collect).
    if (node.disabled || node.type === "note") {
      return { errors: [], warnings: [] };
    }
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
    // error — the Detect node's active detector has its own (swapped-in) params;
    // check their required flags against the hidden ss_<detector> spec node.
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
      // The heatmap style names map 1:1 to the ss_<style> detector that produces
      // the matching raw_results (template→ss_template, …) — derive it directly.
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
    // warning — a Video Source with an empty participant array has nothing to run
    // (the multi-select stores [] when every box is unchecked). Not a hard error:
    // the run/batch simply produces no clips for it.
    if (
      node.type === "video_source" &&
      Array.isArray((node.params || {}).participant) &&
      node.params.participant.length === 0
    ) {
      warnings.push("No participants selected");
    }
    // warning — a filter/partition with an ordering comparison (>=,>,<=,<) needs a
    // numeric value; a non-numeric one fails the backend float() coerce and
    // silently drops every item. (Heuristic on the op, so no need to mirror the
    // backend's per-field numeric/text table.)
    if (
      node.type.indexOf("filter_") === 0 ||
      node.type.indexOf("partition_") === 0
    ) {
      var op = (node.params || {}).op;
      var val = (node.params || {}).value;
      var ordering = op === ">=" || op === ">" || op === "<=" || op === "<";
      var numeric = /^\s*-?(\d+\.?\d*|\.\d+)\s*$/.test(String(val));
      if (ordering && !paramEmpty(val) && !numeric) {
        warnings.push("Value must be a number for this comparison");
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
      // warning — a node with data input ports but none wired runs but produces
      // nothing (e.g. make_clips / measure, whose inputs are all optional so the
      // required-input check above never fires). Suppressed when the clearer
      // "not connected" orphan message below will fire instead.
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

  // ---- Dry-run preview (what would execute) ----

  // Estimate the set of nodes a Run (or "Run to here" with targetNodeId) would
  // execute: everything minus sticky notes, muted nodes, and nodes whose
  // *required* data inputs are fed only by skipped producers (a JS mirror of the
  // runner's _should_skip, minus gate evaluation — gates resolve at run time, so
  // gated branches count as running). Bounded relaxation keeps it cycle-safe.
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
    // "Run to here": intersect with the target's ancestors (inclusive) — a JS
    // port of the runner's _ancestors_inclusive reverse walk.
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

  // Toggle the preview classes on the canvas cards + the "N of M steps" chip.
  // clearRunPreview removes ONLY its own classes — never the run-* tint set the
  // runs satellite owns. (A renderAllNodes rebuild drops the preview classes;
  // re-hovering Run re-applies them, which is fine for a hover-scoped cue.)
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
      chip.textContent =
        plan.count === plan.total
          ? plan.count + steps
          : plan.count + " of " + plan.total + " steps";
      chip.classList.toggle("hidden", !plan.total);
    }
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
  // Dry-run preview (hub wires the Run split-button hover to these).
  WF.computeWouldRun = computeWouldRun;
  WF.showRunPreview = showRunPreview;
  WF.clearRunPreview = clearRunPreview;
})();
