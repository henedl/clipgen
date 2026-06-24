/* Workflows — node card rendering (satellite of workflows.js).
 *
 * Renders a placed node generically from its catalog NodeType: title, domain
 * accent, and static input/output port markers (M2 makes the markers wire
 * connectors; param editors also land in M2). Reads shared state through
 * WF.state — never re-`var`s a divergent `state` (the carve gotcha).
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
      var row = el("div", "wf-port" + (port.optional ? " optional" : ""));
      var dot = el("span", "wf-port-dot");
      dot.setAttribute("data-port", port.name);
      dot.setAttribute("data-port-type", port.type);
      dot.setAttribute("data-port-dir", isOutput ? "out" : "in");
      row.appendChild(dot);
      row.appendChild(el("span", "wf-port-label", port.name));
      col.appendChild(row);
    });
    return col;
  }

  function renderNode(node) {
    var type = state.catalogById[node.type] || {
      label: node.type,
      domain: "",
      inputs: [],
      outputs: [],
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

    card.appendChild(el("div", "wf-node-title", type.label || node.type));
    card.appendChild(el("div", "wf-node-domain", type.domain || ""));

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
  }

  WF.renderNode = renderNode;
  WF.renderAllNodes = renderAllNodes;
})();
