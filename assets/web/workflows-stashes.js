/* Workflows — stash palette (satellite of workflows.js).
 *
 * Owns the stash sidebar: load the built-in recipes + user stashes, render them
 * as draggable/clickable items, save the current selection as a new stash, and
 * instantiate one onto the canvas as a fresh copy (id remap + position offset).
 * The server does CRUD only; instantiation is client-side here. Cross-file calls
 * go through WF.* late-bound, so load order stays flexible.
 */

(function () {
  "use strict";

  var WF = window.ClipgenWorkflows;
  var state = WF.state;

  // Successive click-to-add instantiations cascade so they don't stack exactly.
  var _clickCascade = 0;

  // ---- Data -----------------------------------------------------------------

  function findStash(id) {
    var arr = state.stashes || [];
    for (var i = 0; i < arr.length; i++) {
      if (arr[i].id === id) return arr[i];
    }
    return null;
  }

  // GET returns the read-only built-in recipes first, then the user's stashes.
  function loadStashes() {
    return apiGet("api/stashes").then(function (res) {
      state.stashes = (res && res.stashes) || [];
      renderStashPalette();
      renderEmptyRecipes();
    });
  }

  // Built-in recipes as one-click chips inside #wfCanvasEmpty (hidden with the empty state).
  function renderEmptyRecipes() {
    var host = qs("#wfEmptyRecipes");
    if (!host) return;
    host.innerHTML = "";
    var builtins = (state.stashes || []).filter(function (s) {
      return s.builtin;
    });
    if (!builtins.length) return;
    host.appendChild(el("div", "wf-empty-recipes-label", "Start from a recipe"));
    var row = el("div", "wf-empty-recipes-row");
    builtins.forEach(function (stash) {
      var chip = el("button", "wf-recipe-chip", stash.name);
      chip.type = "button";
      chip.title = "Add this recipe to the canvas";
      chip.addEventListener("click", function () {
        if (state.ready) instantiateStash(stash.id, null);
      });
      row.appendChild(chip);
    });
    host.appendChild(row);
  }

  // ---- Render ---------------------------------------------------------------

  function renderStashPalette() {
    var list = qs("#wfStashList");
    if (!list) return;
    list.removeAttribute("aria-busy"); // boot skeletons are about to go
    list.innerHTML = "";
    var stashes = state.stashes || [];
    if (!stashes.length) {
      list.appendChild(
        el("div", "wf-stash-empty", "No stashes yet. Select nodes, then Stash them.")
      );
      return;
    }
    var frag = document.createDocumentFragment();
    stashes.forEach(function (stash) {
      frag.appendChild(buildStashItem(stash));
    });
    list.appendChild(frag);
  }

  function buildStashItem(stash) {
    var item = el("div", "wf-stash-item");
    if (stash.builtin) item.classList.add("wf-stash-builtin");
    item.dataset.stashId = stash.id;
    item.draggable = true;
    // data-tooltip, not title: native title doesn't render on draggable rows.
    item.setAttribute(
      "data-tooltip",
      stash.builtin
        ? "Built-in recipe. Drag or click to add to the canvas"
        : "Drag or click to add to the canvas"
    );

    item.appendChild(el("span", "wf-stash-label", stash.name));

    if (stash.builtin) {
      // A small read-only badge in place of the edit/delete controls.
      var badge = el("span", "wf-stash-badge");
      badge.title = "Built-in recipe (read-only)";
      item.appendChild(badge);
    } else {
      item.appendChild(iconButton("wf-stash-rename", "Rename stash", function () {
        WF.openPromptDialog({
          title: "Rename stash",
          initial: stash.name,
          confirmLabel: "Rename",
          onConfirm: function (name) {
            name = (name || "").trim();
            if (name && name !== stash.name) renameStash(stash.id, name);
          },
        });
      }));
      item.appendChild(iconButton("wf-stash-delete", "Delete stash", function () {
        WF.openConfirmDialog({
          title: "Delete stash “" + stash.name + "”?",
          body: "This can't be undone.",
          danger: true,
          onConfirm: function () {
            deleteStash(stash.id);
          },
        });
      }));
    }

    item.addEventListener("dragstart", function (e) {
      if (!state.ready) {
        e.preventDefault(); // don't start a drag the canvas can't accept yet
        return;
      }
      e.dataTransfer.setData("application/x-wf-stash", stash.id);
      e.dataTransfer.setData("text/plain", stash.name);
      e.dataTransfer.effectAllowed = "copy";
    });

    // Click-to-add at viewport center; the icon buttons stopPropagation.
    item.addEventListener("click", function () {
      if (!state.ready) return;
      instantiateStash(stash.id, null);
    });

    return item;
  }

  // Icon button (glyph via CSS mask) that swallows the click so click-to-add doesn't fire.
  function iconButton(cls, title, onClick) {
    var btn = el("button", cls);
    btn.type = "button";
    btn.title = title;
    btn.appendChild(el("span", "wf-stash-btn-icon"));
    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      onClick();
    });
    return btn;
  }

  // ---- Save / instantiate ---------------------------------------------------

  function saveSelectionAsStash() {
    var sel = state.selection || [];
    if (!sel.length) return;
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
    if (!nodes.length) return;

    // Induced edges only — both endpoints selected, so a stash never carries a
    // dangling half-edge.
    var edges = (state.edges || [])
      .filter(function (e) {
        return selSet[e.from] && selSet[e.to];
      })
      .map(function (e) {
        return JSON.parse(JSON.stringify(e));
      });

    WF.openPromptDialog({
      title: "Name this stash",
      initial: "Stash",
      confirmLabel: "Save",
      onConfirm: function (name) {
        name = (name || "").trim() || "Stash";
        apiPost("api/stashes", { name: name, nodes: nodes, edges: edges })
          .then(function (res) {
            if (!res || !res.stash) return;
            // Insert after the leading built-ins so user stashes stay grouped
            // below.
            var arr = state.stashes || (state.stashes = []);
            var idx = 0;
            while (idx < arr.length && arr[idx].builtin) idx++;
            arr.splice(idx, 0, res.stash);
            renderStashPalette();
            showToast("Saved stash “" + res.stash.name + "”");
          })
          .catch(function () {
            showToast("Failed to save stash");
          });
      },
    });
  }

  // Stamp deep-cloned nodes and edges with fresh ids; shared with clipboard paste. Returns new ids.
  function instantiateSubgraph(nodes, edges, dropWorld) {
    if (!nodes || !nodes.length) return [];

    // Fresh id for every node, built BEFORE any edge is rewritten.
    var idMap = {};
    nodes.forEach(function (n) {
      idMap[n.id] = "n_" + WF.randomId();
    });

    // Anchor the top-left node at the drop point or cascaded center so nothing stacks.
    var minX = Infinity;
    var minY = Infinity;
    nodes.forEach(function (n) {
      var p = n.position || { x: 0, y: 0 };
      if (p.x < minX) minX = p.x;
      if (p.y < minY) minY = p.y;
    });
    var anchor;
    if (dropWorld) {
      anchor = dropWorld;
    } else {
      var c = viewportCenterWorld();
      var step = (_clickCascade++ % 5) * 40;
      anchor = { x: c.x + step, y: c.y + step };
    }
    var offX = Math.round(anchor.x - minX);
    var offY = Math.round(anchor.y - minY);

    var newIds = [];
    nodes.forEach(function (n) {
      var p = n.position || { x: 0, y: 0 };
      var node = JSON.parse(JSON.stringify(n));
      node.id = idMap[n.id];
      node.position = { x: p.x + offX, y: p.y + offY };
      state.nodes.push(node);
      newIds.push(node.id);
    });

    (edges || []).forEach(function (e) {
      var from = idMap[e.from];
      var to = idMap[e.to];
      if (!from || !to) return; // defensive: drop an edge with an unmapped endpoint
      state.edges.push({
        id: "e_" + WF.randomId(),
        from: from,
        fromPort: e.fromPort,
        to: to,
        toPort: e.toPort,
      });
    });

    state.selection = newIds;
    state.selectedEdge = null;
    if (WF.renderAllNodes) WF.renderAllNodes();
    WF.scheduleSave();
    return newIds;
  }

  // Recipes ship with empty required params; focus the first card that needs one.
  function instantiateStash(stashId, dropWorld) {
    var stash = findStash(stashId);
    if (!stash || !stash.nodes || !stash.nodes.length) return;
    var newIds = instantiateSubgraph(stash.nodes, stash.edges, dropWorld);
    for (var i = 0; i < newIds.length; i++) {
      var node = WF.findNode ? WF.findNode(newIds[i]) : null;
      var issues = node && WF.nodeIssues ? WF.nodeIssues(node) : null;
      if (issues && issues.errors.length) {
        if (WF.focusNode) WF.focusNode(newIds[i]);
        break;
      }
    }
  }

  function viewportCenterWorld() {
    var canvas = qs("#wfCanvas");
    if (canvas && WF.clientToWorld) {
      var r = canvas.getBoundingClientRect();
      return WF.clientToWorld(r.left + r.width / 2, r.top + r.height / 2);
    }
    return { x: 0, y: 0 };
  }

  // ---- Rename / delete ------------------------------------------------------

  function renameStash(id, name) {
    return apiPut("api/stashes/" + encodeURIComponent(id), { name: name })
      .then(function (res) {
        var s = findStash(id);
        if (s && res && res.stash) s.name = res.stash.name;
        renderStashPalette();
      })
      .catch(function () {
        showToast("Failed to rename stash");
      });
  }

  function deleteStash(id) {
    return apiDelete("api/stashes/" + encodeURIComponent(id))
      .then(function () {
        state.stashes = (state.stashes || []).filter(function (s) {
          return s.id !== id;
        });
        renderStashPalette();
      })
      .catch(function () {
        showToast("Failed to delete stash");
      });
  }

  // ---- Toolbar gate + init --------------------------------------------------

  // Enabled only with a selection on a ready canvas; renderAllNodes and init call this.
  function syncStashButton() {
    var btn = qs("#wfSaveStash");
    if (!btn) return;
    btn.disabled = !state.ready || !(state.selection && state.selection.length);
  }

  function initStashes() {
    var btn = qs("#wfSaveStash");
    if (btn) btn.addEventListener("click", saveSelectionAsStash);
    syncStashButton();
  }

  // ---- Satellite interface --------------------------------------------------

  WF.initStashes = initStashes;
  WF.loadStashes = loadStashes;
  WF.renderStashPalette = renderStashPalette;
  WF.saveSelectionAsStash = saveSelectionAsStash;
  WF.instantiateStash = instantiateStash;
  WF.instantiateSubgraph = instantiateSubgraph;
  WF.renameStash = renameStash;
  WF.deleteStash = deleteStash;
  WF.syncStashButton = syncStashButton;
})();
