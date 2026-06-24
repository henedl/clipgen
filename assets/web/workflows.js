/* Workflows hub — the node-canvas frontend (4th top-level surface).
 *
 * Establishes the window.ClipgenWorkflows (WF) namespace that satellite files
 * (workflows-canvas / -wires / -nodes / -runs / -stashes) will share state and
 * functions through — mirroring screenspace.js + window.ClipgenScreenspace and
 * transcripts.js + window.ClipgenTranscripts.
 *
 * M0 (scaffold) only boots the page and publishes `state`. The canvas, typed
 * ports/wires, node catalog, run engine, and stashes arrive in milestones
 * M1-M5 — see plans/WORKFLOWS-PLAN.md. Vanilla JS, ES5-style (no build step).
 */

(function () {
  "use strict";

  // Shared mutable state. Routed through WF.state (not bare `var`s) so future
  // satellites read/write the same object without cross-file ReferenceErrors —
  // the carve gotcha that bit Screenspace and Transcripts.
  var state = {
    catalog: null, // node-type registry, fetched from /api/catalog (M2)
    nodes: [], // placed cards: {id, type, params, position:{x,y}}
    edges: [], // wires: {from, fromPort, to, toPort}
    viewport: { x: 0, y: 0, zoom: 1 }, // pan/zoom (M1)
    selection: [], // selected node ids
    activeBlueprintId: null,
  };

  function boot() {
    // TopNav renders the theme toggle (#themeToggle) and Settings (#settingsBtn)
    // buttons synchronously before this hub loads, so wire them here as the
    // other surfaces do. M0 has no page-specific settings, so Settings just
    // opens the shared modal; page-scoped onSave/onReset land with later config.
    if (typeof initThemeToggle === "function") {
      initThemeToggle();
    }
    var settingsBtn = qs("#settingsBtn");
    if (settingsBtn && typeof window.openSettingsModal === "function") {
      settingsBtn.addEventListener("click", function () {
        window.openSettingsModal({});
      });
    }
  }

  // ---- Satellite interface (window.ClipgenWorkflows) ----
  // The hub publishes `state` (and, in later milestones, shared helpers) onto
  // this namespace; satellites attach their own functions back onto it.
  var WF = (window.ClipgenWorkflows = window.ClipgenWorkflows || {});
  WF.state = state;
  WF.boot = boot;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
