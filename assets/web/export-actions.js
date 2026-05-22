/* Shared Export quick-action wiring for Studio / Screenspace / Transcripts.
 *
 * Endpoints /api/export and /api/export/status are registered on the combined
 * Flask app (server.py), so the same client logic drives all three surfaces.
 *
 * Depends on globals from utils.js: apiGet, showToast, clipgenPluralUnit.
 *
 * Public API on window.ClipgenExportActions:
 *   runExport()                   — POST /api/export, toast the result
 *   exportQuickAction()           — TopNav quick-action item with current enabled state
 *   refreshExportStatus(rebuild)  — GET /api/export/status; calls rebuild() on flag flip
 */
(function () {
  "use strict";

  var _enabled = false;

  function runExport() {
    // Manual fetch (not apiPost) so a server-supplied j.error on non-2xx still
    // surfaces in the toast — apiPost would throw "Server error <code>" and
    // drop the body.
    fetch("/api/export", { method: "POST" })
      .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }); })
      .then(function (res) {
        if (res.ok && res.j && res.j.ok) {
          var n = (res.j.written || []).length;
          showToast("Exported " + clipgenPluralUnit(n, "file", "files") + " to " + res.j.output_dir);
        } else {
          showToast((res.j && res.j.error) || "Export failed");
        }
      })
      .catch(function (err) { showToast("Export failed: " + err.message); });
  }

  function exportQuickAction() {
    return {
      icon: "arrow-down-tray",
      label: "Export",
      action: runExport,
      disabled: !_enabled,
      title: _enabled
        ? "Write JSON+CSV exports of Screenspace and Transcripts manifests"
        : "Run a Screenspace task or transcribe a video first to enable Export.",
    };
  }

  function refreshExportStatus(rebuild) {
    return apiGet("/api/export/status")
      .then(function (j) {
        var next = !!(j && j.any);
        if (next === _enabled) return;
        _enabled = next;
        if (typeof rebuild === "function") rebuild();
      })
      .catch(function () {});
  }

  window.ClipgenExportActions = {
    runExport: runExport,
    exportQuickAction: exportQuickAction,
    refreshExportStatus: refreshExportStatus,
  };
})();
