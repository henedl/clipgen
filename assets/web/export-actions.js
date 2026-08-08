/* Shared Export quick-action wiring for Studio / Screenspace / Transcripts.
 *
 * Endpoints /api/export and /api/export/status are registered on the combined
 * Flask app (server.py), so the same client logic drives all three surfaces.
 *
 * This is *not* a download: the server writes clipgen_export_<surface>.json and
 * .csv straight into the output directory (data_export.write_export_bundle),
 * and the only data it covers is the Screenspace and Transcripts manifests —
 * never generated clips. The label and tooltips have to say all of that, since
 * the item also shows up verbatim in the command palette.
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

  // "clipgen_export_screenspace_events.json" -> "screenspace_events", so the
  // toast names what was exported instead of only counting files (each surface
  // writes one .json and one .csv, hence the dedupe).
  function surfaceNames(written) {
    var seen = [];
    (written || []).forEach(function (name) {
      var surface = name.replace(/^clipgen_export_/, "").replace(/\.(json|csv)$/, "");
      if (seen.indexOf(surface) === -1) seen.push(surface);
    });
    if (seen.length <= 3) return seen.join(", ");
    return seen.slice(0, 3).join(", ") + " +" + (seen.length - 3) + " more";
  }

  function runExport() {
    // Manual fetch (not apiPost) so a server-supplied j.error on non-2xx still
    // surfaces in the toast — apiPost would throw "Server error <code>" and
    // drop the body.
    fetch("/api/export", { method: "POST" })
      .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }); })
      .then(function (res) {
        if (res.ok && res.j && res.j.ok) {
          var written = res.j.written || [];
          var names = surfaceNames(written);
          showToast(
            "Exported " + (names ? names + " — " : "") +
              clipgenPluralUnit(written.length, "file", "files") + " in " + res.j.output_dir
          );
        } else {
          showToast((res.j && res.j.error) || "Export failed");
        }
      })
      .catch(function (err) { showToast("Export failed: " + err.message); });
  }

  function exportQuickAction() {
    return {
      icon: "arrow-down-tray",
      label: "Export analysis data (JSON+CSV)",
      action: runExport,
      disabled: !_enabled,
      title: _enabled
        ? "Write Screenspace and Transcript analysis data into the output folder as clipgen_export_*.json and .csv tables. Does not export clips."
        : "Run a Screenspace scan or transcribe a video first — this exports their data, not your clips.",
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
