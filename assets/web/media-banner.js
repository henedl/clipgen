/* Media banner: warns when a participant's source recording is a fragmented
 * MP4 the browser cannot seek, and drives the one-click remux that fixes it.
 *
 * Shared by Transcripts, Composer and Screenspace — the three pages that play
 * source video — so the wording, the poll and the recovery flow exist once.
 * CSS lives in media-banner.css.
 *
 * Why this exists: OBS's "fragmented recording" writes an MP4 with no movie
 * duration and no sample index. ffmpeg reads it via the mfra tail box, so every
 * server-side path is fine — but browsers don't read mfra, so video.duration is
 * Infinity and a seek lands wherever the download happens to have reached. See
 * video.probe_container_seekability for the backend half.
 */

(function () {
  var POLL_MS = 1000;

  // One banner per page; the pages only ever show one participant at a time.
  var _host = null; // { el, participant }
  var _poller = null;

  // A remux or restore swaps the streaming file; reload. Pages persist the selected participant.
  function reloadAfterFileSwap() {
    window.location.reload();
  }

  function config(key, fallback) {
    var cfg = window.CLIPGEN_CONFIG || {};
    return cfg[key] === undefined ? fallback : cfg[key];
  }

  function iconSpan(mode) {
    var name = mode === "done" ? "check-circle" : "exclamation-triangle";
    var span = el("span", "media-banner-icon");
    span.style.maskImage = 'url("icons/' + name + '.svg")';
    span.style.webkitMaskImage = 'url("icons/' + name + '.svg")';
    return span;
  }

  // ---- Rendering ----

  function render(state) {
    if (!_host) return;
    var root = _host.el;
    root.innerHTML = "";
    if (!state) {
      root.classList.add("hidden");
      return;
    }
    root.classList.remove("hidden");
    root.classList.toggle("is-done", state.mode === "done");

    root.appendChild(iconSpan(state.mode));

    var text = el(
      "span",
      "media-banner-text" + (state.mode === "running" ? " cg-shimmer" : ""),
      state.text
    );
    if (state.tooltip) text.setAttribute("data-tooltip", state.tooltip);
    root.appendChild(text);

    if (state.mode === "running") {
      var track = el("span", "media-banner-progress");
      var fill = el("span", "media-banner-progress-fill");
      fill.style.width = Math.round((state.progress || 0) * 100) + "%";
      track.appendChild(fill);
      root.appendChild(track);
      return;
    }

    var actions = el("span", "media-banner-actions");
    (state.actions || []).forEach(function (action) {
      var btn = el("button", "btn btn-small" + (action.primary ? " btn-primary" : ""));
      btn.type = "button";
      btn.textContent = action.label;
      if (action.tooltip) btn.setAttribute("data-tooltip", action.tooltip);
      btn.addEventListener("click", action.onClick);
      actions.appendChild(btn);
    });
    root.appendChild(actions);
    if (window.clipgenInitDataTooltips) window.clipgenInitDataTooltips();
  }

  // ---- Actions ----

  function startRemux() {
    var pid = _host && _host.participant;
    if (!pid) return;
    render({ mode: "running", text: "Remuxing " + pid + "…", progress: 0 });
    apiPost("api/remux/" + encodeURIComponent(pid), {})
      .then(function () {
        startPolling();
      })
      .catch(function (error) {
        render({
          mode: "warn",
          text: "Could not start the remux: " + error.message,
          actions: [{ label: "Retry", primary: true, onClick: startRemux }],
        });
      });
  }

  function originalAction(path, label, reloads) {
    var pid = _host && _host.participant;
    if (!pid) return;
    apiPost("api/remux/" + encodeURIComponent(pid) + "/" + path, {})
      .then(function (result) {
        // Multi-part participants can partially fail; the server reports `warnings`. Say so.
        var warnings = (result && result.warnings) || [];
        showToast(
          warnings.length
            ? label + " " + warnings.length + " part(s) unchanged: " + warnings.join(" ")
            : label
        );
        // Reload even after a partial restore: swapped parts left the player streaming a dead file.
        if (reloads) reloadAfterFileSwap();
        else refresh();
      })
      .catch(function (error) {
        showToast("Failed: " + error.message);
      });
  }

  // ---- Polling ----

  function stopPolling() {
    if (!_poller) return;
    _poller.stop();
    _poller = null;
  }

  // createPoller skips hidden tabs itself; the job continues server-side.
  function startPolling() {
    if (_poller) return;
    _poller = createPoller(poll, POLL_MS, { label: "media-banner.remux" });
    _poller.start();
  }

  function poll() {
    if (!_host) return;
    return apiGet("api/remux/status")
      .then(function (data) {
        if (!_host) return;
        var job = (data.jobs || {})[_host.participant];
        if (job && job.state === "running") {
          render({
            mode: "running",
            text: "Remuxing " + _host.participant + "…",
            progress: job.progress,
          });
          return;
        }
        stopPolling();
        if (!job) {
          // Another page finished or discarded it; show the on-disk state.
          refresh();
          return;
        }
        if (job.state === "error") {
          renderError(job.error);
          return;
        }
        showToast("Remux complete — " + _host.participant + " is now seekable.");
        reloadAfterFileSwap();
      })
      .catch(function () {
        // A transient poll failure is not a job failure; the next tick retries.
      });
  }

  // Shared by the poll and refresh(), which rebuilds state after a participant switch.
  function renderError(reason) {
    render({
      mode: "warn",
      text: reason || "The remux failed; the original is untouched.",
      actions: [{ label: "Try again", primary: true, onClick: startRemux }],
    });
  }

  function renderKept(kept) {
    if (!_host) return;
    var names = kept[_host.participant];
    if (!names || !names.length) {
      render(null);
      return;
    }
    render({
      mode: "done",
      // "Original kept", not "Remuxed": Normalize Audio parks its original in the same .orig slot.
      text: "Original kept as " + names.join(", ") + ".",
      actions: [
        {
          label: "Delete original",
          tooltip: "Free the disk space — this cannot be undone",
          onClick: function () {
            originalAction("discard-original", "Original deleted.", false);
          },
        },
        {
          label: "Restore",
          tooltip: "Put the original recording back and discard the rewrite",
          onClick: function () {
            originalAction("restore-original", "Original restored.", true);
          },
        },
      ],
    });
  }

  // ---- Public API ----

  // `entry` is the /api/participants record; only `browser_seekable === false` warns (null = unclassified).
  function show(container, entry) {
    stopPolling();
    if (!container) return;
    if (!_host || _host.el.parentNode !== container) {
      var node = el("div", "media-banner hidden");
      container.insertBefore(node, container.firstChild);
      _host = { el: node, participant: null };
    }
    _host.participant = entry && entry.id;

    if (!config("mediaContainerWarning", true) || !entry || !entry.id) {
      render(null);
      return;
    }
    refresh(entry);
  }

  // Re-read server state and repaint; after reload no in-memory job records the kept original.
  function refresh(entry) {
    if (!_host) return;
    // Fallback only: the participants snapshot goes stale whenever any page rewrites a file.
    var staleUnseekable = entry ? entry.browser_seekable === false : false;
    apiGet("api/remux/status")
      .then(function (data) {
        if (!_host) return;
        var job = (data.jobs || {})[_host.participant];
        var kept = data.kept || {};
        if (job && job.state === "running") {
          startPolling();
          return;
        }
        // show() stopped the poll, so this alone rebuilds the finished state after a switch.
        if (job && job.state === "error") {
          renderError(job.error);
          return;
        }
        if (kept[_host.participant]) {
          renderKept(kept);
          return;
        }
        if ((data.unseekable || []).indexOf(_host.participant) >= 0) renderWarning();
        else render(null);
      })
      .catch(function () {
        if (staleUnseekable) renderWarning();
      });
  }

  function renderWarning() {
    render({
      mode: "warn",
      text: "This recording is a fragmented MP4 — the browser can't seek it.",
      tooltip:
        "Recorded with OBS's fragmented-recording option. The file has no " +
        "duration or seek index, so the player must download all of it before " +
        "timestamps work, and seeks land at the wrong moment until it does. " +
        "Remuxing rewrites the container only — the video and audio are copied " +
        "untouched.",
      actions: [
        {
          label: "Remux to fix",
          primary: true,
          tooltip: "Stream copy — no re-encode. The original is kept until you delete it.",
          onClick: startRemux,
        },
      ],
    });
  }

  function hide() {
    stopPolling();
    if (_host) render(null);
  }

  function teardown() {
    stopPolling();
    _host = null;
  }

  window.addEventListener("pagehide", teardown);

  window.clipgenMediaBanner = {
    show: show,
    hide: hide,
    refresh: refresh,
    teardown: teardown,
  };
})();
