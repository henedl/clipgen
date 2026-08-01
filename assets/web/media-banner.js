// Media banner: warns when a participant's source recording is a fragmented
// MP4 the browser cannot seek, and drives the one-click remux that fixes it.
//
// Shared by Transcripts, Composer and Screenspace — the three pages that play
// source video — so the wording, the poll and the recovery flow exist once.
// CSS lives in media-banner.css.
//
// Why this exists: OBS's "fragmented recording" writes an MP4 with no movie
// duration and no sample index. ffmpeg reads it via the mfra tail box, so every
// server-side path is fine; browsers do not read mfra, so video.duration comes
// back Infinity and a seek lands wherever the download happens to have reached.
// See video.probe_container_seekability for the backend half.

(function () {
  var POLL_MS = 1000;

  // One banner per page; the pages only ever show one participant at a time.
  var _host = null; // { el, participant }
  var _pollTimer = 0;
  var _polling = false;

  // A remux or a restore replaces the file the <video> is streaming, so the
  // page has to start over: the element is holding a part-downloaded body for
  // a byte range that no longer exists, and every cache-bust token derived from
  // the old mtime is stale. All three pages persist their selected participant,
  // so a reload lands back where the user was. Discarding the kept original
  // changes nothing playable and deliberately does not reload.
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

    var text = el("span", "media-banner-text", state.text);
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
        // A multi-part participant can succeed on some parts and fail on
        // others; the server reports that as `warnings`. Claiming a flat
        // "Original deleted." there would be a lie, and the untouched parts
        // would silently keep their backups.
        var warnings = (result && result.warnings) || [];
        showToast(
          warnings.length
            ? label + " " + warnings.length + " part(s) unchanged: " + warnings.join(" ")
            : label
        );
        // Still reload after a partial restore — the parts that did swap left
        // the player streaming a file that no longer exists. The rebuilt banner
        // reports whichever backups are still on disk.
        if (reloads) reloadAfterFileSwap();
        else refresh();
      })
      .catch(function (error) {
        showToast("Failed: " + error.message);
      });
  }

  // ---- Polling ----

  function stopPolling() {
    if (_pollTimer) clearTimeout(_pollTimer);
    _pollTimer = 0;
    _polling = false;
  }

  function startPolling() {
    if (_polling) return;
    _polling = true;
    poll();
  }

  function poll() {
    if (!_polling || !_host) return;
    // Never poll a hidden tab; the job keeps running server-side and the next
    // visible tick picks up wherever it got to.
    if (document.hidden) {
      _pollTimer = setTimeout(poll, POLL_MS);
      return;
    }
    apiGet("api/remux/status")
      .then(function (data) {
        if (!_host) return;
        var job = (data.jobs || {})[_host.participant];
        if (!job || job.state === "running") {
          render({
            mode: "running",
            text: "Remuxing " + _host.participant + "…",
            progress: job ? job.progress : 0,
          });
          _pollTimer = setTimeout(poll, POLL_MS);
          return;
        }
        stopPolling();
        if (job.state === "error") {
          renderError(job.error);
          return;
        }
        showToast("Remux complete — " + _host.participant + " is now seekable.");
        reloadAfterFileSwap();
      })
      .catch(function () {
        // A transient poll failure is not a job failure; try again next tick.
        _pollTimer = setTimeout(poll, POLL_MS);
      });
  }

  // Shared by the live poll and by refresh(), which has to rebuild this state
  // from scratch when the user leaves a participant mid-remux and comes back.
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
      text: "Remuxed. Original kept as " + names.join(", ") + ".",
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
          tooltip: "Put the original recording back and discard the remux",
          onClick: function () {
            originalAction("restore-original", "Original restored.", true);
          },
        },
      ],
    });
  }

  // ---- Public API ----

  // Point the banner at a participant. `entry` is the /api/participants record;
  // `browser_seekable === false` is the only value that warns — null/undefined
  // means the probe could not classify the container (not an MP4, unreadable)
  // and guessing there would flag every non-MP4 source.
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

  // Re-read server state and repaint. Called after a remux/discard/restore, and
  // on show(); keeps the "original kept" state correct across page reloads,
  // where no job is left in memory to report it.
  function refresh(entry) {
    if (!_host) return;
    // Only a fallback for a failed status fetch. The page's /api/participants
    // snapshot is stale the moment anything rewrites a file — including a remux
    // run from Composer or Screenspace against the same input dir — so the
    // status response, which re-probes, is the authority.
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
        // Every branch below has to survive the user switching away mid-remux
        // and back: show() stopped the poll, so this is the only thing that
        // rebuilds the finished state.
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
