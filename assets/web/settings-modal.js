/* clipgen shared settings modal — settings-modal.js
 *
 * Exposes window.openSettingsModal({ initialTab, onSave, onReset }).
 * Loaded on every frontend page that needs the settings UI.
 *
 * Tabs are rendered in a fixed order; settings are grouped by the `tab`
 * metadata returned from /api/settings (backed by config.STUDIO_SETTINGS).
 * Within a tab, existing `group` labels are preserved as sub-headings.
 */

(function () {
  var TAB_ORDER = [
    "General",
    "Hotkeys",
    "Video & Clips",
    "Screenspace",
    "Composer",
    "Transcription",
    "Summaries",
    "CLI",
  ];

  // EXIT_MS must match --duration-veil (tokens.css) so veil and overlay finish together.
  var EXIT_MS = 360;

  var _root = null;
  var _tabsEl = null;
  var _panelsEl = null;
  var _statusEl = null;
  var _actionsEl = null;
  var _settings = null;
  var _activeTab = "General";
  var _saveTimer = null;
  var _opts = {};
  var _modelsCache = null;
  var _modelsCachePromise = null;
  // Running GGUF downloads by model ref; a rebuilt row joins the existing poll.
  var _llmDownloadWatch = {};
  // The Summaries tab's model block refresh, set when the block is built.
  var _llmBlockRefresh = null;
  var _closeTimer = null;
  // Titlecard/endcard picker state (shared by the title + end pickers).
  var _cardsCache = null;
  var _cardsCachePromise = null;
  var _cardPickers = [];
  // Desktop-only footer button; built after load, once the server says native window.
  var _revealBtn = null;
  var _SAMPLE_TITLE_TEXT = "Sample description";

  function _getApiRoot() {
    // Settings and models live at the combined-app root, not under a page prefix.
    return "/api";
  }

  function _formatSize(mb) {
    if (mb >= 1024) return (mb / 1024).toFixed(1) + " GB";
    return mb + " MB";
  }

  function _fetchModels() {
    if (_modelsCache) return Promise.resolve(_modelsCache);
    if (_modelsCachePromise) return _modelsCachePromise;
    _modelsCachePromise = apiGet(_getApiRoot() + "/models")
      .then(function (data) {
        // Cache only when the AI server answered; otherwise retry on the next open.
        if (data && data.ok && !(data.llm && data.llm.available === false)) {
          _modelsCache = data;
        } else {
          _modelsCachePromise = null;
        }
        return data;
      })
      .catch(function () { _modelsCachePromise = null; return null; });
    return _modelsCachePromise;
  }

  function _invalidateModels() {
    _modelsCache = null;
    _modelsCachePromise = null;
  }

  // Refetch once and repaint the model block plus every AI-model dropdown.
  function _refreshLlmViews() {
    _invalidateModels();
    if (_llmBlockRefresh) _llmBlockRefresh();
    _refreshLlmSelects();
  }

  function _refreshLlmSelects() {
    if (!_panelsEl) return;
    var sels = _panelsEl.querySelectorAll(".settings-row[data-setting] .settings-model-dropdown");
    for (var i = 0; i < sels.length; i++) {
      var row = sels[i].parentNode;
      while (row && !(row.getAttribute && row.getAttribute("data-setting"))) row = row.parentNode;
      var s = row && _findSetting(row.getAttribute("data-setting"));
      if (s && s.provider === "llm") {
        _loadModelsForSelect(sels[i], s.provider, s.value, s.emptyLabel);
      }
    }
  }

  // Start or join a GGUF download; one shared poll per model survives modal reopen.
  function _watchLlmDownload(model, onProgress) {
    var watch = _llmDownloadWatch[model];
    if (watch) { watch.listeners.push(onProgress); return; }
    watch = { listeners: [onProgress] };
    _llmDownloadWatch[model] = watch;

    function emit(st) {
      for (var i = 0; i < watch.listeners.length; i++) watch.listeners[i](st);
    }
    function finish(st) {
      delete _llmDownloadWatch[model];
      emit(st);
      if (st.succeeded) _refreshLlmViews();
    }

    apiPost(_getApiRoot() + "/models/llm/download", { model: model }).then(function (data) {
      if (!data || !data.ok) {
        finish({ done: true, succeeded: false, error: (data && data.error) || "Download failed" });
        return;
      }
      var misses = 0;
      var poller = createPoller(function () {
        return apiGet(_getApiRoot() + "/models/llm/download-status?model=" + encodeURIComponent(model))
          .then(function (st) {
            if (!st || !st.ok || !st.found) {
              if (++misses >= 20) {
                poller.stop();
                finish({ done: true, succeeded: false, error: "Download failed" });
              }
              return;
            }
            misses = 0;
            if (st.done) { poller.stop(); finish(st); } else emit(st);
          })
          .catch(function () {
            if (++misses >= 20) {
              poller.stop();
              finish({ done: true, succeeded: false, error: "Download failed" });
            }
          });
      }, 1000, { runImmediately: true, label: "settings.llmDownload" });
      poller.start();
    }).catch(function () {
      finish({ done: true, succeeded: false, error: "Download failed" });
    });
  }

  // Downloaded models with show/delete. Deleting a symlinked external model removes only the link.
  function _buildLlmModelsBlock() {
    var wrap = el("div", "settings-llm-models");
    wrap.appendChild(el("div", "settings-group-label", "Downloaded models"));
    var list = el("div", "settings-llm-models-list");
    wrap.appendChild(list);
    wrap.appendChild(el("div", "settings-group-label", "Suggested models"));
    var suggestedList = el("div", "settings-llm-models-list");
    wrap.appendChild(suggestedList);

    function refresh() {
      _fetchModels().then(function (data) {
        list.textContent = "";
        suggestedList.textContent = "";
        var models = (data && data.llm && data.llm.models) || [];
        if (!models.length) {
          list.appendChild(el("div", "settings-model-note", "No models downloaded yet."));
        }
        for (var i = 0; i < models.length; i++) {
          list.appendChild(_buildLlmModelRow(models[i]));
        }
        var suggested = (data && data.llm && data.llm.suggested) || [];
        for (var j = 0; j < suggested.length; j++) {
          suggestedList.appendChild(_buildSuggestedRow(suggested[j]));
        }
      });
    }

    // Friendly name over mono id; without a catalog label the id stays primary.
    function _modelNameBlock(model) {
      var name = el("div", "settings-llm-model-name");
      var title = el("span", "settings-llm-model-title", model.label || model.name);
      if (!model.label) title.classList.add("settings-llm-model-title--mono");
      name.appendChild(title);
      if (model.label) {
        name.appendChild(el("span", "settings-llm-model-id", model.name));
      }
      return name;
    }

    // Only catalog models carry a source page; the rest get no link at all.
    function _modelLinkButton(model) {
      if (!model.model_url) return null;
      var link = el("a", "settings-llm-model-reveal");
      link.href = model.model_url;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.title = "View this model on Hugging Face";
      link.setAttribute("aria-label", "View this model on Hugging Face");
      link.appendChild(
        el("span", "settings-llm-model-icon settings-llm-model-icon--link")
      );
      return link;
    }

    // A curated model: Download with an in-row progress bar, or "Downloaded".
    function _buildSuggestedRow(model) {
      var row = el("div", "settings-llm-model-row");
      var name = _modelNameBlock(model);
      name.appendChild(el("span", "settings-llm-model-desc", model.description));
      if (model.unusable) {
        name.appendChild(el("span", "settings-llm-model-reason", model.unusable));
        row.classList.add("settings-llm-model-row--unusable");
      }
      row.appendChild(name);
      var size = el("span", "settings-llm-model-size", _formatSize(model.size_mb));
      row.appendChild(size);
      var sugLink = _modelLinkButton(model);
      if (sugLink) row.appendChild(sugLink);

      if (model.installed) {
        var done = el("span", "settings-llm-model-state");
        done.appendChild(el("span", "settings-llm-model-icon settings-llm-model-icon--done"));
        done.appendChild(document.createTextNode("Downloaded"));
        row.appendChild(done);
        return row;
      }

      var bar = el("div", "settings-llm-model-bar");
      var fill = el("div", "settings-llm-model-bar-fill");
      bar.appendChild(fill);
      var dlBtn = el("button", "btn btn-small btn-icon");
      dlBtn.type = "button";
      dlBtn.appendChild(el("span", "settings-llm-model-icon settings-llm-model-icon--download"));
      dlBtn.appendChild(document.createTextNode("Download"));
      row.appendChild(dlBtn);

      function onProgress(st) {
        if (st.done) {
          if (!st.succeeded) {
            bar.remove();
            dlBtn.disabled = false;
            size.textContent = _formatSize(model.size_mb);
            _setStatus(st.error || "Download failed");
          }
          return;
        }
        if (st.total > 0) {
          var pct = Math.max(0, Math.min(100, Math.round((st.completed / st.total) * 100)));
          fill.style.width = pct + "%";
          size.textContent = _formatSize(Math.round(st.completed / 1048576)) +
            " / " + _formatSize(model.size_mb);
        }
      }
      function startWatching() {
        dlBtn.disabled = true;
        name.appendChild(bar);
        _watchLlmDownload(model.name, onProgress);
      }
      dlBtn.addEventListener("click", startWatching);
      // A download started before the modal was closed is still running.
      apiGet(_getApiRoot() + "/models/llm/download-status?model=" + encodeURIComponent(model.name))
        .then(function (st) {
          if (st && st.ok && st.found && !st.done) startWatching();
        })
        .catch(function () {});
      return row;
    }

    function _buildLlmModelRow(model) {
      var row = el("div", "settings-llm-model-row");
      var name = _modelNameBlock(model);
      if (model.unusable) {
        // Show the router's own reason, not a generic "incompatible".
        name.appendChild(el("span", "settings-llm-model-reason", model.unusable));
        row.classList.add("settings-llm-model-row--unusable");
      }
      row.appendChild(name);
      if (model.size_mb) {
        row.appendChild(el("span", "settings-llm-model-size", _formatSize(model.size_mb)));
      }
      var link = _modelLinkButton(model);
      if (link) row.appendChild(link);

      var showBtn = el("button", "settings-llm-model-reveal");
      showBtn.type = "button";
      showBtn.title = "Show in file browser";
      showBtn.setAttribute("aria-label", "Show in file browser");
      showBtn.appendChild(
        el("span", "settings-llm-model-icon settings-llm-model-icon--reveal")
      );
      showBtn.addEventListener("click", function () {
        apiPost(_getApiRoot() + "/models/llm/reveal", { model: model.name })
          .catch(function (e) {
            _setStatus((e && e.message) || "Could not open the folder");
          });
      });
      row.appendChild(showBtn);

      var delBtn = el("button", "btn btn-small btn-icon");
      delBtn.type = "button";
      delBtn.appendChild(
        el("span", "settings-llm-model-icon settings-llm-model-icon--delete")
      );
      delBtn.appendChild(document.createTextNode("Delete"));
      delBtn.addEventListener("click", function () {
        delBtn.disabled = true;
        apiDelete(_getApiRoot() + "/models/llm/" + encodeURIComponent(model.name))
          .then(function () { _refreshLlmViews(); })
          .catch(function (e) {
            delBtn.disabled = false;
            _setStatus((e && e.message) || "Delete failed");
          });
      });
      row.appendChild(delBtn);
      return row;
    }

    _llmBlockRefresh = refresh;
    refresh();
    return wrap;
  }

  function _loadModelsForSelect(sel, provider, currentValue, emptyLabel) {
    _fetchModels().then(function (data) {
      if (!data || !data.ok) { sel.disabled = false; return; }

      var models = [];
      if (provider === "whisper") {
        models = (data.whisper && data.whisper.models) || [];
      } else if (provider === "llm") {
        models = (data.llm && data.llm.models) || [];
      }

      sel.innerHTML = "";
      // Optional "inherit" choice (e.g. friction → same as summary model),
      // represented by a blank value.
      if (emptyLabel) {
        var inheritOpt = document.createElement("option");
        inheritOpt.value = "";
        inheritOpt.textContent = emptyLabel;
        if (!currentValue) inheritOpt.selected = true;
        sel.appendChild(inheritOpt);
      }
      // Settings hold HF refs, installed entries file stems; the catalog maps stem to ref.
      var suggested = (provider === "llm" && data.llm && data.llm.suggested) || [];
      var refByStem = {};
      for (var si = 0; si < suggested.length; si++) {
        refByStem[suggested[si].stem] = suggested[si].name;
      }
      var hasCurrentValue = false;
      for (var i = 0; i < models.length; i++) {
        var m = models[i];
        var opt = document.createElement("option");
        opt.value = refByStem[m.name] || m.name;
        // No room for the raw id in an option; it goes in the tooltip.
        var label = m.label || m.name;
        if (m.size_mb) label += " (" + _formatSize(m.size_mb) + ")";
        if (m.description) label += " \u2014 " + m.description;
        // Still selectable: a llama.cpp upgrade may fix it; the mark warns.
        if (m.unusable) label += " \u2014 won't load";
        opt.textContent = label;
        opt.title = m.unusable ? m.name + " \u2014 " + m.unusable : m.name;
        if (opt.value === currentValue || m.name === currentValue) {
          opt.selected = true;
          hasCurrentValue = true;
        }
        sel.appendChild(opt);
      }
      // Selectable before download; the Transcripts gate fetches on first use.
      var group = null;
      for (var sj = 0; sj < suggested.length; sj++) {
        var sm = suggested[sj];
        if (sm.installed) continue;
        if (!group) {
          group = document.createElement("optgroup");
          group.label = "Suggested";
        }
        var sopt = document.createElement("option");
        sopt.value = sm.name;
        sopt.textContent = (sm.label || sm.name) + " (" + _formatSize(sm.size_mb) +
          ") \u2014 not downloaded";
        sopt.title = sm.name + " \u2014 " + sm.description;
        if (sm.name === currentValue) {
          sopt.selected = true;
          hasCurrentValue = true;
        }
        group.appendChild(sopt);
      }
      if (group) sel.appendChild(group);
      if (!hasCurrentValue && currentValue) {
        var custom = document.createElement("option");
        custom.value = currentValue;
        custom.textContent = currentValue + " (current)";
        custom.selected = true;
        sel.insertBefore(custom, sel.firstChild);
      }
      sel.disabled = false;

      // A single "(current)" option looks populated; say next to it when no runtime exists.
      if (provider === "llm") {
        var status = clipgenLlmStatus(data.llm);
        var note = sel.parentNode && sel.parentNode.querySelector(".settings-model-note");
        if (status.state !== "ok" && sel.parentNode) {
          if (!note) {
            note = document.createElement("div");
            note.className = "settings-model-note";
            sel.parentNode.appendChild(note);
          }
          var extra = "";
          if (status.state === "missing" && status.hint.length) {
            extra = " " + status.hint[0];
          }
          note.textContent = status.message + extra;
        } else if (note) {
          note.remove();
        }
      }
    });
  }

  function _buildDom() {
    if (_root) return;

    var overlay = el("div", "settings-overlay cg-modal-overlay cg-modal-veil hidden");
    overlay.id = "settingsOverlay";

    var panel = el("div", "settings-panel");
    overlay.appendChild(panel);

    var header = el("div", "settings-header");
    var title = el("h3", null, "Settings");
    var versionEl = el("span", "settings-version cg-mono");
    versionEl.id = "settingsVersion";
    var closeBtn = el("button", "btn btn-small");
    closeBtn.type = "button";
    closeBtn.textContent = "Close";
    closeBtn.addEventListener("click", _close);
    var rightGroup = el("div", "settings-header-right");
    rightGroup.appendChild(versionEl);
    rightGroup.appendChild(closeBtn);
    header.appendChild(title);
    header.appendChild(rightGroup);
    panel.appendChild(header);

    var tabs = el("div", "settings-tabs");
    tabs.setAttribute("role", "tablist");
    panel.appendChild(tabs);

    var panels = el("div", "settings-tab-panels");
    panels.id = "settingsContent";
    // Focus target for the row cursor on open; programmatic only, never in Tab order.
    panels.tabIndex = -1;
    panel.appendChild(panels);

    var footer = el("div", "settings-footer");
    var actions = el("div", "settings-footer-actions");
    var resetAll = el("button", "btn btn-small settings-reset-all");
    resetAll.type = "button";
    resetAll.textContent = "Reset all to defaults";
    resetAll.setAttribute("data-hotkey", "settings.resetAll");
    resetAll.addEventListener("click", function () { _resetAll(); });
    actions.appendChild(resetAll);
    var status = el("span", "settings-save-status");
    footer.appendChild(actions);
    footer.appendChild(status);
    panel.appendChild(footer);

    overlay.addEventListener("click", function (e) {
      if (e.target === overlay) _close();
    });

    document.body.appendChild(overlay);

    _root = overlay;
    _actionsEl = actions;
    _tabsEl = tabs;
    _panelsEl = panels;
    _statusEl = status;
  }

  // Footer status line; `working` shimmers in-flight messages, terminal messages write flat.
  function _setStatus(text, working) {
    if (!_statusEl) return;
    _statusEl.classList.toggle("cg-shimmer", working === true);
    _statusEl.textContent = text;
  }

  function _open() {
    _buildDom();
    var v = _opts && _opts.version;
    var vEl = document.getElementById("settingsVersion");
    if (vEl) vEl.textContent = v ? "v" + v : "";

    if (_closeTimer) {
      clearTimeout(_closeTimer);
      _closeTimer = null;
    }

    var panel = _root.querySelector(".settings-panel");
    if (panel) panel.classList.remove("is-in");
    _root.classList.remove("is-veiled");

    _root.classList.remove("hidden");
    document.body.classList.add("modal-open");
    // Start in mouse mode; focus the list so arrow keys reach the nav listener.
    _navVisible = false;
    _navRow = null;
    if (_panelsEl) _panelsEl.focus();
    // Set the active-modal root ourselves: this modal bypasses openBlockingModal (own Escape/focus).
    if (typeof setActiveModalRoot === "function") setActiveModalRoot(_root);

    // Next frame: build in the backdrop blur and slide/scale the panel.
    requestAnimationFrame(function () {
      _root.classList.add("is-veiled");
      if (panel) panel.classList.add("is-in");
    });
  }

  function _close() {
    if (!_root || _root.classList.contains("hidden")) return;
    // Release the root now; showHints() keeps hints suppressed through the fade-out.
    if (typeof setActiveModalRoot === "function") setActiveModalRoot(null);
    // A hotkey recording capture-listener must never outlive the modal.
    _hkStopRecording();
    // The color popover lives on document.body and would outlive the modal.
    if (window.ClipgenColorPicker) window.ClipgenColorPicker.close();
    // Drop the keyboard cursor so a re-open starts in mouse mode again.
    _navVisible = false;
    _selectNavRow(null, false);
    var panel = _root.querySelector(".settings-panel");
    if (panel) panel.classList.remove("is-in");
    _root.classList.remove("is-veiled");
    if (_closeTimer) clearTimeout(_closeTimer);
    _closeTimer = setTimeout(function () {
      if (_root) _root.classList.add("hidden");
      // Keep the topnav covered until the veil has faded.
      document.body.classList.remove("modal-open");
      _closeTimer = null;
    }, EXIT_MS);
  }

  function _load() {
    // Inner span, not the container: .cg-shimmer styles would leak onto the settings UI.
    _panelsEl.textContent = "";
    _panelsEl.appendChild(el("span", "cg-shimmer", "Loading settings\u2026"));
    // Refetch the card list on every open so external uploads show up.
    if (window.ClipgenColorPicker) window.ClipgenColorPicker.close();
    _cardsCache = null;
    _cardsCachePromise = null;
    apiGet(_getApiRoot() + "/settings")
      .then(function (data) {
        if (!data.ok) {
          _panelsEl.textContent = "Failed to load settings.";
          return;
        }
        _settings = data.settings;
        _syncRevealBtn(data.desktop, data.path);
        _render();
      })
      .catch(function () {
        _panelsEl.textContent = "Failed to load settings.";
      });
  }

  // Native windows have no address bar. Gate on GUI_LAUNCH: data-desktop-chrome is macOS-only.
  function _syncRevealBtn(isDesktop, path) {
    if (!isDesktop || !_actionsEl) return;
    if (!_revealBtn) {
      _revealBtn = el("button", "btn btn-small");
      _revealBtn.type = "button";
      _revealBtn.textContent = "Show settings file";
      _revealBtn.addEventListener("click", _reveal);
      _actionsEl.appendChild(_revealBtn);
    }
    if (path) _revealBtn.title = path;
  }

  function _reveal() {
    apiPost(_getApiRoot() + "/settings/reveal", {})
      .catch(function (err) {
        _setStatus((err && err.message) || "Could not open the folder");
      });
  }

  function _findSetting(name) {
    if (!_settings) return null;
    for (var i = 0; i < _settings.length; i++) {
      if (_settings[i].name === name) return _settings[i];
    }
    return null;
  }

  function _isChanged(s) {
    // Object-valued settings (mark_categories, hotkeys) need structural comparison.
    if (s.value !== null && typeof s.value === "object") {
      return JSON.stringify(s.value) !== JSON.stringify(s.default);
    }
    return s.value !== s.default;
  }

  function _updateChanged(name) {
    var s = _findSetting(name);
    if (!s) return;
    var row = _panelsEl.querySelector('.settings-row[data-setting="' + name + '"]');
    if (!row) return;
    if (_isChanged(s)) row.classList.add("settings-changed");
    else row.classList.remove("settings-changed");
  }

  function _scheduleSave() {
    if (_saveTimer) clearTimeout(_saveTimer);
    _saveTimer = setTimeout(_save, 400);
  }

  function _save() {
    if (!_settings) return;
    var payload = {};
    for (var i = 0; i < _settings.length; i++) {
      var s = _settings[i];
      payload[s.name] = s.value;
    }
    _setStatus("Saving\u2026", true);

    // Manual fetch, not apiPut: keep data.error from non-2xx responses and r.ok.
    fetch(_getApiRoot() + "/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ settings: payload }),
    })
      .then(function (r) {
        return r.json().then(
          function (j) { return { ok: r.ok, body: j }; },
          function () { return { ok: r.ok, body: null }; }
        );
      })
      .then(function (res) {
        var data = res.body;
        if (!res.ok || !data || !data.ok) {
          _setStatus(data && data.error ? "Save failed: " + data.error : "Save failed");
          return;
        }
        _setStatus("Saved");
        setTimeout(function () { _setStatus(""); }, 2000);
        // Save already persisted; isolate the hook so its errors don't read as "Save failed".
        if (typeof _opts.onSave === "function") {
          try {
            _opts.onSave(data.applied || {}, _settings.slice());
          } catch (err) {
            if (window.console && console.error) {
              console.error("settings onSave hook failed:", err);
            }
          }
        }
      })
      .catch(function () {
        _setStatus("Save failed");
      });
  }

  function _resetTab(tabName) {
    apiPut(_getApiRoot() + "/settings", { reset: "tab:" + tabName })
      .then(function (data) {
        if (!data.ok) {
          _setStatus("Reset failed");
          return;
        }
        _setStatus("Reset " + tabName);
        setTimeout(function () { _setStatus(""); }, 2000);
        _reloadAfterReset(tabName);
      })
      .catch(function () {
        _setStatus("Reset failed");
      });
  }

  function _resetAll() {
    apiPut(_getApiRoot() + "/settings", { reset: "all" })
      .then(function (data) {
        if (!data.ok) {
          _setStatus("Reset failed");
          return;
        }
        _setStatus("Reset to defaults");
        setTimeout(function () { _setStatus(""); }, 2000);
        _reloadAfterReset("all");
      })
      .catch(function () {
        _setStatus("Reset failed");
      });
  }

  function _reloadAfterReset(scope) {
    apiGet(_getApiRoot() + "/settings")
      .then(function (data) {
        if (!data.ok) return;
        _settings = data.settings;
        _render();
        if (typeof _opts.onReset === "function") {
          _opts.onReset(scope, _settings.slice());
        }
      })
      .catch(function () {});
  }

  function _buildRow(s) {
    var row = el("div", "settings-row");
    if (_isChanged(s)) row.classList.add("settings-changed");
    row.setAttribute("data-setting", s.name);

    var labelDiv = el("div", "settings-label");
    var friendlyName = s.name
      .replace(/_/g, " ").toLowerCase()
      .replace(/\b\w/g, function (c) { return c.toUpperCase(); })
      .replace(/Mb$/i, "(MB)").replace(/Seconds$/i, "(s)")
      // Title-casing lowercases acronyms ("Ffmpeg", "Llm"); restore them.
      .replace(/^Ffmpeg\b/, "FFmpeg")
      .replace(/\bLlm\b/g, "LLM")
      .replace(/\bGif\b/g, "GIF")
      .replace(/\bUrl\b/g, "URL");
    labelDiv.appendChild(el("div", "settings-label-name", friendlyName));
    labelDiv.appendChild(el("div", "settings-label-desc", s.description));

    var controlDiv = el("div", "settings-control");
    var settingName = s.name;

    if (s.type === "bool") {
      var toggle = document.createElement("input");
      toggle.type = "checkbox";
      toggle.className = "settings-toggle";
      toggle.checked = !!s.value;
      toggle.addEventListener("change", function () {
        var setting = _findSetting(settingName);
        if (setting) setting.value = this.checked;
        _updateChanged(settingName);
        _scheduleSave();
      });
      controlDiv.appendChild(toggle);
    } else if (s.type === "select" && s.options) {
      var sel = document.createElement("select");
      for (var oi = 0; oi < s.options.length; oi++) {
        var opt = document.createElement("option");
        opt.value = s.options[oi];
        opt.textContent = s.options[oi];
        if (s.options[oi] === s.value) opt.selected = true;
        sel.appendChild(opt);
      }
      sel.addEventListener("change", function () {
        var setting = _findSetting(settingName);
        if (setting) setting.value = this.value;
        _updateChanged(settingName);
        _scheduleSave();
      });
      controlDiv.appendChild(sel);
    } else if (s.type === "model_select") {
      var msel = document.createElement("select");
      msel.className = "settings-model-dropdown";
      var curOpt = document.createElement("option");
      curOpt.value = s.value;
      // Blank value shows the inherit label; avoids an empty placeholder before models load.
      curOpt.textContent = (!s.value && s.emptyLabel) ? s.emptyLabel : s.value;
      curOpt.selected = true;
      msel.appendChild(curOpt);
      msel.disabled = true;
      msel.addEventListener("change", function () {
        var setting = _findSetting(settingName);
        if (setting) setting.value = this.value;
        _updateChanged(settingName);
        _scheduleSave();
      });
      controlDiv.appendChild(msel);
      _loadModelsForSelect(msel, s.provider, s.value, s.emptyLabel);
    } else if (s.type === "str") {
      var txtInput = document.createElement("input");
      txtInput.type = "text";
      txtInput.autocomplete = "off";
      txtInput.value = s.value || "";
      txtInput.placeholder = String(s.default || "");
      txtInput.addEventListener("change", function () {
        var setting = _findSetting(settingName);
        if (setting) setting.value = this.value;
        _updateChanged(settingName);
        _scheduleSave();
      });
      controlDiv.appendChild(txtInput);
    } else if (s.type === "float") {
      var fInput = document.createElement("input");
      fInput.type = "number";
      if (s.min !== undefined && s.min !== null) fInput.min = s.min;
      if (s.max !== undefined && s.max !== null) fInput.max = s.max;
      fInput.step = s.step !== undefined && s.step !== null ? s.step : "any";
      fInput.value = s.value;
      fInput.placeholder = String(s.default);
      fInput.addEventListener("change", function () {
        var setting = _findSetting(settingName);
        if (setting) {
          var n = parseFloat(this.value);
          if (isNaN(n)) n = setting.default;
          if (setting.min !== undefined && setting.min !== null && n < setting.min) n = setting.min;
          if (setting.max !== undefined && setting.max !== null && n > setting.max) n = setting.max;
          setting.value = n;
          this.value = n;
        }
        _updateChanged(settingName);
        _scheduleSave();
      });
      controlDiv.appendChild(fInput);
    } else if (s.type === "mark_categories") {
      row.classList.add("settings-row-stacked");
      _renderMarkCategoriesEditor(controlDiv, settingName);
    } else if (s.type === "hotkeys") {
      row.classList.add("settings-row-stacked");
      _renderHotkeysEditor(controlDiv, settingName);
    } else if (s.type === "card_picker") {
      row.classList.add("settings-row-stacked");
      var kind = s.kind === "end" ? "end" : "title";
      _cardPickers.push({ container: controlDiv, settingName: settingName, kind: kind });
      _renderCardPicker(controlDiv, settingName, kind);
    } else if (s.type === "prompt") {
      row.classList.add("settings-row-stacked");
      var phList = s.placeholders || [];
      var hint = el("div", "settings-prompt-placeholders");
      if (phList.length) {
        hint.appendChild(el("span", "settings-prompt-hint-label", "Placeholders:"));
        for (var pi = 0; pi < phList.length; pi++) {
          hint.appendChild(el("code", "settings-prompt-chip", "{" + phList[pi] + "}"));
        }
      } else {
        hint.appendChild(
          el("span", "settings-prompt-hint-label", "Sent verbatim (no placeholders)."),
        );
      }
      controlDiv.appendChild(hint);

      var ta = document.createElement("textarea");
      ta.className = "settings-prompt-textarea";
      ta.autocomplete = "off";
      ta.spellcheck = false;
      ta.rows = 8;
      ta.value = s.value || "";
      ta.addEventListener("change", function () {
        var setting = _findSetting(settingName);
        if (setting) setting.value = this.value;
        _updateChanged(settingName);
        _scheduleSave();
      });
      controlDiv.appendChild(ta);

      var resetBtn = el("button", "btn btn-small settings-prompt-reset", "Reset to default");
      resetBtn.type = "button";
      resetBtn.addEventListener("click", function () {
        var setting = _findSetting(settingName);
        if (!setting) return;
        setting.value = setting.default;
        ta.value = setting.default || "";
        _updateChanged(settingName);
        _scheduleSave();
      });
      controlDiv.appendChild(resetBtn);
    } else {
      var input = document.createElement("input");
      input.type = "number";
      if (s.min !== undefined && s.min !== null) input.min = s.min;
      if (s.max !== undefined && s.max !== null) input.max = s.max;
      if (s.step !== undefined && s.step !== null) input.step = s.step;
      input.value = s.value;
      input.placeholder = String(s.default);
      input.addEventListener("change", function () {
        var setting = _findSetting(settingName);
        if (setting) setting.value = parseInt(this.value, 10) || 0;
        _updateChanged(settingName);
        _scheduleSave();
      });
      controlDiv.appendChild(input);
    }

    row.appendChild(labelDiv);
    row.appendChild(controlDiv);
    return row;
  }

  // ---- Hotkeys editor ----
  // Value holds only overrides {actionId: "combo"}; "" means disabled.

  var _hkRecordCleanup = null;

  function _hkStopRecording() {
    if (_hkRecordCleanup) {
      _hkRecordCleanup();
      _hkRecordCleanup = null;
    }
  }

  function _hkCommit(container, settingName, action, combo) {
    var s = _findSetting(settingName);
    if (!s) return;
    if (!s.value || typeof s.value !== "object") s.value = {};
    if (combo === (action.combos || []).join(" ")) delete s.value[action.id];
    else s.value[action.id] = combo;
    // Apply live for this page; other pages read config on next load.
    window.ClipgenHotkeys.applyOverrides(s.value);
    CLIPGEN_CONFIG.hotkeyOverrides = s.value;
    _updateChanged(settingName);
    _scheduleSave();
    _renderHotkeysEditor(container, settingName);
  }

  function _hkStartRecording(container, settingName, action, chipsEl) {
    _hkStopRecording();
    chipsEl.classList.add("hotkey-recording");
    chipsEl.innerHTML = "";
    chipsEl.appendChild(el("span", "hotkey-record-hint", "Press keys… Esc cancels · ⌫ disables"));
    var conflictEl = null;

    function cleanup() {
      document.removeEventListener("keydown", onRec, true);
      chipsEl.classList.remove("hotkey-recording");
      if (conflictEl && conflictEl.parentNode) conflictEl.parentNode.removeChild(conflictEl);
    }

    function showConflict(combo, conflicts) {
      // One decision at a time: swap the hint for a Replace/Cancel prompt.
      document.removeEventListener("keydown", onRec, true);
      chipsEl.innerHTML = "";
      var names = [];
      for (var i = 0; i < conflicts.length; i++) names.push(conflicts[i].label);
      conflictEl = el("span", "hotkey-conflict");
      conflictEl.appendChild(el("span", "", "Already used by " + names.join(", ") + "."));
      var replaceBtn = el("button", "btn btn-small", "Replace");
      replaceBtn.type = "button";
      replaceBtn.addEventListener("click", function () {
        var s = _findSetting(settingName);
        if (!s) return;
        if (!s.value || typeof s.value !== "object") s.value = {};
        // Never leave two live bindings: disable the previous owners.
        for (var c = 0; c < conflicts.length; c++) s.value[conflicts[c].id] = "";
        _hkRecordCleanup = null;
        cleanup();
        _hkCommit(container, settingName, action, combo);
      });
      var cancelBtn = el("button", "btn btn-small", "Cancel");
      cancelBtn.type = "button";
      cancelBtn.addEventListener("click", function () {
        _hkRecordCleanup = null;
        cleanup();
        _renderHotkeysEditor(container, settingName);
      });
      conflictEl.appendChild(replaceBtn);
      conflictEl.appendChild(cancelBtn);
      chipsEl.appendChild(conflictEl);
    }

    function onRec(e) {
      e.preventDefault();
      e.stopPropagation();
      if (e.key === "Escape") {
        _hkRecordCleanup = null;
        cleanup();
        _renderHotkeysEditor(container, settingName);
        return;
      }
      if (e.key === "Backspace" && !e.metaKey && !e.ctrlKey && !e.altKey && !e.shiftKey) {
        _hkRecordCleanup = null;
        cleanup();
        _hkCommit(container, settingName, action, "");
        return;
      }
      var combo = window.ClipgenHotkeys.normalizeEvent(e);
      if (!combo) return; // modifier-only keydown; keep waiting
      var conflicts = window.ClipgenHotkeys.comboConflicts(combo, action.id);
      if (conflicts.length) {
        showConflict(combo, conflicts);
        return;
      }
      _hkRecordCleanup = null;
      cleanup();
      _hkCommit(container, settingName, action, combo);
    }

    document.addEventListener("keydown", onRec, true);
    _hkRecordCleanup = cleanup;
  }

  function _renderHotkeysEditor(container, settingName) {
    _hkStopRecording();
    container.innerHTML = "";
    var setting = _findSetting(settingName);
    if (!setting) return;
    if (!window.ClipgenHotkeys) {
      container.appendChild(
        el("div", "settings-label-desc", "Hotkey catalog unavailable on this page."),
      );
      return;
    }
    if (!setting.value || typeof setting.value !== "object") setting.value = {};
    // Keep the live registry in sync with the editor (also after resets).
    window.ClipgenHotkeys.applyOverrides(setting.value);
    CLIPGEN_CONFIG.hotkeyOverrides = setting.value;

    var cat = window.ClipgenHotkeys.catalog();
    var editor = el("div", "hotkey-editor");
    for (var si = 0; si < cat.sections.length; si++) {
      var section = cat.sections[si];
      var rows = [];
      for (var ai = 0; ai < cat.actions.length; ai++) {
        var a = cat.actions[ai];
        if (a.section === section.id && !a.note) rows.push(a);
      }
      if (!rows.length) continue;
      editor.appendChild(el("div", "settings-group-label", section.label));
      for (var ri = 0; ri < rows.length; ri++) {
        editor.appendChild(_hkBuildActionRow(container, settingName, setting, rows[ri]));
      }
    }
    container.appendChild(editor);
  }

  function _hkBuildActionRow(container, settingName, setting, action) {
    var rowEl = el("div", "hotkey-row");
    var overridden = Object.prototype.hasOwnProperty.call(setting.value, action.id);
    if (overridden) rowEl.classList.add("hotkey-row-changed");

    var chips = el("span", "hotkey-chips");
    var combos = window.ClipgenHotkeys.resolvedCombos(action.id);
    if (action.rebindable === false) {
      chips.classList.add("hotkey-fixed");
      chips.appendChild(el("kbd", "hotkey-chip", action.displayKeys || (action.combos || []).join(" ")));
      chips.title = "This shortcut cannot be rebound.";
    } else if (!combos.length) {
      var offChip = el("span", "hotkey-chip hotkey-chip-disabled", "Disabled");
      chips.appendChild(offChip);
    } else {
      for (var ci = 0; ci < combos.length; ci++) {
        chips.appendChild(window.ClipgenHotkeys.fillKeycap(el("kbd", "hotkey-chip"), combos[ci]));
      }
    }
    if (action.rebindable !== false) {
      chips.title = "Click to rebind";
      chips.setAttribute("role", "button");
      chips.tabIndex = 0;
      chips.addEventListener("click", function () {
        _hkStartRecording(container, settingName, action, chips);
      });
    }
    rowEl.appendChild(chips);
    rowEl.appendChild(el("span", "hotkey-label", action.label));

    if (overridden && action.rebindable !== false) {
      var resetBtn = el("button", "btn btn-small hotkey-reset", "Reset");
      resetBtn.type = "button";
      resetBtn.addEventListener("click", function () {
        var s = _findSetting(settingName);
        if (!s || !s.value) return;
        delete s.value[action.id];
        window.ClipgenHotkeys.applyOverrides(s.value);
        CLIPGEN_CONFIG.hotkeyOverrides = s.value;
        _updateChanged(settingName);
        _scheduleSave();
        _renderHotkeysEditor(container, settingName);
      });
      rowEl.appendChild(resetBtn);
    }
    return rowEl;
  }

  function _slugifyKey(label, existingKeys) {
    var base = String(label || "").toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
    if (!base) base = "category";
    var key = base;
    var n = 2;
    while (existingKeys.indexOf(key) !== -1) {
      key = base + "_" + n;
      n++;
    }
    return key;
  }

  function _renderMarkCategoriesEditor(container, settingName) {
    container.innerHTML = "";
    var setting = _findSetting(settingName);
    if (!setting) return;
    var value = setting.value && typeof setting.value === "object" ? setting.value : {};

    var editor = el("div", "mark-cat-editor");
    var keys = Object.keys(value);
    for (var i = 0; i < keys.length; i++) {
      (function (key) {
        var entry = value[key] || { label: "", color: "#888888" };
        var rowEl = el("div", "mark-cat-row");

        var keyEl = el("span", "mark-cat-key", key);
        keyEl.title = key;
        rowEl.appendChild(keyEl);

        var labelInput = document.createElement("input");
        labelInput.type = "text";
        labelInput.className = "mark-cat-label";
        labelInput.autocomplete = "off";
        labelInput.value = entry.label || "";
        labelInput.placeholder = "Label";
        labelInput.addEventListener("change", function () {
          var s = _findSetting(settingName);
          if (!s || !s.value || !s.value[key]) return;
          var v = this.value.trim() || key;
          s.value[key].label = v;
          this.value = v;
          _updateChanged(settingName);
          _scheduleSave();
        });
        rowEl.appendChild(labelInput);

        var colorInput = document.createElement("input");
        colorInput.type = "color";
        colorInput.className = "mark-cat-color";
        colorInput.value = entry.color || "#888888";
        colorInput.addEventListener("change", function () {
          var s = _findSetting(settingName);
          if (!s || !s.value || !s.value[key]) return;
          s.value[key].color = this.value;
          _updateChanged(settingName);
          _scheduleSave();
        });
        rowEl.appendChild(colorInput);

        var delBtn = el("button", "btn btn-small mark-cat-delete", "Remove");
        delBtn.type = "button";
        delBtn.addEventListener("click", function () {
          var s = _findSetting(settingName);
          if (!s || !s.value) return;
          delete s.value[key];
          _updateChanged(settingName);
          _renderMarkCategoriesEditor(container, settingName);
          _scheduleSave();
        });
        rowEl.appendChild(delBtn);

        editor.appendChild(rowEl);
      })(keys[i]);
    }

    var addRow = el("div", "mark-cat-add-row");
    var addLabel = document.createElement("input");
    addLabel.type = "text";
    addLabel.className = "mark-cat-label";
    addLabel.autocomplete = "off";
    addLabel.placeholder = "New category label";
    addRow.appendChild(addLabel);

    var addColor = document.createElement("input");
    addColor.type = "color";
    addColor.className = "mark-cat-color";
    addColor.value = "#888888";
    addRow.appendChild(addColor);

    var addBtn = el("button", "btn btn-small", "Add");
    addBtn.type = "button";
    addBtn.addEventListener("click", function () {
      var s = _findSetting(settingName);
      if (!s) return;
      var label = addLabel.value.trim();
      if (!label) {
        addLabel.focus();
        return;
      }
      if (!s.value || typeof s.value !== "object") s.value = {};
      var newKey = _slugifyKey(label, Object.keys(s.value));
      s.value[newKey] = { label: label, color: addColor.value };
      _updateChanged(settingName);
      _renderMarkCategoriesEditor(container, settingName);
      _scheduleSave();
    });
    addRow.appendChild(addBtn);

    editor.appendChild(addRow);
    container.appendChild(editor);
  }

  // ── Titlecard / endcard background picker ──────────────────────────────

  function _fetchCards(force) {
    if (force) {
      _cardsCache = null;
      _cardsCachePromise = null;
    }
    if (_cardsCache) return Promise.resolve(_cardsCache);
    if (_cardsCachePromise) return _cardsCachePromise;
    _cardsCachePromise = apiGet(_getApiRoot() + "/titlecards")
      .then(function (data) {
        _cardsCachePromise = null;
        if (data && data.ok) _cardsCache = data;
        return data;
      })
      .catch(function () {
        _cardsCachePromise = null;
        return null;
      });
    return _cardsCachePromise;
  }

  function _refreshAllCardPickers() {
    _fetchCards(true).then(function () {
      for (var i = 0; i < _cardPickers.length; i++) {
        var p = _cardPickers[i];
        if (document.body.contains(p.container)) {
          _renderCardPicker(p.container, p.settingName, p.kind);
        }
      }
    });
  }

  function _cardColorSettingName(kind) {
    return kind === "end" ? "ENDCARD_COLOR" : "TITLECARD_COLOR";
  }

  function _cardCurrentColor(kind) {
    var s = _findSetting(_cardColorSettingName(kind));
    return s && s.value ? s.value : "#000000";
  }

  function _cardTile(item, settingName, kind, selectedId) {
    var tile = el("div", "card-tile");
    tile.setAttribute("data-card-id", item.id);
    if (item.id === selectedId) tile.classList.add("is-selected");

    var preview = el("div", "card-tile-preview card-tile-preview--" + item.kind);
    if (item.url) {
      var img = document.createElement("img");
      img.decoding = "async";
      img.src = item.url;
      img.alt = item.label;
      img.loading = "lazy";
      preview.appendChild(img);
    } else if (item.kind === "none") {
      preview.appendChild(el("span", "card-tile-placeholder", "No endcard"));
    } else if (item.kind === "color") {
      preview.style.background = _cardCurrentColor(kind);
    }
    // Approximate the ffmpeg drawtext overlay; endcards carry no text.
    if (kind === "title" && item.kind !== "none") {
      preview.appendChild(el("span", "card-tile-text-overlay", _SAMPLE_TITLE_TEXT));
    }
    tile.appendChild(preview);

    function selectTile() {
      var s = _findSetting(settingName);
      if (!s) return;
      s.value = item.id;
      _updateChanged(settingName);
      var siblings = tile.parentNode.querySelectorAll(".card-tile");
      for (var t = 0; t < siblings.length; t++) {
        siblings[t].classList.remove("is-selected");
      }
      tile.classList.add("is-selected");
      _scheduleSave();
    }
    tile.addEventListener("click", selectTile);

    if (item.kind === "color") {
      // Label row carries an inline swatch that opens the color picker.
      var labelRow = el("div", "card-tile-label-row");
      labelRow.appendChild(el("span", "card-tile-label", item.label));
      var box = el("button", "card-tile-color-box");
      box.type = "button";
      box.title = "Pick color";
      box.setAttribute("aria-label", "Pick solid color");
      box.style.background = _cardCurrentColor(kind);
      box.addEventListener("click", function (ev) {
        ev.stopPropagation();
        selectTile();
        _openCardColorPicker(box, preview, kind);
      });
      labelRow.appendChild(box);
      tile.appendChild(labelRow);
    } else {
      tile.appendChild(el("span", "card-tile-label", item.label));
    }

    if (item.deletable) {
      var del = el("button", "card-tile-delete", "×");
      del.type = "button";
      del.title = "Delete";
      del.setAttribute("aria-label", "Delete " + item.label);
      del.addEventListener("click", function (ev) {
        ev.stopPropagation();
        _deleteCard(item.id);
      });
      tile.appendChild(del);
    }
    return tile;
  }

  function _openCardColorPicker(box, preview, kind) {
    if (!window.ClipgenColorPicker) return;
    var colorSettingName = _cardColorSettingName(kind);
    window.ClipgenColorPicker.open({
      anchor: box,
      value: _cardCurrentColor(kind),
      onInput: function (hex) {
        box.style.background = hex;
        preview.style.background = hex;
        var s = _findSetting(colorSettingName);
        if (s) s.value = hex;
      },
      onChange: function (hex) {
        var s = _findSetting(colorSettingName);
        if (s) {
          s.value = hex;
          _updateChanged(colorSettingName);
        }
        _scheduleSave();
      },
    });
  }

  function _cardUploadTile(settingName) {
    var tile = el("div", "card-tile card-tile--upload");
    tile.appendChild(el("span", "card-tile-upload-icon", "+"));
    tile.appendChild(el("span", "card-tile-label", "Upload"));

    var input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,image/jpeg,image/webp";
    input.style.display = "none";
    input.addEventListener("change", function () {
      var file = input.files && input.files[0];
      if (file) _uploadCard(file, settingName);
      input.value = "";
    });
    tile.appendChild(input);
    tile.addEventListener("click", function () { input.click(); });
    return tile;
  }

  function _renderCardPicker(container, settingName, kind) {
    container.innerHTML = "";
    var grid = el("div", "settings-card-picker");
    grid.appendChild(el("div", "card-tile card-tile--loading cg-shimmer", "Loading…"));
    container.appendChild(grid);

    _fetchCards(false).then(function (data) {
      grid.innerHTML = "";
      if (!data || !data.ok || !data[kind]) {
        grid.appendChild(el("div", "card-tile card-tile--loading", "Failed to load"));
        return;
      }
      var setting = _findSetting(settingName);
      var selectedId = setting ? setting.value : data[kind].selected;
      var items = data[kind].items || [];
      var frag = document.createDocumentFragment();
      for (var i = 0; i < items.length; i++) {
        frag.appendChild(_cardTile(items[i], settingName, kind, selectedId));
      }
      frag.appendChild(_cardUploadTile(settingName));
      grid.appendChild(frag);
    });
  }

  function _uploadCard(file, settingName) {
    _setStatus("Uploading…", true);
    var form = new FormData();
    form.append("file", file);
    // Manual fetch, not apiPost: FormData upload; keep data.error from non-2xx and r.ok.
    fetch(_getApiRoot() + "/titlecards/upload", { method: "POST", body: form })
      .then(function (r) {
        return r.json().then(
          function (j) { return { ok: r.ok, body: j }; },
          function () { return { ok: r.ok, body: null }; }
        );
      })
      .then(function (res) {
        var data = res.body;
        if (!res.ok || !data || !data.ok) {
          _setStatus(data && data.error ? data.error : "Upload failed");
          return;
        }
        _setStatus("Uploaded");
        // Auto-select the new image for the picker that triggered the upload.
        var s = _findSetting(settingName);
        if (s && data.item) {
          s.value = data.item.id;
          _updateChanged(settingName);
          _scheduleSave();
        }
        _refreshAllCardPickers();
      })
      .catch(function () {
        _setStatus("Upload failed");
      });
  }

  function _deleteCard(name) {
    _setStatus("Deleting…", true);
    // Manual fetch, not apiDelete: keep data.error from non-2xx responses and r.ok.
    fetch(_getApiRoot() + "/titlecards/image/" + encodeURIComponent(name), {
      method: "DELETE",
    })
      .then(function (r) {
        return r.json().then(
          function (j) { return { ok: r.ok, body: j }; },
          function () { return { ok: r.ok, body: null }; }
        );
      })
      .then(function (res) {
        var data = res.body;
        if (!res.ok || !data || !data.ok) {
          _setStatus(data && data.error ? data.error : "Delete failed");
          return;
        }
        _setStatus("Deleted");
        // Mirror the server's reset of selections pointing at the deleted file.
        if (data.reset) {
          for (var key in data.reset) {
            var s = _findSetting(key);
            if (s) {
              s.value = data.reset[key];
              _updateChanged(key);
            }
          }
        }
        _refreshAllCardPickers();
      })
      .catch(function () {
        _setStatus("Delete failed");
      });
  }

  function _render() {
    _tabsEl.innerHTML = "";
    _panelsEl.innerHTML = "";
    _cardPickers = [];

    // Partition settings by tab, preserving insertion order from the API.
    var byTab = {};
    for (var i = 0; i < _settings.length; i++) {
      var s = _settings[i];
      var tab = s.tab || "General";
      if (!byTab[tab]) byTab[tab] = [];
      byTab[tab].push(s);
    }

    // Only show tabs that have settings; keep the fixed order.
    var tabsToShow = [];
    for (var ti = 0; ti < TAB_ORDER.length; ti++) {
      if (byTab[TAB_ORDER[ti]]) tabsToShow.push(TAB_ORDER[ti]);
    }
    // Append any unknown tabs (forward-compat).
    for (var k in byTab) {
      if (TAB_ORDER.indexOf(k) === -1) tabsToShow.push(k);
    }

    if (tabsToShow.indexOf(_activeTab) === -1 && tabsToShow.length) {
      _activeTab = tabsToShow[0];
    }

    for (var j = 0; j < tabsToShow.length; j++) {
      var name = tabsToShow[j];
      var tabBtn = el("button", "settings-tab", name);
      tabBtn.type = "button";
      tabBtn.setAttribute("role", "tab");
      tabBtn.setAttribute("data-tab", name);
      // First nine tabs get Alt-hold chips; the settings.tab hotkey (1–9) switches.
      if (j < 9) {
        tabBtn.setAttribute("data-hotkey", "settings.tab");
        tabBtn.setAttribute("data-hotkey-combo", String(j));
      }
      if (name === _activeTab) {
        tabBtn.classList.add("is-active");
        tabBtn.setAttribute("aria-selected", "true");
      }
      tabBtn.addEventListener("click", _makeTabClickHandler(name));
      _tabsEl.appendChild(tabBtn);

      var panel = el("div", "settings-tab-panel");
      panel.setAttribute("data-tab", name);
      if (name !== _activeTab) panel.classList.add("hidden");

      var items = byTab[name];
      // Group by sub-heading, preserving order.
      var groups = {};
      var groupOrder = [];
      for (var gi = 0; gi < items.length; gi++) {
        var it = items[gi];
        // Hidden settings ship to the client but get no row (e.g. card colors).
        if (it.type === "hidden") continue;
        var g = it.group || "";
        if (!groups[g]) {
          groups[g] = [];
          groupOrder.push(g);
        }
        groups[g].push(it);
      }
      for (var go = 0; go < groupOrder.length; go++) {
        var gname = groupOrder[go];
        if (gname) panel.appendChild(el("div", "settings-group-label", gname));
        var gitems = groups[gname];
        for (var gii = 0; gii < gitems.length; gii++) {
          panel.appendChild(_buildRow(gitems[gii]));
        }
      }

      if (name === "Summaries") panel.appendChild(_buildLlmModelsBlock());

      var resetTabBtn = el("button", "btn btn-small settings-tab-reset", "Reset this tab");
      resetTabBtn.type = "button";
      resetTabBtn.setAttribute("data-hotkey", "settings.resetTab");
      resetTabBtn.addEventListener("click", _makeTabResetHandler(name));
      panel.appendChild(resetTabBtn);

      _panelsEl.appendChild(panel);
    }
    _resetNavCursor(true);
  }

  function _makeTabClickHandler(name) {
    return function () { _switchTab(name); };
  }

  function _makeTabResetHandler(name) {
    return function () { _resetTab(name); };
  }

  function _switchTab(name) {
    _activeTab = name;
    var tabBtns = _tabsEl.querySelectorAll(".settings-tab");
    for (var i = 0; i < tabBtns.length; i++) {
      var b = tabBtns[i];
      if (b.getAttribute("data-tab") === name) {
        b.classList.add("is-active");
        b.setAttribute("aria-selected", "true");
      } else {
        b.classList.remove("is-active");
        b.removeAttribute("aria-selected");
      }
    }
    var panels = _panelsEl.querySelectorAll(".settings-tab-panel");
    for (var j = 0; j < panels.length; j++) {
      var p = panels[j];
      if (p.getAttribute("data-tab") === name) p.classList.remove("hidden");
      else p.classList.add("hidden");
    }
    // Park the cursor on the new panel's first row.
    _resetNavCursor(true);
  }

  // ---- Keyboard list navigation ----
  // _navRow is position, _navVisible the ring; mouse mode paints none.
  var _navRow = null;
  var _navVisible = false;

  function _navPanel() {
    return _panelsEl ? _panelsEl.querySelector(".settings-tab-panel:not(.hidden)") : null;
  }

  function _navRows() {
    var panel = _navPanel();
    if (!panel) return [];
    return Array.prototype.filter.call(
      panel.querySelectorAll(".settings-row"),
      function (r) { return r.offsetParent !== null; }
    );
  }

  function _selectNavRow(row, focus) {
    var rows = _navRows();
    for (var i = 0; i < rows.length; i++) {
      var on = rows[i] === row;
      // Ring only in keyboard mode; the roving tabindex follows the position regardless.
      rows[i].classList.toggle("is-nav-selected", on && _navVisible);
      rows[i].tabIndex = on ? 0 : -1;
    }
    _navRow = row || null;
    if (row && focus !== false) {
      row.focus();
      if (row.scrollIntoView) row.scrollIntoView({ block: "nearest" });
    }
  }

  // Reveal the cursor before the key acts so a move paints the destination.
  function _showNavCursor() {
    if (_navVisible) return;
    _navVisible = true;
    _selectNavRow(_navRow, false);
  }

  // Park on the first row; mouse mode clears instead so ArrowDown hits row 0.
  function _resetNavCursor(focus) {
    if (!_navVisible) {
      _selectNavRow(null, false);
      return;
    }
    var rows = _navRows();
    _selectNavRow(rows.length ? rows[0] : null, !!focus && _isModalOpen());
  }

  function _moveNav(delta) {
    var rows = _navRows();
    if (!rows.length) return false;
    var idx = _navRow ? rows.indexOf(_navRow) : -1;
    var next = idx === -1 ? (delta < 0 ? rows.length - 1 : 0) : idx + delta;
    if (next < 0 || next >= rows.length) return false; // boundary: fall through
    _selectNavRow(rows[next]);
    return true;
  }

  function _fireChange(node) {
    node.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function _isEditableField(t) {
    if (!t || !t.tagName) return false;
    if (t.tagName === "TEXTAREA" || t.tagName === "SELECT") return true;
    if (t.tagName === "INPUT") {
      var ty = (t.type || "text").toLowerCase();
      return ty !== "checkbox" && ty !== "radio" && ty !== "button" && ty !== "submit";
    }
    return false;
  }

  // Left/Right: toggle, cycle, or step. Fires change so per-control handlers persist.
  function _actuateNav(row, dir) {
    if (!row) return;
    var control = row.querySelector(".settings-control");
    if (!control) return;
    var cb = control.querySelector("input.settings-toggle");
    if (cb) { cb.checked = dir > 0; _fireChange(cb); return; }
    var sel = control.querySelector("select");
    if (sel && !sel.disabled && sel.options.length) {
      var n = sel.options.length;
      sel.selectedIndex = ((sel.selectedIndex + dir) % n + n) % n;
      _fireChange(sel);
      return;
    }
    var num = control.querySelector('input[type="number"]');
    if (num) {
      var step = parseFloat(num.step);
      if (isNaN(step)) step = 1;
      var cur = parseFloat(num.value);
      if (isNaN(cur)) cur = 0;
      num.value = cur + dir * step;
      _fireChange(num); // the float handler re-clamps to min/max
      return;
    }
  }

  // Enter: toggle a bool, else focus the first editable field.
  function _editNav(row) {
    if (!row) return;
    var control = row.querySelector(".settings-control");
    if (!control) return;
    var cb = control.querySelector("input.settings-toggle");
    if (cb) { cb.checked = !cb.checked; _fireChange(cb); return; }
    var field = control.querySelector(
      'input[type="text"], input[type="number"], textarea, select'
    );
    if (field) {
      field.focus();
      if (field.select) { try { field.select(); } catch (e) {} }
    }
  }

  function _navActive() {
    return _isModalOpen() && !_hkRecordCleanup &&
      !(window.ClipgenColorPicker && window.ClipgenColorPicker.isOpen && window.ClipgenColorPicker.isOpen());
  }

  document.addEventListener("keydown", function (e) {
    if (!_navActive()) return;
    var t = e.target;

    // Edit mode: Esc (or Enter on single-line fields) returns to the cursor.
    if (_isEditableField(t) && _panelsEl && _panelsEl.contains(t)) {
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation(); // beat the modal-close (bubble) listener below
        _selectNavRow(t.closest ? t.closest(".settings-row") : _navRow);
      } else if (e.key === "Enter" && t.tagName !== "TEXTAREA") {
        e.preventDefault();
        e.stopPropagation();
        _fireChange(t);
        _selectNavRow(t.closest ? t.closest(".settings-row") : _navRow);
      }
      return;
    }

    // Nav mode: only while focus is within the settings list.
    var inPanel = (_panelsEl && _panelsEl.contains(t)) || (_navRow && t === _navRow);
    if (!inPanel) return;
    // Sync the cursor to wherever focus actually landed (e.g. a clicked toggle).
    var here = t && t.closest ? t.closest(".settings-row") : null;
    if (here && here !== _navRow) _selectNavRow(here, false);

    switch (e.key) {
      case "ArrowDown": e.preventDefault(); _showNavCursor(); _moveNav(1); break;
      case "ArrowUp":   e.preventDefault(); _showNavCursor(); _moveNav(-1); break;
      case "Tab":
        // At the list edges fall through to native Tab so focus can leave.
        _showNavCursor();
        if (_moveNav(e.shiftKey ? -1 : 1)) e.preventDefault();
        break;
      case "ArrowLeft":  e.preventDefault(); _showNavCursor(); _actuateNav(_navRow, -1); break;
      case "ArrowRight": e.preventDefault(); _showNavCursor(); _actuateNav(_navRow, 1); break;
      case "Enter":      e.preventDefault(); _showNavCursor(); _editNav(_navRow); break;
      // Space is left to the browser so it still scrolls the settings list.
      default: break;
    }
  }, true);

  // Escape closes. Bubble-phase so the hotkey recorder and color picker (capture, stopPropagation) win.
  document.addEventListener("keydown", function (e) {
    if (e.key !== "Escape") return;
    if (!_root || _root.classList.contains("hidden")) return;
    if (window.ClipgenColorPicker && window.ClipgenColorPicker.isOpen && window.ClipgenColorPicker.isOpen()) return;
    e.preventDefault();
    _close();
  });

  function _isModalOpen() {
    return !!_root && !_root.classList.contains("hidden");
  }

  // Cycle tabs with wrap via the buttons' click path (keeps aria state right).
  function _cycleTab(delta) {
    if (!_tabsEl) return;
    var btns = _tabsEl.querySelectorAll(".settings-tab");
    if (!btns.length) return;
    var cur = 0;
    for (var i = 0; i < btns.length; i++) {
      if (btns[i].classList.contains("is-active")) { cur = i; break; }
    }
    btns[((cur + delta) % btns.length + btns.length) % btns.length].click();
  }

  // Modal-only tab hotkeys: digits 1–9 jump, Z/X cycle like Screenspace's tool tabs.
  if (window.ClipgenHotkeys) {
    window.ClipgenHotkeys.register([
      {
        id: "settings.tab",
        inModal: true,
        when: _isModalOpen,
        handler: function (e, combo) {
          var n = parseInt(combo, 10);
          if (!_tabsEl || isNaN(n)) return;
          var btns = _tabsEl.querySelectorAll(".settings-tab");
          var btn = btns[n - 1];
          if (btn) btn.click();
        }
      },
      { id: "settings.cyclePrev", inModal: true, when: _isModalOpen, handler: function () { _cycleTab(-1); } },
      { id: "settings.cycleNext", inModal: true, when: _isModalOpen, handler: function () { _cycleTab(1); } },
      // Reset hotkeys reuse the buttons' unconfirmed handlers; Shift+R resets everything.
      { id: "settings.resetTab", inModal: true, when: _isModalOpen, handler: function () { _resetTab(_activeTab); } },
      { id: "settings.resetAll", inModal: true, when: _isModalOpen, handler: function () { _resetAll(); } }
    ]);
  }

  window.openSettingsModal = function (options) {
    _opts = options || {};
    if (_opts.initialTab) _activeTab = _opts.initialTab;
    _open();
    _load();
  };

  // Wire TopNav #settingsBtn to the modal. opts: initialTab, version (string|function), onApply(applied, settings).
  window.wireSettingsButton = function (opts) {
    opts = opts || {};
    var btn = document.getElementById("settingsBtn");
    if (!btn) return;
    btn.addEventListener("click", function () {
      var options = {
        initialTab: opts.initialTab,
        version: typeof opts.version === "function" ? opts.version() : opts.version,
      };
      if (opts.onApply) {
        options.onSave = function (applied, settings) {
          opts.onApply(applied, settings);
        };
        options.onReset = function (_scope, settings) {
          opts.onApply(null, settings);
        };
      }
      window.openSettingsModal(options);
    });
  };
})();
