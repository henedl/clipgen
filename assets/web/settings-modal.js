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
    "Video & Clips",
    "Screenspace",
    "Transcription",
    "Summaries",
    "CLI",
  ];

  // Entry/exit animation tuning. Mirrors the Start overlay's vocabulary
  // (--host-blur / --veil-alpha on the overlay root, .is-in on the panel)
  // but dialled lighter — Settings is a sheet, not a full-screen launcher.
  var INTRO_BLUR_PX = 12;
  var INTRO_VEIL_ALPHA = 0.45;
  var EXIT_MS = 360;

  var _root = null;
  var _tabsEl = null;
  var _panelsEl = null;
  var _statusEl = null;
  var _settings = null;
  var _activeTab = "General";
  var _saveTimer = null;
  var _opts = {};
  var _modelsCache = null;
  var _modelsCachePromise = null;
  var _closeTimer = null;

  function _getApiRoot() {
    // Each page is served under a different prefix (/studio/, /transcripts/,
    // /screenspace/). Settings + models are registered at the combined-app
    // root, so request them from an absolute path.
    return "/api";
  }

  function _formatSize(mb) {
    if (mb >= 1024) return (mb / 1024).toFixed(1) + " GB";
    return mb + " MB";
  }

  function _fetchModels() {
    if (_modelsCache) return Promise.resolve(_modelsCache);
    if (_modelsCachePromise) return _modelsCachePromise;
    _modelsCachePromise = fetch(_getApiRoot() + "/models")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (data && data.ok) _modelsCache = data;
        return data;
      })
      .catch(function () { return null; });
    return _modelsCachePromise;
  }

  function _loadModelsForSelect(sel, provider, currentValue) {
    _fetchModels().then(function (data) {
      if (!data || !data.ok) { sel.disabled = false; return; }

      var models = [];
      if (provider === "whisper") {
        models = (data.whisper && data.whisper.models) || [];
      } else if (provider === "ollama") {
        models = (data.ollama && data.ollama.models) || [];
      }

      sel.innerHTML = "";
      var hasCurrentValue = false;
      for (var i = 0; i < models.length; i++) {
        var m = models[i];
        var opt = document.createElement("option");
        opt.value = m.name;
        var label = m.name;
        if (m.size_mb) label += " (" + _formatSize(m.size_mb) + ")";
        if (m.parameter_size) label += " \u00B7 " + m.parameter_size;
        if (m.description) label += " \u2014 " + m.description;
        opt.textContent = label;
        if (m.name === currentValue) {
          opt.selected = true;
          hasCurrentValue = true;
        }
        sel.appendChild(opt);
      }
      if (!hasCurrentValue && currentValue) {
        var custom = document.createElement("option");
        custom.value = currentValue;
        custom.textContent = currentValue + " (current)";
        custom.selected = true;
        sel.insertBefore(custom, sel.firstChild);
      }
      sel.disabled = false;
    });
  }

  function _buildDom() {
    if (_root) return;

    var overlay = el("div", "settings-overlay hidden");
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
    panel.appendChild(panels);

    var footer = el("div", "settings-footer");
    var resetAll = el("button", "btn btn-small settings-reset-all");
    resetAll.type = "button";
    resetAll.textContent = "Reset all to defaults";
    resetAll.addEventListener("click", function () { _resetAll(); });
    var status = el("span", "settings-save-status");
    footer.appendChild(resetAll);
    footer.appendChild(status);
    panel.appendChild(footer);

    overlay.addEventListener("click", function (e) {
      if (e.target === overlay) _close();
    });

    document.body.appendChild(overlay);

    _root = overlay;
    _tabsEl = tabs;
    _panelsEl = panels;
    _statusEl = status;
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
    _root.style.setProperty("--host-blur", "0px");
    _root.style.setProperty("--veil-alpha", "0");

    _root.classList.remove("hidden");

    // Next frame: build in the backdrop blur and slide/scale the panel.
    requestAnimationFrame(function () {
      _root.style.setProperty("--host-blur", INTRO_BLUR_PX + "px");
      _root.style.setProperty("--veil-alpha", String(INTRO_VEIL_ALPHA));
      if (panel) panel.classList.add("is-in");
    });
  }

  function _close() {
    if (!_root || _root.classList.contains("hidden")) return;
    var panel = _root.querySelector(".settings-panel");
    if (panel) panel.classList.remove("is-in");
    _root.style.setProperty("--host-blur", "0px");
    _root.style.setProperty("--veil-alpha", "0");
    if (_closeTimer) clearTimeout(_closeTimer);
    _closeTimer = setTimeout(function () {
      if (_root) _root.classList.add("hidden");
      _closeTimer = null;
    }, EXIT_MS);
  }

  function _load() {
    _panelsEl.textContent = "Loading settings\u2026";
    fetch(_getApiRoot() + "/settings")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data.ok) {
          _panelsEl.textContent = "Failed to load settings.";
          return;
        }
        _settings = data.settings;
        _render();
      })
      .catch(function () {
        _panelsEl.textContent = "Failed to load settings.";
      });
  }

  function _findSetting(name) {
    if (!_settings) return null;
    for (var i = 0; i < _settings.length; i++) {
      if (_settings[i].name === name) return _settings[i];
    }
    return null;
  }

  function _updateChanged(name) {
    var s = _findSetting(name);
    if (!s) return;
    var row = _panelsEl.querySelector('.settings-row[data-setting="' + name + '"]');
    if (!row) return;
    if (s.value !== s.default) row.classList.add("settings-changed");
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
    if (_statusEl) _statusEl.textContent = "Saving\u2026";

    fetch(_getApiRoot() + "/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ settings: payload }),
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (_statusEl) {
          if (data.ok) {
            _statusEl.textContent = "Saved";
            setTimeout(function () { if (_statusEl) _statusEl.textContent = ""; }, 2000);
          } else {
            _statusEl.textContent = data.error ? "Save failed: " + data.error : "Save failed";
          }
        }
        if (data.ok && typeof _opts.onSave === "function") {
          _opts.onSave(data.applied || {}, _settings.slice());
        }
      })
      .catch(function () {
        if (_statusEl) _statusEl.textContent = "Save failed";
      });
  }

  function _resetTab(tabName) {
    fetch(_getApiRoot() + "/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reset: "tab:" + tabName }),
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data.ok) {
          if (_statusEl) _statusEl.textContent = "Reset failed";
          return;
        }
        if (_statusEl) {
          _statusEl.textContent = "Reset " + tabName;
          setTimeout(function () { if (_statusEl) _statusEl.textContent = ""; }, 2000);
        }
        _reloadAfterReset(tabName);
      })
      .catch(function () {
        if (_statusEl) _statusEl.textContent = "Reset failed";
      });
  }

  function _resetAll() {
    fetch(_getApiRoot() + "/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reset: "all" }),
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data.ok) {
          if (_statusEl) _statusEl.textContent = "Reset failed";
          return;
        }
        if (_statusEl) {
          _statusEl.textContent = "Reset to defaults";
          setTimeout(function () { if (_statusEl) _statusEl.textContent = ""; }, 2000);
        }
        _reloadAfterReset("all");
      })
      .catch(function () {
        if (_statusEl) _statusEl.textContent = "Reset failed";
      });
  }

  function _reloadAfterReset(scope) {
    fetch(_getApiRoot() + "/settings")
      .then(function (r) { return r.json(); })
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
    if (s.value !== s.default) row.classList.add("settings-changed");
    row.setAttribute("data-setting", s.name);

    var labelDiv = el("div", "settings-label");
    var friendlyName = s.name
      .replace(/_/g, " ").toLowerCase()
      .replace(/\b\w/g, function (c) { return c.toUpperCase(); })
      .replace(/Mb$/i, "(MB)").replace(/Seconds$/i, "(s)");
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
      curOpt.textContent = s.value;
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
      _loadModelsForSelect(msel, s.provider, s.value);
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

  function _render() {
    _tabsEl.innerHTML = "";
    _panelsEl.innerHTML = "";

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

      var resetTabBtn = el("button", "btn btn-small settings-tab-reset", "Reset this tab");
      resetTabBtn.type = "button";
      resetTabBtn.addEventListener("click", _makeTabResetHandler(name));
      panel.appendChild(resetTabBtn);

      _panelsEl.appendChild(panel);
    }
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
  }

  window.openSettingsModal = function (options) {
    _opts = options || {};
    if (_opts.initialTab) _activeTab = _opts.initialTab;
    _open();
    _load();
  };
})();
