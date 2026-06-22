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
  // Titlecard/endcard picker state (shared by the title + end pickers).
  var _cardsCache = null;
  var _cardsCachePromise = null;
  var _cardPickers = [];
  var _SAMPLE_TITLE_TEXT = "Sample description";

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
    // Dismiss the inline color popover; it lives on document.body at --z-toast
    // and would otherwise outlive the modal.
    if (window.ClipgenColorPicker) window.ClipgenColorPicker.close();
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
    // Dismiss any stale color popover and refetch the card list each time the
    // modal opens so externally added or removed uploads show up.
    if (window.ClipgenColorPicker) window.ClipgenColorPicker.close();
    _cardsCache = null;
    _cardsCachePromise = null;
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
    } else if (s.type === "card_picker") {
      row.classList.add("settings-row-stacked");
      var kind = s.kind === "end" ? "end" : "title";
      _cardPickers.push({ container: controlDiv, settingName: settingName, kind: kind });
      _renderCardPicker(controlDiv, settingName, kind);
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

  // ── Titlecard / endcard background picker ──────────────────────────────

  function _fetchCards(force) {
    if (force) {
      _cardsCache = null;
      _cardsCachePromise = null;
    }
    if (_cardsCache) return Promise.resolve(_cardsCache);
    if (_cardsCachePromise) return _cardsCachePromise;
    _cardsCachePromise = fetch(_getApiRoot() + "/titlecards")
      .then(function (r) { return r.json(); })
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
      img.src = item.url;
      img.alt = item.label;
      img.loading = "lazy";
      preview.appendChild(img);
    } else if (item.kind === "none") {
      preview.appendChild(el("span", "card-tile-placeholder", "No endcard"));
    } else if (item.kind === "color") {
      preview.style.background = _cardCurrentColor(kind);
    }
    // Approximate the ffmpeg drawtext overlay on titlecards: centered sample
    // text over the chosen background. Endcards carry no text.
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
    grid.appendChild(el("div", "card-tile card-tile--loading", "Loading…"));
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
    if (_statusEl) _statusEl.textContent = "Uploading…";
    var form = new FormData();
    form.append("file", file);
    fetch(_getApiRoot() + "/titlecards/upload", { method: "POST", body: form })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data || !data.ok) {
          if (_statusEl) {
            _statusEl.textContent = data && data.error ? data.error : "Upload failed";
          }
          return;
        }
        if (_statusEl) _statusEl.textContent = "Uploaded";
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
        if (_statusEl) _statusEl.textContent = "Upload failed";
      });
  }

  function _deleteCard(name) {
    if (_statusEl) _statusEl.textContent = "Deleting…";
    fetch(_getApiRoot() + "/titlecards/image/" + encodeURIComponent(name), {
      method: "DELETE",
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data || !data.ok) {
          if (_statusEl) {
            _statusEl.textContent = data && data.error ? data.error : "Delete failed";
          }
          return;
        }
        if (_statusEl) _statusEl.textContent = "Deleted";
        // The server resets any selection that pointed at the deleted file;
        // mirror that into the in-memory settings so the UI stays in sync.
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
        if (_statusEl) _statusEl.textContent = "Delete failed";
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
        // Hidden settings are persisted + sent to the client but have no row of
        // their own (e.g. the card colors edited via the card picker's swatch).
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
