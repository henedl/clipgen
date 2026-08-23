"""Source-level wiring assertions for the settings modal's model rows.

The modal opens from every page, so every request it makes is absolute
(``_getApiRoot()`` returns ``/api``) and must hit a rule registered on the
combined app itself, not on a blueprint. A page-relative path would resolve
under whichever prefix happened to serve the page; a leading slash on a
blueprint-only rule 404s from everywhere. Both mistakes are invisible to the
Python tests, which drive the routes directly.
"""

from __future__ import annotations

from _frontend_source import assert_es5, read, strip_comments

_JS = strip_comments(read("settings-modal.js"))


def test_model_rows_carry_reveal_and_delete():
    row = _JS[
        _JS.index("function _buildLlmModelRow") : _JS.index("function _loadModels")
    ]
    # Reveal is chromeless (no .btn), delete keeps the button chrome + label.
    assert '"settings-llm-model-reveal"' in row
    assert '"btn btn-small btn-icon"' in row
    assert '"Delete"' in row
    # Reveal comes first: delete is destructive, so it sits furthest out.
    assert row.index("model-icon--reveal") < row.index("model-icon--delete")


def test_model_icons_are_css_masks_with_an_accessible_name():
    """Icons come from the shared heroicons, never inline SVG path data.

    The reveal button has no label of its own, so it needs the title.
    """
    row = _JS[
        _JS.index("function _buildLlmModelRow") : _JS.index("function _loadModels")
    ]
    assert "<svg" not in row and "<path" not in row
    assert 'showBtn.title = "Show in file browser"' in row
    css = read("settings-modal.css")
    assert 'url("icons/magnifying-glass.svg")' in css
    assert 'url("icons/trash.svg")' in css


def test_model_routes_go_through_the_api_root():
    """No hardcoded prefix and no page-relative path for either model action."""
    assert '_getApiRoot() + "/models/llm/reveal"' in _JS
    assert '_getApiRoot() + "/models/llm/"' in _JS
    assert 'apiPost("api/models' not in _JS
    assert '"/transcripts/api/models' not in _JS


def test_model_actions_report_their_errors():
    """``apiDelete``/``apiPost`` reject on a non-2xx, so the catch owns the message."""
    row = _JS[
        _JS.index("function _buildLlmModelRow") : _JS.index("function _loadModels")
    ]
    assert row.count("_setStatus(") == 2
    assert "res.ok" not in row


def test_settings_modal_is_es5():
    assert_es5(_JS, "settings-modal.js")


def test_unusable_models_are_marked_with_their_reason():
    """A model the router refused must say so before it wastes another run.

    Marked, never hidden or disabled: a llama.cpp upgrade can fix one, and the
    row still has to be deletable.
    """
    row = _JS[
        _JS.index("function _buildLlmModelRow") : _JS.index("function _loadModels")
    ]
    assert "model.unusable" in row
    assert '"settings-llm-model-reason"' in row
    assert "settings-llm-model-row--unusable" in row
    assert "disabled" not in row.split("delBtn")[0], "the row stays actionable"

    select = _JS[_JS.index("function _loadModelsForSelect") :]
    assert "m.unusable" in select
    assert "won't load" in select

    css = read("settings-modal.css")
    assert ".settings-llm-model-reason {" in css
    assert ".settings-llm-model-row--unusable" in css


def test_suggested_rows_offer_download_or_downloaded():
    """A catalog entry downloads in place; an installed one only says so."""
    row = _JS[
        _JS.index("function _buildSuggestedRow") : _JS.index(
            "function _buildLlmModelRow"
        )
    ]
    assert "model.installed" in row
    assert '"Downloaded"' in row
    assert '"Download"' in row
    assert '"btn btn-small btn-icon"' in row
    assert "model-icon--download" in row
    assert "model-icon--done" in row
    assert '"settings-llm-model-bar-fill"' in row
    assert "_watchLlmDownload(model.name" in row

    css = read("settings-modal.css")
    assert 'url("icons/arrow-down-tray.svg")' in css
    assert 'url("icons/check.svg")' in css
    assert ".settings-llm-model-bar-fill {" in css


def test_download_routes_go_through_the_api_root():
    assert '_getApiRoot() + "/models/llm/download"' in _JS
    assert '_getApiRoot() + "/models/llm/download-status?model="' in _JS


def test_download_completion_refreshes_block_and_selects():
    """Cache cleared first, so the block and every dropdown share one fetch."""
    watch = _JS[
        _JS.index("function _watchLlmDownload") : _JS.index(
            "function _buildLlmModelsBlock"
        )
    ]
    assert "_refreshLlmViews()" in watch
    views = _JS[
        _JS.index("function _refreshLlmViews") : _JS.index(
            "function _refreshLlmSelects"
        )
    ]
    assert views.index("_invalidateModels()") < views.index("_llmBlockRefresh()")
    assert "_refreshLlmSelects()" in views


def test_suggested_models_join_the_select():
    """Not-yet-downloaded suggestions are selectable under their own group."""
    select = _JS[_JS.index("function _loadModelsForSelect") :]
    assert 'createElement("optgroup")' in select
    assert 'group.label = "Suggested"' in select
    assert "not downloaded" in select
    assert "if (sm.installed) continue;" in select
    # An installed catalog model is offered under its HF ref, not its stem.
    assert "refByStem[m.name] || m.name" in select
