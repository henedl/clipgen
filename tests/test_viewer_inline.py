import json

import viewer

from _frontend_source import WEB

_VIEWER_JS = WEB / "viewer.js"


def test_generate_timeline_viewer_inlines_css_and_js(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    data = {
        "meta": {},
        "artifacts": [],
        "timeline": {"duration": 0, "startOffset": 0},
    }

    out_path = viewer.generate_timeline_viewer(data, output_basename="viewer.html")
    assert out_path is not None and out_path.is_file()

    html = out_path.read_text(encoding="utf-8")

    # External link/script tags must be replaced by inlined blocks.
    assert '<link rel="stylesheet" href="viewer.css">' not in html
    assert '<script src="viewer.js" defer></script>' not in html
    assert '<script src="utils.js" defer></script>' not in html

    # The card-scrubber module is inlined too (external tags stripped).
    assert '<link rel="stylesheet" href="card-scrubber.css">' not in html
    assert '<script src="card-scrubber.js" defer></script>' not in html
    assert "clipgenCardScrubber" in html
    assert ".waveform-canvas" in html

    # Inlined blocks present, with the data payload bound.
    assert "<style>" in html
    assert "<script defer>" in html
    assert "window.CLIPGEN_DATA" in html


def test_viewer_data_escapes_script_break_sequences(tmp_path, monkeypatch):
    # Regression: user text is embedded into an HTML <script> tag as JSON. json.dumps
    # does not escape `<`, so `</script>` (and the `<!--<script>` double-escape variant)
    # in an observation could close/confuse the tag, breaking the exported viewer and
    # injecting arbitrary HTML. The embed must escape these so the page stays intact.
    monkeypatch.chdir(tmp_path)
    hostile = "see </script><!--<script><b>x</b> &  more"
    data = {
        "meta": {},
        "artifacts": [{"desc": hostile}],
        "timeline": {"duration": 0, "startOffset": 0},
    }

    out_path = viewer.generate_timeline_viewer(data, output_basename="viewer.html")
    assert out_path is not None and out_path.is_file()
    html = out_path.read_text(encoding="utf-8")

    # Isolate the embedded payload from the surrounding template/inlined assets.
    marker = "window.CLIPGEN_DATA="
    start = html.index(marker) + len(marker)
    end = html.index(";</script>", start)
    payload = html[start:end]

    # The raw break sequences from user text must not survive into the payload.
    assert "</script>" not in payload
    assert "<script" not in payload
    assert "<!--" not in payload
    assert " " not in payload

    # ...but the payload is still valid JSON and round-trips the original text unchanged.
    assert json.loads(payload)["artifacts"][0]["desc"] == hostile


def test_viewer_renders_reels_from_data():
    # Regression: reels live in their own `data.reels` slot (no start/end), so the
    # viewer must surface them as playable cards — else a Build Reel → Viewer
    # workflow (or CLI `reel --viewer`) produces an empty viewer.
    src = _VIEWER_JS.read_text(encoding="utf-8")
    assert "data.reels" in src
