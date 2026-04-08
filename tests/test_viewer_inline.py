import viewer


def test_generate_timeline_viewer_inlines_css_and_js(tmp_path, monkeypatch):
    # Run in an isolated working directory so the generated file is easy to inspect.
    monkeypatch.chdir(tmp_path)

    # Use a minimal data payload.
    data = {
        "meta": {},
        "artifacts": [],
        "timeline": {"duration": 0, "startOffset": 0},
    }

    out_path = viewer.generate_timeline_viewer(data, output_basename="viewer.html")
    assert out_path is not None
    assert out_path.is_file()

    html = out_path.read_text(encoding="utf-8")

    # The output should inline CSS/JS instead of linking external files.
    assert '<link rel="stylesheet" href="viewer.css">' not in html
    assert '<script src="viewer.js" defer></script>' not in html
    assert '<script src="utils.js" defer></script>' not in html

    # Expect a style block that contains a recognizable piece of the source CSS.
    assert "<style>" in html
    assert "clipgen Timeline Viewer" in html or "artifact-marker" in html

    # Expect a script block with defer and a recognizable piece of the source JS.
    assert "<script defer>" in html
    assert (
        "clipgen Timeline Viewer \u2013 viewer.js" in html
        or "window.CLIPGEN_DATA" in html
    )

    # Shared utilities from utils.js should be inlined into the script block.
    assert "clipgen shared utilities" in html or "var severityClass" in html
