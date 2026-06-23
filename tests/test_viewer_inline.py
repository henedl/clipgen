import viewer


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
