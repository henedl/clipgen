"""Static regression checks for Screenspace frontend sources.

Focused on the calibration slider annotation (the hairline at the pin-derived
cutoff plus the pass/fail tint on the value readout), whose failure mode is
silent: a renamed class or custom property, or a `background` shorthand
reintroduced on the shared slider rule, makes the mark quietly stop rendering
with no error anywhere.
"""

import re

from _frontend_source import assert_es5, read

CALIBRATION_JS = read("screenspace-calibration.js")
SCREENSPACE_CSS = read("screenspace.css")


def _shared_slider_rule() -> str:
    """The declaration block shared by all three Screenspace range variants."""
    start = SCREENSPACE_CSS.index(
        '.param-control input[type="range"],\n.scene-ref-thresh {'
    )
    return SCREENSPACE_CSS[start : SCREENSPACE_CSS.index("}", start)]


def test_shared_slider_rule_uses_background_color_not_shorthand():
    """`background:` would implicitly reset background-image to none, so the
    .cal-mark hairline would only survive on source-order luck."""
    rule = _shared_slider_rule()
    assert re.search(r"^\s*background-color:", rule, re.MULTILINE), (
        "shared slider rule must set background-color"
    )
    assert not re.search(r"^\s*background:", rule, re.MULTILINE), (
        "shared slider rule must not use the `background` shorthand — it resets "
        "background-image and silently kills the .cal-mark hairline"
    )
    assert "background-repeat: no-repeat;" in rule


def test_slider_thumb_size_variable_is_defined_where_the_thumbs_read_it():
    """The thumb pseudo-elements and .cal-mark both read --slider-thumb-size;
    it has to be declared on the originating element to inherit into them."""
    assert "--slider-thumb-size: 12px;" in _shared_slider_rule()
    for pseudo in ("::-webkit-slider-thumb", "::-moz-range-thumb"):
        block_start = SCREENSPACE_CSS.index(
            f'.param-control input[type="range"]{pseudo},'
        )
        block = SCREENSPACE_CSS[block_start : SCREENSPACE_CSS.index("}", block_start)]
        assert block.count("var(--slider-thumb-size") == 2, (
            f"{pseudo} should size both width and height from --slider-thumb-size"
        )


def test_calibration_slider_classes_have_css_rules():
    """CSS toggle completeness: every class the updater adds must be styled."""
    updater_start = CALIBRATION_JS.index("function updateCalibrationSliderMarks(")
    updater = CALIBRATION_JS[
        updater_start : CALIBRATION_JS.index("\n  }", updater_start)
    ]
    toggled = set(re.findall(r"classList\.(?:add|remove)\(([^)]*)\)", updater))
    classes = {c for group in toggled for c in re.findall(r'"([\w-]+)"', group)}
    assert classes == {"cal-mark", "cal-pass", "cal-fail"}, classes
    for name in classes:
        assert f".{name}" in SCREENSPACE_CSS, f".{name} is toggled but never styled"


def test_calibration_custom_properties_round_trip():
    """Custom-property names have no toggle-completeness guard of their own, and
    a rename typo degrades silently — pin the JS↔CSS pairing both ways."""
    written = set(
        re.findall(r'(?:setProperty|removeProperty)\("(--cal-[\w-]+)"', CALIBRATION_JS)
    )
    assert written == {"--cal-mark-frac"}, written
    for name in written:
        assert f"var({name})" in SCREENSPACE_CSS, f"{name} is set in JS but never read"
    read_in_css = set(re.findall(r"var\((--cal-mark-[\w-]+)", SCREENSPACE_CSS))
    # --cal-mark-x is derived inside the rule itself, not written from JS.
    assert read_in_css - {"--cal-mark-x"} == written, read_in_css


def test_calibration_satellite_is_es5():
    assert_es5(CALIBRATION_JS, "screenspace-calibration.js")
