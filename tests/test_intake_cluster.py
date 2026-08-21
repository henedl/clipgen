"""Behaviour tests for ``assets/web/intake-cluster.js``, run through node.

The rest of the frontend suite greps source text; this module actually *runs*
JS. ``intake-cluster.js`` is the one file where that is cheap: it is pure by
design (no DOM, no module state, plain in/out helpers), so a small ``vm``
context plus a stub ``document`` is enough to load it — along with the real
``utils.js`` ``severityRank``, so the severity hoisting is checked against the
shipped implementation rather than a copy.

Why it earns its place: Studio, Composer, and the Transcripts "Clip Marked
Lines" action all cluster through these two functions, and the navigational
(boundary) rules below used to be recorded only in a source comment. A merged
boundary cluster silently draws one tick and hides the rest, and padding a
boundary by ±5s puts every downstream card range and clip window off by five
seconds. Neither shows up in a syntax or wiring check.

Skips cleanly where node is absent, exactly like ``test_frontend_syntax.py``.
"""

import json
import shutil
import subprocess

import pytest

from _frontend_source import WEB

NODE = shutil.which("node")

# Minimal browser surface: enough for utils.js to reach its own bottom.
_HARNESS = """
const fs = require("fs"), vm = require("vm");
const el = () => ({
  style: {}, dataset: {},
  classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
  appendChild() {}, setAttribute() {}, getAttribute: () => null,
  addEventListener() {}, removeEventListener() {},
  querySelector: () => null, querySelectorAll: () => [],
});
const ctx = { console };
ctx.window = ctx;
ctx.globalThis = ctx;
ctx.addEventListener = () => {};
ctx.removeEventListener = () => {};
ctx.document = Object.assign(el(), {
  documentElement: el(), body: el(), head: el(),
  createElement: el, readyState: "complete",
});
ctx.localStorage = { getItem: () => null, setItem() {}, removeItem() {} };
ctx.sessionStorage = { getItem: () => null, setItem() {}, removeItem() {} };
ctx.navigator = { userAgent: "node", platform: "node" };
ctx.location = { href: "http://x/", search: "", hash: "", pathname: "/" };
ctx.matchMedia = () => ({ matches: false, addEventListener() {}, addListener() {} });
ctx.getComputedStyle = () => ({ getPropertyValue: () => "" });
ctx.setTimeout = setTimeout;
ctx.clearTimeout = clearTimeout;
ctx.requestAnimationFrame = (f) => setTimeout(f, 0);
ctx.fetch = () => Promise.resolve({ json: () => Promise.resolve({}) });
vm.createContext(ctx);

for (const f of ["utils.js", "intake-cluster.js"]) {
  vm.runInContext(fs.readFileSync(WEB + "/" + f, "utf8"), ctx, { filename: f });
}

const API = ctx.window.ClipgenIntakeCluster;
const cases = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const out = {};
for (const [name, spec] of Object.entries(cases)) {
  out[name] = API[spec.fn](spec.items, spec.threshold);
}
process.stdout.write(JSON.stringify(out));
"""


def _event(participant, time_in, time_out, **kw):
    ev = {
        "participant": participant,
        "event_type": kw.get("event_type", "color"),
        "time_in": time_in,
        "time_out": time_out,
        "confidence": kw.get("confidence", 1.0),
        "source_video": kw.get("source_video", "s_P01.mp4"),
        "detector": kw.get("detector", "d"),
        "region": kw.get("region", "r"),
    }
    if "navigational" in kw:
        ev["navigational"] = kw["navigational"]
    return ev


def _mark(participant, start, end, **kw):
    return {
        "participant": participant,
        "start": start,
        "end": end,
        "text": kw.get("text", ""),
        "label": kw.get("label", ""),
        "category": kw.get("category", ""),
        "severity": kw.get("severity", ""),
    }


CASES = {
    "merge_within_threshold": {
        "fn": "clusterIntakeEvents",
        "threshold": 30,
        "items": [_event("P01", 10, 12), _event("P01", 20, 22)],
    },
    "split_beyond_threshold": {
        "fn": "clusterIntakeEvents",
        "threshold": 5,
        "items": [_event("P01", 10, 12), _event("P01", 40, 42)],
    },
    "split_by_event_type": {
        "fn": "clusterIntakeEvents",
        "threshold": 300,
        "items": [_event("P01", 10, 12), _event("P01", 11, 13, event_type="text")],
    },
    "split_by_participant": {
        "fn": "clusterIntakeEvents",
        "threshold": 300,
        "items": [_event("P01", 10, 12), _event("P02", 11, 13)],
    },
    # Three boundaries well inside the threshold: still three clusters.
    "navigational_never_merges": {
        "fn": "clusterIntakeEvents",
        "threshold": 300,
        "items": [
            _event("P01", 10, 10, event_type="boundary", navigational=True),
            _event("P01", 11, 11, event_type="boundary", navigational=True),
            _event("P01", 12, 12, event_type="boundary", navigational=True),
        ],
    },
    "navigational_keeps_exact_time": {
        "fn": "clusterIntakeEvents",
        "threshold": 300,
        "items": [_event("P01", 30, 30, event_type="boundary", navigational=True)],
    },
    "zero_width_gets_padded": {
        "fn": "clusterIntakeEvents",
        "threshold": 300,
        "items": [_event("P01", 30, 30)],
    },
    "padding_floors_at_zero": {
        "fn": "clusterIntakeEvents",
        "threshold": 300,
        "items": [_event("P01", 2, 2)],
    },
    "confidence_is_averaged": {
        "fn": "clusterIntakeEvents",
        "threshold": 300,
        "items": [
            _event("P01", 10, 12, confidence=0.4),
            _event("P01", 14, 16, confidence=0.8),
        ],
    },
    "marks_merge_within_threshold": {
        "fn": "clusterTranscriptMarks",
        "threshold": 30,
        "items": [_mark("P01", 10, 12, text="a"), _mark("P01", 20, 25, text="b")],
    },
    "marks_split_beyond_threshold": {
        "fn": "clusterTranscriptMarks",
        "threshold": 5,
        "items": [_mark("P01", 10, 12), _mark("P01", 40, 42)],
    },
    "marks_split_by_participant": {
        "fn": "clusterTranscriptMarks",
        "threshold": 300,
        "items": [_mark("P01", 10, 12), _mark("P02", 11, 13)],
    },
    "marks_hoist_worst_severity": {
        "fn": "clusterTranscriptMarks",
        "threshold": 300,
        "items": [
            _mark("P01", 10, 12, severity="Low"),
            _mark("P01", 14, 16, severity="Critical"),
            _mark("P01", 18, 20, severity="Medium"),
        ],
    },
}


@pytest.fixture(scope="module")
def clusters(tmp_path_factory) -> dict:
    """Run every case in one node process; each test asserts on its slice."""
    if NODE is None:
        pytest.skip("node not installed; intake-cluster behaviour gate skipped")
    tmp = tmp_path_factory.mktemp("intake_cluster")
    cases_path = tmp / "cases.json"
    cases_path.write_text(json.dumps(CASES), encoding="utf-8")
    script = tmp / "harness.js"
    script.write_text(
        f"const WEB = {json.dumps(str(WEB))};\n{_HARNESS}", encoding="utf-8"
    )
    result = subprocess.run(
        [NODE, str(script), str(cases_path)],
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, f"harness failed:\n{result.stderr}"
    return json.loads(result.stdout)


def test_merges_runs_closer_than_threshold(clusters) -> None:
    got = clusters["merge_within_threshold"]
    assert len(got) == 1
    assert (got[0]["start"], got[0]["end"]) == (10, 22)
    assert len(got[0]["events"]) == 2


def test_splits_runs_beyond_threshold(clusters) -> None:
    assert len(clusters["split_beyond_threshold"]) == 2


def test_never_merges_across_event_type(clusters) -> None:
    assert len(clusters["split_by_event_type"]) == 2


def test_never_merges_across_participant(clusters) -> None:
    assert len(clusters["split_by_participant"]) == 2


def test_navigational_events_each_get_a_cluster(clusters) -> None:
    """Boundaries are point ticks: a merged cluster would hide all but the first."""
    got = clusters["navigational_never_merges"]
    assert len(got) == 3, "boundary events must not merge, however close"
    assert all(c["navigational"] is True for c in got), (
        "the flag must reach the cluster"
    )


def test_navigational_cluster_is_not_padded(clusters) -> None:
    """A boundary is a precise instant; ±5s would skew ranges and clip windows."""
    got = clusters["navigational_keeps_exact_time"]
    assert (got[0]["start"], got[0]["end"]) == (30, 30)


def test_zero_width_cluster_is_padded(clusters) -> None:
    got = clusters["zero_width_gets_padded"]
    assert (got[0]["start"], got[0]["end"]) == (25, 35)


def test_padding_never_goes_negative(clusters) -> None:
    got = clusters["padding_floors_at_zero"]
    assert (got[0]["start"], got[0]["end"]) == (0, 7)


def test_confidence_is_averaged_across_members(clusters) -> None:
    got = clusters["confidence_is_averaged"]
    assert got[0]["confidence_avg"] == pytest.approx(0.6)


def test_marks_merge_and_concatenate_text(clusters) -> None:
    got = clusters["marks_merge_within_threshold"]
    assert len(got) == 1
    assert (got[0]["start"], got[0]["end"]) == (10, 25)
    assert got[0]["text"] == "a b"


def test_marks_split_beyond_threshold(clusters) -> None:
    assert len(clusters["marks_split_beyond_threshold"]) == 2


def test_marks_never_merge_across_participant(clusters) -> None:
    assert len(clusters["marks_split_by_participant"]) == 2


def test_marks_hoist_the_worst_severity(clusters) -> None:
    """Ranked through the real ``severityRank``: lower rank wins."""
    got = clusters["marks_hoist_worst_severity"]
    assert len(got) == 1
    assert got[0]["severity"] == "Critical"
