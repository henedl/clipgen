"""Behaviour tests for ``clipgenPerf`` in ``assets/web/utils.js``, run under node.

The browser timing API is what ``tests/ui/shot.py --perf`` and the UI bench
read; a lost sample there is a benchmark that quietly under-reports. These
pin the rules the harness relies on: nested same-label spans each record,
recording is bounded and counted past the cap, promise rejections still
record, and nothing records when profiling is off.

Skips cleanly where node is absent, like ``test_intake_cluster.py``.
"""

import json
import shutil
import subprocess

import pytest

from _frontend_source import WEB

NODE = shutil.which("node")

_HARNESS = """
const fs = require("fs"), vm = require("vm");
const el = () => ({
  style: {}, dataset: {},
  classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
  appendChild() {}, setAttribute() {}, getAttribute: () => null,
  addEventListener() {}, removeEventListener() {},
  querySelector: () => null, querySelectorAll: () => [],
});
const ctx = { console, performance };
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
vm.runInContext(fs.readFileSync(WEB + "/utils.js", "utf8"), ctx, { filename: "utils.js" });

const P = ctx.window.clipgenPerf;
const out = {};
const busy = () => { const t = performance.now(); while (performance.now() - t < 1) {} };

ctx.CLIPGEN_CONFIG.profiling = false;
P.span("off", busy);
P.record("off.record", 5);
out.off = P.snapshot();

ctx.CLIPGEN_CONFIG.profiling = true;
P.span("nest", () => P.span("nest", busy));
P.begin("inter"); P.begin("inter"); P.end("inter"); P.end("inter");
P.end("never-begun");
try { P.wrap("throws", () => { throw new Error("x"); })(); } catch (_) {}
P.observe();
const rejected = P.wrap("rejects", () => Promise.reject(new Error("no")))();
rejected.catch(() => {}).then(() => {
  for (let i = 0; i < 600; i++) P.record("cap." + i, 1);
  out.on = P.snapshot();
  out.labelCount = Object.keys(out.on.measures).length;
  process.stdout.write(JSON.stringify(out));
});
"""


@pytest.fixture(scope="module")
def perf(tmp_path_factory) -> dict:
    if NODE is None:
        pytest.skip("node not installed; clipgenPerf behaviour gate skipped")
    script = tmp_path_factory.mktemp("clipgen_perf") / "harness.js"
    script.write_text(
        f"const WEB = {json.dumps(str(WEB))};\n{_HARNESS}", encoding="utf-8"
    )
    result = subprocess.run(
        [NODE, str(script)], check=False, capture_output=True, text=True
    )
    assert result.returncode == 0, f"harness failed:\n{result.stderr}"
    return json.loads(result.stdout)


def test_nothing_records_when_profiling_is_off(perf) -> None:
    assert perf["off"]["measures"] == {}
    assert perf["off"]["dropped"] == 0


def test_nested_same_label_spans_each_record(perf) -> None:
    assert perf["on"]["measures"]["nest"]["n"] == 2
    assert perf["on"]["measures"]["inter"]["n"] == 2
    assert "never-begun" not in perf["on"]["measures"]


def test_wrap_records_throws_and_rejections(perf) -> None:
    assert perf["on"]["measures"]["throws"]["n"] == 1
    assert perf["on"]["measures"]["rejects"]["n"] == 1


def test_label_cap_counts_overflow(perf) -> None:
    assert perf["labelCount"] == 512
    # 600 cap labels + nest/inter/throws/rejects, minus the 512 kept.
    assert perf["on"]["dropped"] == 600 + 4 - 512


def test_unsupported_longtask_observer_is_explicit(perf) -> None:
    # node has no PerformanceObserver in the vm context: false, never 0-as-fact.
    assert perf["on"]["supported"]["longtask"] is False
