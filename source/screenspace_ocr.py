"""Screenspace OCR + numeric helpers."""

import difflib
import math
import os
import queue
import re
import threading
import time
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

import cv2
import numpy as np

import config
import profiling
from utils import get_bundled_assets_root
from screenspace_primitives import extract_region, point_in_mask_points


# ---------------------------------------------------------------------------
# Engine pool
# ---------------------------------------------------------------------------
#
# RapidOCR gives no official thread-safety guarantee for a shared engine, and
# each engine owns private onnxruntime sessions, so each caller borrows its own
# engine from a small per-model pool rather than sharing one behind a global
# lock (which would cap OCR concurrency at 1 regardless of
# SCREENSPACE_PARALLEL_WORKERS). Bounded rather than per-thread: Flask's
# ephemeral calibration request threads would otherwise accumulate one model
# copy each.

_ocr_pools: dict[str, queue.Queue] = {}
_ocr_pool_lock = threading.Lock()  # guards _ocr_pools creation
_ocr_build_lock = threading.Lock()  # serializes first-time engine construction

# Language → recognition-model family. Detection/orientation always use the
# wheel's bundled defaults; only the recognition model varies by script.
# "default" is the wheel's offline PP-OCR Chinese+English model; the other
# families are vendored at build time (build/fetch_binaries.py OCR_MODEL_PINS,
# bundled as ocr_models/) and auto-downloaded by rapidocr in source checkouts
# that skipped the fetch. The UI language dropdown mirrors this table's keys.
_OCR_MODEL_DEFAULT = "default"
_OCR_LANG_TO_MODEL: dict[str, str] = {
    "en": _OCR_MODEL_DEFAULT,
    "zh": _OCR_MODEL_DEFAULT,  # the default rec model covers Chinese + English
    "es": "latin",
    "fr": "latin",
    "de": "latin",
    "ja": "japan",
    "ko": "korean",
}


def _resolve_ocr_model(languages: list[str] | None) -> str:
    """Map a UI language list onto one recognition-model family."""
    langs = list(languages or ["en"])
    unknown = [lang for lang in langs if lang not in _OCR_LANG_TO_MODEL]
    if unknown:
        raise ValueError(f"Unsupported OCR language(s): {unknown}")
    models = {_OCR_LANG_TO_MODEL[lang] for lang in langs}
    if models == {_OCR_MODEL_DEFAULT, "latin"}:
        return "latin"  # the latin model covers English glyphs too
    if len(models) == 1:
        return next(iter(models))
    raise ValueError(f"OCR languages {langs} need incompatible recognition models")


def _vendored_rec_model(model: str) -> Any:
    """Locate the build-time-vendored recognition model for a model family.

    ONNX rec models embed their character dict in the model metadata, so a
    single ``.onnx`` file is the whole vendored artifact. Frozen bundles carry
    them under ``<_MEIPASS>/ocr_models/``; source checkouts that ran
    ``build/fetch_binaries.py`` have them in ``build/vendor/ocr/``. Returns the
    path, or ``None`` when absent (dev fallback: rapidocr's own pinned
    download).
    """
    root = get_bundled_assets_root()
    for base in (root / "ocr_models", root / "build" / "vendor" / "ocr"):
        onnx = base / f"{model}_rec.onnx"
        if onnx.is_file():
            return onnx
    return None


def _build_ocr_reader(model: str) -> Any:
    """Construct one RapidOCR engine (the sole engine-construction site)."""
    from rapidocr import LangRec, ModelType, OCRVersion, RapidOCR

    params: dict[str, Any] = {
        "Global.log_level": "error",
        # Each pooled engine owns a private onnxruntime session; cap intra-op
        # threads so pool_size × threads cannot oversubscribe the CPU.
        "EngineConfig.onnxruntime.intra_op_num_threads": max(
            1, (os.cpu_count() or 4) // _ocr_pool_size()
        ),
    }
    if model != _OCR_MODEL_DEFAULT:
        vendored = _vendored_rec_model(model)
        if vendored is not None:
            params["Rec.model_path"] = str(vendored)
        else:
            # Source checkout without build/vendor/ocr: let rapidocr fetch its
            # own pinned model (dev only — frozen bundles always hit the
            # vendored path). Versions match OCR_MODEL_PINS in
            # build/fetch_binaries.py: v5 has no japan ONNX model, hence v4.
            params["Rec.lang_type"] = {
                "latin": LangRec.LATIN,
                "japan": LangRec.JAPAN,
                "korean": LangRec.KOREAN,
            }[model]
            params["Rec.ocr_version"] = (
                OCRVersion.PPOCRV4 if model == "japan" else OCRVersion.PPOCRV5
            )
            params["Rec.model_type"] = ModelType.MOBILE
    return RapidOCR(params=params)


def _ocr_pool_size() -> int:
    """Max concurrent engines per model family (auto = parallel worker count)."""
    size = config.SCREENSPACE_OCR_POOL_SIZE or config.SCREENSPACE_PARALLEL_WORKERS
    return max(1, size)


def _get_ocr_pool(languages: list[str]) -> queue.Queue:
    """Return the (lazily created) engine pool for the given language set.

    The pool is seeded with ``_ocr_pool_size()`` ``None`` placeholder slots;
    each slot's engine is built on first checkout. Keyed by the resolved
    recognition-model family, so e.g. es/fr/de share one pool.
    """
    key = _resolve_ocr_model(languages)
    with _ocr_pool_lock:
        pool = _ocr_pools.get(key)
        if pool is None:
            pool = queue.Queue()
            for _ in range(_ocr_pool_size()):
                pool.put(None)
            _ocr_pools[key] = pool
        return pool


@contextmanager
def _checkout_ocr_reader(languages: list[str]) -> Iterator[Any]:
    """Borrow an engine from the bounded per-model pool for one call.

    ``pool.get()`` blocks while every engine is busy, capping concurrency at the
    pool size. ``None`` slots are built on first use under ``_ocr_build_lock`` so
    a first-use model fetch (dev fallback) can't race. The slot always goes
    back — ``None`` again if construction raised — so the pool never shrinks.
    """
    model = _resolve_ocr_model(languages)
    pool = _get_ocr_pool(languages)
    # The one number that settles SCREENSPACE_OCR_POOL_SIZE, which is otherwise
    # tuned on reasoning alone: time spent blocked here is OCR concurrency the
    # pool is refusing. Direct analogue of worker.progress_lock_wait, and one
    # add() per OCR call rather than per frame, so the hot-loop rule holds.
    _t0 = time.perf_counter() if config.PROFILING else 0.0
    reader = pool.get()
    if _t0:
        profiling.add("ocr.pool_wait", time.perf_counter() - _t0)
    try:
        if reader is None:
            with _ocr_build_lock, profiling.span("ocr.reader_build"):
                reader = _build_ocr_reader(model)
        yield reader
    finally:
        pool.put(reader)


def _readings_from_result(result: Any) -> list[Any]:
    """Adapt a ``RapidOCROutput`` to raw ``[(bbox, text, conf)]`` readings.

    ``bbox`` is a list of four ``(x, y)`` floats in pixel space — the same shape
    EasyOCR produced, which every downstream consumer unpacks. Plain tuples,
    not ndarrays: readings cross the server's per-pin OCR cache boundary and
    nothing numpy may leak toward JSON. An empty result carries ``None`` fields.
    """
    boxes = getattr(result, "boxes", None)
    txts = getattr(result, "txts", None)
    if boxes is None or txts is None:
        return []
    scores = getattr(result, "scores", None)
    if scores is None:
        scores = [0.0] * len(txts)
    readings: list[Any] = []
    for box, text, score in zip(boxes, txts, scores, strict=False):
        points = [
            (float(x), float(y)) for x, y in np.asarray(box, dtype=float).reshape(-1, 2)
        ]
        readings.append((points, str(text), float(score)))
    return readings


def _ocr_readtext(languages: list[str], image: np.ndarray) -> list[Any]:
    """Run a pooled RapidOCR engine over *image* for the given language set."""
    with _checkout_ocr_reader(languages) as reader:
        return _readings_from_result(reader(image))


def _preprocess_for_ocr(pixels: np.ndarray, *, min_height: int = 0) -> np.ndarray:
    """Enhance a region crop for OCR: upscale small crops and boost local contrast.

    Compressed HUDs render text below the OCR engine's comfortable size and contrast,
    so crops shorter than ``min_height`` are cubic-upscaled (aspect preserved) and
    CLAHE-equalized. Opt-in per task: it costs a few ms/frame and can ring on
    already-clean text. Returns 3-channel BGR so the downstream call is identical
    to the raw-crop path.
    """
    if min_height <= 0:
        min_height = config.SCREENSPACE_OCR_MIN_HEIGHT
    h, w = pixels.shape[:2]
    if h == 0 or w == 0:
        return pixels
    if h < min_height:
        scale = min_height / float(h)
        new_w = max(1, round(w * scale))
        pixels = cv2.resize(pixels, (new_w, min_height), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY) if pixels.ndim == 3 else pixels
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    return cv2.cvtColor(enhanced, cv2.COLOR_GRAY2BGR)


# Opt-in confusion-collapsing for the text tool: fold the glyphs OCR engines
# most often swap on compressed footage before the fuzzy compare. Either toward
# digits ("100" matches a reading of "l00") or letters ("stop" matches "5top").
# The pairs were tuned against EasyOCR's misread profile; re-tune against
# PP-OCR's once real-footage misreads accumulate (the mechanism is unchanged).
_OCR_FOLD_TO_DIGITS = str.maketrans(
    {"o": "0", "l": "1", "i": "1", "|": "1", "s": "5", "b": "8"}
)
_OCR_FOLD_TO_LETTERS = str.maketrans({"0": "o", "1": "l", "5": "s", "8": "b"})


def _normalize_ocr_text(s: str, mode: str) -> str:
    """Lowercase, then fold easily-confused glyphs toward ``mode``'s canonical form.

    ``mode`` is ``"digits"`` (fold letters→digits), ``"letters"`` (fold
    digits→letters), or ``"off"`` (lowercase only — also the fallback for any
    unrecognized value).
    """
    s = s.lower()
    if mode == "digits":
        return s.translate(_OCR_FOLD_TO_DIGITS)
    if mode == "letters":
        return s.translate(_OCR_FOLD_TO_LETTERS)
    return s


def _effective_ocr_confidence_threshold(value: Any = None) -> float:
    """Return OCR confidence cutoff, using config default only when omitted."""
    if value is None:
        return config.SCREENSPACE_OCR_MIN_CONFIDENCE
    try:
        threshold = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError("ocr_confidence_threshold must be a number") from exc
    if not math.isfinite(threshold):
        raise ValueError("ocr_confidence_threshold must be a finite number")
    if threshold < 0 or threshold > 1:
        raise ValueError("ocr_confidence_threshold must be between 0 and 1")
    return threshold


_NUMBERS_RE = re.compile(r"-?\d+(?:\.\d+)?")
_VALID_OPERATORS = ("eq", "gt", "lt", "gte", "lte", "range")


def _number_matches(
    value: float,
    operator: str,
    target_value: float = 0,
    range_min: float | None = None,
    range_max: float | None = None,
) -> bool:
    """Check if *value* satisfies the given numeric comparison."""
    if operator == "eq":
        return value == target_value
    elif operator == "gt":
        return value > target_value
    elif operator == "lt":
        return value < target_value
    elif operator == "gte":
        return value >= target_value
    elif operator == "lte":
        return value <= target_value
    elif operator == "range":
        return (
            range_min is not None
            and range_max is not None
            and range_min <= value <= range_max
        )
    return False


def _ocr_region_readings(
    region_pixels: np.ndarray,
    *,
    languages: list[str] | None = None,
    preprocess: bool = False,
    mask_points: list[Any] | None = None,
) -> list[Any]:
    """Run OCR over a region crop and return raw ``(bbox, text, conf)`` tuples.

    Pure transport — no fuzzy/threshold logic — so the same readings can be
    re-scored under different settings (the calibration OCR cache relies on this).

    For shaped regions, *mask_points* (bbox-relative contours) drops readings whose
    bbox center falls outside every contour. OCR still sees the full rect, since
    masking glyph pixels would corrupt recognition. Centers normalize by the
    *post-preprocess* shape so the test survives the OCR upscale.
    """
    langs = languages or ["en"]
    pixels = _preprocess_for_ocr(region_pixels) if preprocess else region_pixels
    readings = _ocr_readtext(langs, pixels)
    if mask_points:
        img_h, img_w = pixels.shape[:2]
        if img_h > 0 and img_w > 0:
            readings = [
                r
                for r in readings
                if point_in_mask_points(
                    sum(float(p[0]) for p in r[0]) / (len(r[0]) * img_w),
                    sum(float(p[1]) for p in r[0]) / (len(r[0]) * img_h),
                    mask_points,
                )
            ]
    return readings


def _score_text_readings(
    readings: list[Any], params: dict[str, Any]
) -> tuple[bool, dict[str, Any]]:
    """Score text OCR readings: best fuzzy ratio among readings clearing min conf.

    ``fuzzy_ratio`` is the calibration scalar; ``text_found``/``confidence`` carry
    the best-matching reading for the strip tooltip. ``passed`` is the fuzzy
    match at the current threshold.
    """
    search_string = params.get("search_string", "")
    fuzzy_threshold = params.get(
        "fuzzy_threshold", config.SCREENSPACE_OCR_FUZZY_THRESHOLD
    )
    ocr_min_conf = _effective_ocr_confidence_threshold(
        params.get("ocr_confidence_threshold")
    )
    ocr_normalize = params.get("ocr_normalize") or "off"
    if ocr_normalize not in ("digits", "letters"):
        ocr_normalize = "off"
    search_cmp = _normalize_ocr_text(search_string, ocr_normalize)
    best_ratio = 0.0
    best_text = ""
    best_conf = 0.0
    for _, text, conf in readings:
        if conf < ocr_min_conf:
            continue
        ocr_cmp = _normalize_ocr_text(text, ocr_normalize)
        ratio = difflib.SequenceMatcher(None, search_cmp, ocr_cmp).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_text = text
            best_conf = conf
    return best_ratio >= fuzzy_threshold, {
        "fuzzy_ratio": round(best_ratio, 4),
        "text_found": best_text,
        "confidence": round(best_conf, 4),
    }


def _score_numbers_readings(
    readings: list[Any], params: dict[str, Any]
) -> tuple[bool, dict[str, Any]]:
    """Score numbers OCR readings by the best condition-matching OCR confidence.

    ``confidence`` is the calibration scalar for the OCR-confidence slider: it
    reflects the best reading that satisfies operator / target_value / range,
    not an unrelated high-confidence number. ``passed`` then applies the current
    confidence threshold to that matching reading.

    ``integers_only`` rejects any extracted value carrying a decimal part or a
    sign, so "3.5" or "-12" can never satisfy a whole-number HUD condition.
    (Post-filtering replaces the old English-only EasyOCR recognition allowlist;
    it now applies to every language.)
    """
    operator = params.get("operator", "gt")
    target_value = params.get("target_value", 0)
    range_min = params.get("range_min")
    range_max = params.get("range_max")
    integers_only = bool(params.get("integers_only", False))
    ocr_min_conf = _effective_ocr_confidence_threshold(
        params.get("ocr_confidence_threshold")
    )
    matched_number: float | None = None
    matched_conf = 0.0
    for _, text, conf in readings:
        cleaned = text.replace(",", "")
        for match in _NUMBERS_RE.findall(cleaned):
            if integers_only and not match.isdigit():
                continue  # decimal or signed reading — not a whole-number value
            num = float(match)
            if _number_matches(num, operator, target_value, range_min, range_max) and (
                matched_number is None or conf > matched_conf
            ):
                matched_number = num
                matched_conf = conf
    detail: dict[str, Any] = {"confidence": round(matched_conf, 4)}
    if matched_number is not None:
        detail["number_found"] = matched_number
    return matched_number is not None and matched_conf >= ocr_min_conf, detail


def run_calibration_ocr(
    frame: np.ndarray,
    region: dict[str, Any],
    params: dict[str, Any],
) -> list[Any]:
    """Run OCR for one calibration frame/region (text or numbers tool).

    Public entry point for the server's per-pin OCR cache: returns the raw
    ``(bbox, text, conf)`` readings so they can be memoized and re-scored under
    changed fuzzy/confidence settings without re-running OCR. Text and numbers
    pins share readings — ``integers_only`` and the numeric operators are
    applied at scoring time, not here.
    """
    languages = params.get("languages") or ["en"]
    pixels = extract_region(frame, region)
    return _ocr_region_readings(
        pixels,
        languages=languages,
        preprocess=params.get("ocr_preprocess", False),
        mask_points=region.get("mask_points"),
    )
