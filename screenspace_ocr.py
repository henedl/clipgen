"""Screenspace OCR + numeric helpers.

Pooled EasyOCR readers, region preprocessing, glyph confusion-folding, the
numeric comparison helpers and allowlists, region reading/scoring, and the
calibration OCR entry point. Imports region cropping from screenspace_primitives.
"""

import difflib
import math
import queue
import re
import threading
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

import cv2
import numpy as np

import config
from screenspace_primitives import extract_region, point_in_mask_points


# ---------------------------------------------------------------------------
# Reader pool
# ---------------------------------------------------------------------------
#
# EasyOCR/torch inference on a shared Reader is not thread-safe: with parallel
# Screenspace workers (and the calibration route thread) hitting one Reader,
# concurrent readtext calls can corrupt results or crash. Rather than serialize
# every readtext behind a global lock (which capped OCR concurrency at 1 despite
# SCREENSPACE_PARALLEL_WORKERS), keep a small bounded pool of Readers per
# language set and hand a distinct Reader to each concurrent caller. A Reader is
# only ever touched by one thread at a time, so no inference lock is needed and
# up to pool-size OCR calls run truly in parallel. The pool is bounded (not a
# per-thread reader) because the Flask calibration path spawns ephemeral request
# threads, which would otherwise accumulate one model copy per thread.

_ocr_pools: dict[tuple, queue.Queue] = {}
_ocr_pool_lock = threading.Lock()  # guards _ocr_pools creation
_ocr_build_lock = threading.Lock()  # serializes first-time Reader construction


def _build_ocr_reader(languages: list[str]) -> Any:
    """Construct one EasyOCR Reader (the sole reader-construction site)."""
    import easyocr

    return easyocr.Reader(
        list(languages), gpu=config.SCREENSPACE_OCR_GPU, verbose=False
    )


def _ocr_pool_size() -> int:
    """Max concurrent Readers per language set (auto = parallel worker count)."""
    size = config.SCREENSPACE_OCR_POOL_SIZE or config.SCREENSPACE_PARALLEL_WORKERS
    return max(1, size)


def _get_ocr_pool(languages: list[str]) -> queue.Queue:
    """Return the (lazily created) Reader pool for the given language set.

    The pool is seeded with ``_ocr_pool_size()`` ``None`` placeholder slots;
    each slot's Reader is built on first checkout.
    """
    key = tuple(sorted(languages))
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
    """Borrow a Reader from the bounded per-language pool for the duration of a call.

    ``pool.get()`` blocks when every Reader is busy, so concurrency is capped at
    the pool size. A ``None`` slot is built on first use under ``_ocr_build_lock``
    (serializing only construction, so the initial model-file download can't
    race). The slot is always returned to the pool — ``None`` again if
    construction raised — so the pool's slot count never shrinks.
    """
    key = tuple(sorted(languages))
    pool = _get_ocr_pool(languages)
    reader = pool.get()
    try:
        if reader is None:
            with _ocr_build_lock:
                reader = _build_ocr_reader(list(key))
        yield reader
    finally:
        pool.put(reader)


def _ocr_readtext(languages: list[str], image: np.ndarray, **kwargs: Any) -> list[Any]:
    """Run ``readtext`` on a pooled Reader for the given language set."""
    with _checkout_ocr_reader(languages) as reader:
        return reader.readtext(image, **kwargs)


def _preprocess_for_ocr(pixels: np.ndarray, *, min_height: int = 0) -> np.ndarray:
    """Enhance a region crop for OCR: upscale small crops and boost local contrast.

    Compressed game HUDs frequently render text below EasyOCR's comfortable size
    and at low contrast. Crops shorter than ``min_height`` are upscaled with cubic
    interpolation (preserving aspect ratio), then CLAHE (contrast-limited adaptive
    histogram equalization) separates faint glyphs from the background. Opt-in per
    task because the resize + equalization costs a few ms/frame and can introduce
    ringing on already-clean text.

    Returns a 3-channel BGR array (EasyOCR accepts grayscale, but stacking keeps
    the downstream call identical to the raw crop).
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


# Opt-in OCR confusion-collapsing for the text tool. Folds glyphs EasyOCR most
# often swaps on compressed footage toward a canonical form before the fuzzy
# compare. Two directions: fold toward digits (so a search for "100" matches an
# OCR reading of "l00") or toward letters (so a search for "stop" matches "5top").
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

# EasyOCR character allowlist for the numbers tool. Numbers mode is digits-only,
# so constraining recognition kills glyph confusions (O↔0, S↔5, l↔1) at the
# source. Only passed for the default English reader — some language combos
# reject ``allowlist``. Mirrors what the downstream parser accepts (``-``, ``.``,
# ``,`` thousands separators, digits).
_OCR_NUMBER_ALLOWLIST = "0123456789.,-"
# integers_only narrows the allowlist further: dropping ``.,-`` stops a separator
# or sign glyph from surviving OCR as a digit and inflating the parsed value.
# Use for whole-number HUD targets (scores, counts) where decimals never appear.
_OCR_DIGITS_ONLY_ALLOWLIST = "0123456789"


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


def _numbers_ocr_allowlist(languages: list[str], integers_only: bool) -> str | None:
    """EasyOCR allowlist for the numbers tool (English only; ``None`` otherwise)."""
    if languages == ["en"]:
        return _OCR_DIGITS_ONLY_ALLOWLIST if integers_only else _OCR_NUMBER_ALLOWLIST
    return None


def _ocr_region_readings(
    region_pixels: np.ndarray,
    *,
    languages: list[str] | None = None,
    allowlist: str | None = None,
    preprocess: bool = False,
    mask_points: list[Any] | None = None,
) -> list[Any]:
    """Run EasyOCR over a region crop and return raw ``(bbox, text, conf)`` tuples.

    Pure transport over the cached reader — no fuzzy/threshold logic — so the
    same readings can be re-scored under different fuzzy/confidence settings
    (the basis of the calibration OCR cache).

    For shaped regions, *mask_points* (bbox-relative contour list) drops
    readings whose bbox center falls outside every contour. OCR always sees the
    full rect (masking glyph pixels would corrupt recognition); only the
    readings are filtered. Centers are normalized by the *post-preprocess*
    image shape so the test is unaffected by the OCR upscale.
    """
    langs = languages or ["en"]
    pixels = _preprocess_for_ocr(region_pixels) if preprocess else region_pixels
    kwargs: dict[str, Any] = {"detail": 1}
    if allowlist is not None:
        kwargs["allowlist"] = allowlist
    readings = _ocr_readtext(langs, pixels, **kwargs)
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
    """
    operator = params.get("operator", "gt")
    target_value = params.get("target_value", 0)
    range_min = params.get("range_min")
    range_max = params.get("range_max")
    ocr_min_conf = _effective_ocr_confidence_threshold(
        params.get("ocr_confidence_threshold")
    )
    matched_number: float | None = None
    matched_conf = 0.0
    for _, text, conf in readings:
        cleaned = text.replace(",", "")
        for match in _NUMBERS_RE.findall(cleaned):
            num = float(match)
            if _number_matches(num, operator, target_value, range_min, range_max):
                if matched_number is None or conf > matched_conf:
                    matched_number = num
                    matched_conf = conf
    detail: dict[str, Any] = {"confidence": round(matched_conf, 4)}
    if matched_number is not None:
        detail["number_found"] = matched_number
    return matched_number is not None and matched_conf >= ocr_min_conf, detail


def run_calibration_ocr(
    frame: np.ndarray,
    region: dict[str, Any],
    tool_type: str,
    params: dict[str, Any],
) -> list[Any]:
    """Run EasyOCR for one calibration frame/region (text or numbers tool).

    Public entry point for the server's per-pin OCR cache: returns the raw
    ``(bbox, text, conf)`` readings so they can be memoized and re-scored under
    changed fuzzy/confidence settings without re-running OCR.
    """
    languages = params.get("languages") or ["en"]
    allowlist = (
        _numbers_ocr_allowlist(languages, params.get("integers_only", False))
        if tool_type == "numbers"
        else None
    )
    pixels = extract_region(frame, region)
    return _ocr_region_readings(
        pixels,
        languages=languages,
        allowlist=allowlist,
        preprocess=params.get("ocr_preprocess", False),
        mask_points=region.get("mask_points"),
    )
