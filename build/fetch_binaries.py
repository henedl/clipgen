#!/usr/bin/env python3
"""Fetch the pinned ffmpeg/ffprobe builds and OCR models for the desktop bundle.

Run with `uv run build/fetch_binaries.py` (stdlib-only on purpose: CI runs it
before the project venv matters, and it must never pull dependencies). It
downloads the pinned archives for the host platform, verifies each archive's
SHA256 against PINS below, extracts ffmpeg + ffprobe into
build/vendor/<platform>/bin/, and verifies the extracted binaries' SHA256 too.
It also fetches the pinned RapidOCR recognition models (OCR_MODEL_PINS,
platform-independent) into build/vendor/ocr/ so the frozen bundle never
downloads a model at runtime. The extracted-file hashes double as the
idempotency check: when the files are already present and hash-clean the
script exits without downloading. clipgen.spec refuses to build if any of
these files are missing.

Provenance (THIRD-PARTY-LICENSES cites this block):
  macos-arm64: Martin Riedl's FFmpeg build server, https://ffmpeg.martin-riedl.de
    Build 1783011502_8.1.2 — FFmpeg 8.1.2 release, GPLv3 build
    (--enable-gpl --enable-version3, with libx264/libvpx/libwebp/libfreetype;
    VideoToolbox enabled by default on macOS). Permanent per-build URLs.
  windows-x64: BtbN FFmpeg-Builds, https://github.com/BtbN/FFmpeg-Builds
    Release tag autobuild-2026-08-16-13-00 (a dated tag — never pin "latest",
    its assets are replaced daily), asset
    ffmpeg-n8.1.2-44-g7c533d0f86-win64-gpl-8.1.zip — FFmpeg 8.1 release
    branch, GPLv3 build. Archive hash cross-checked against the release's
    published checksums.sha256.
    CAUTION: BtbN prunes dated autobuild tags after roughly two weeks, so
    this pin has a shelf life. CI usually survives on the build/vendor cache,
    but any change to this file rotates the cache key and forces a real
    download — so refresh this pin (tag, asset name, all three hashes)
    whenever you touch this file, and expect a 404 here if the pin has
    lapsed.

Updating the pins: pick a new build/tag on the provider, update the archive
URL + sha256 (from the provider's published .sha256 / checksums.sha256 files),
clear build/vendor/, run this script — it prints the extracted-file hashes on
mismatch so the member sha256s can be copied in — then let CI's post-build
feature check (libx264/libwebp/libvpx-vp9/drawtext) confirm the new build
still carries everything clipgen's soft gates probe for.
"""

import hashlib
import platform
import sys
import tempfile
import urllib.request
import zipfile
from pathlib import Path

PINS: dict[str, list[dict]] = {
    "macos-arm64": [
        {
            "url": "https://ffmpeg.martin-riedl.de/download/macos/arm64/1783011502_8.1.2/ffmpeg.zip",
            "sha256": "ef1aa60006c7b77ce170c1608c08d8e4ba1c30c5746f2ac986ded932d0ac2c3c",
            "members": {
                "ffmpeg": {
                    "target": "ffmpeg",
                    "sha256": "eaf91238e104dd0e262bc6510e25061855cc99a6955a721b0ac99660d58c473d",
                },
            },
        },
        {
            "url": "https://ffmpeg.martin-riedl.de/download/macos/arm64/1783011502_8.1.2/ffprobe.zip",
            "sha256": "c39787f4af7a3932502d2d48db6f6feaaa836b48a73ef78c32cc3285df61dfaf",
            "members": {
                "ffprobe": {
                    "target": "ffprobe",
                    "sha256": "ed9dc5871914b466b96b402c9ec0ba68ce4f836e72faa464b1b4e279835bd4a6",
                },
            },
        },
    ],
    "windows-x64": [
        {
            "url": "https://github.com/BtbN/FFmpeg-Builds/releases/download/autobuild-2026-08-16-13-00/ffmpeg-n8.1.2-44-g7c533d0f86-win64-gpl-8.1.zip",
            "sha256": "d2425b12dc746a2b044148c6100440d4065876ac4ed6e3eb13a68437b7719796",
            "members": {
                "ffmpeg-n8.1.2-44-g7c533d0f86-win64-gpl-8.1/bin/ffmpeg.exe": {
                    "target": "ffmpeg.exe",
                    "sha256": "361c161fa536922d32badacad5f32fbd8561f945bd6e0bf4cb1f1017deb03541",
                },
                "ffmpeg-n8.1.2-44-g7c533d0f86-win64-gpl-8.1/bin/ffprobe.exe": {
                    "target": "ffprobe.exe",
                    "sha256": "9fe11967029cff5562e7b0c2e74987690bea1b8b877fcadd6c9939a334fc9fb3",
                },
            },
        },
    ],
}

# RapidOCR PP-OCR recognition models (Apache-2.0), vendored so the frozen
# bundle never downloads at runtime. Detection/orientation models ship inside
# the rapidocr wheel; only the non-default recognition models are fetched here.
# ONNX rec models embed their character dict in the model metadata, so one
# .onnx per script family is the whole artifact. URLs and SHA256 come from the
# pinned rapidocr wheel's default_models.yaml (onnxruntime section) —
# re-derive both on any rapidocr version bump. Target names match
# screenspace_ocr._vendored_rec_model. Japan has no PP-OCRv5 ONNX model
# upstream, hence v4 there.
OCR_MODEL_PINS: list[dict] = [
    {
        "url": "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.9.2/onnx/PP-OCRv5/rec/latin_PP-OCRv5_rec_mobile.onnx",
        "sha256": "b20bd37c168a570f583afbc8cd7925603890efbcdc000a59e22c269d160b5f5a",
        "target": "latin_rec.onnx",
    },
    {
        "url": "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.9.2/onnx/PP-OCRv4/rec/japan_PP-OCRv4_rec_mobile.onnx",
        "sha256": "e1075a67dba758ecfc7ebc78a10ae61c95ac8fb66a9c86fab5541e33f085cb7a",
        "target": "japan_rec.onnx",
    },
    {
        "url": "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.9.2/onnx/PP-OCRv5/rec/korean_PP-OCRv5_rec_mobile.onnx",
        "sha256": "cd6e2ea50f6943ca7271eb8c56a877a5a90720b7047fe9c41a2e541a25773c9b",
        "target": "korean_rec.onnx",
    },
]

_CHUNK = 1024 * 1024


def host_platform() -> str:
    """Map the host to a PINS key, refusing hosts with no pinned build."""
    if sys.platform == "darwin":
        if platform.machine() != "arm64":
            raise SystemExit(
                "fetch_binaries: no pinned ffmpeg build for macOS "
                f"{platform.machine()} — only arm64 is pinned."
            )
        return "macos-arm64"
    if sys.platform == "win32":
        return "windows-x64"
    raise SystemExit(
        f"fetch_binaries: no pinned ffmpeg build for {sys.platform}. The "
        "desktop bundle is only built for macOS and Windows."
    )


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(_CHUNK):
            digest.update(chunk)
    return digest.hexdigest()


def vendor_up_to_date(vendor_bin: Path, archives: list[dict]) -> bool:
    for archive in archives:
        for member in archive["members"].values():
            target = vendor_bin / member["target"]
            if not target.is_file() or file_sha256(target) != member["sha256"]:
                return False
    return True


def download_archive(url: str, expected_sha256: str, dest: Path) -> None:
    """Stream *url* to *dest*, hashing as it downloads; delete on mismatch."""
    print(f"fetch_binaries: downloading {url}")
    digest = hashlib.sha256()
    received = 0
    # martin-riedl.de returns 403 to urllib's default Python-urllib/3.x agent.
    request = urllib.request.Request(
        url, headers={"User-Agent": "clipgen-fetch-binaries"}
    )
    try:
        with (
            urllib.request.urlopen(request, timeout=60) as response,
            dest.open("wb") as out,
        ):
            while chunk := response.read(_CHUNK):
                digest.update(chunk)
                out.write(chunk)
                received += len(chunk)
    except OSError as exc:
        dest.unlink(missing_ok=True)
        raise SystemExit(f"fetch_binaries: download failed for {url}: {exc}") from exc
    if digest.hexdigest() != expected_sha256:
        dest.unlink(missing_ok=True)
        raise SystemExit(
            f"fetch_binaries: SHA256 mismatch for {url}\n"
            f"  expected {expected_sha256}\n"
            f"  received {digest.hexdigest()}\n"
            "The pinned archive changed upstream. Do not bypass this check — "
            "re-verify the pin against the provider's published checksums."
        )
    print(f"fetch_binaries: verified archive ({received / 1e6:.0f} MB)")


def extract_members(
    archive_path: Path, members: dict[str, dict], vendor_bin: Path
) -> None:
    with zipfile.ZipFile(archive_path) as bundle:
        for member_name, member in members.items():
            target = vendor_bin / member["target"]
            digest = hashlib.sha256()
            with bundle.open(member_name) as src, target.open("wb") as out:
                while chunk := src.read(_CHUNK):
                    digest.update(chunk)
                    out.write(chunk)
            if digest.hexdigest() != member["sha256"]:
                target.unlink(missing_ok=True)
                raise SystemExit(
                    f"fetch_binaries: SHA256 mismatch for extracted {member_name}\n"
                    f"  expected {member['sha256']}\n"
                    f"  received {digest.hexdigest()}"
                )
            if sys.platform != "win32":
                target.chmod(0o755)
            print(f"fetch_binaries: extracted {target}")


def fetch_ocr_models() -> None:
    """Fetch the pinned RapidOCR recognition models into build/vendor/ocr/."""
    vendor_ocr = Path(__file__).resolve().parent / "vendor" / "ocr"
    stale = [
        pin
        for pin in OCR_MODEL_PINS
        if not (vendor_ocr / pin["target"]).is_file()
        or file_sha256(vendor_ocr / pin["target"]) != pin["sha256"]
    ]
    if not stale:
        print(f"fetch_binaries: {vendor_ocr} is up to date, nothing to do.")
        return
    vendor_ocr.mkdir(parents=True, exist_ok=True)
    for pin in stale:
        # download_archive streams + hashes plain files just as well as zips,
        # and deletes on mismatch, so a failure leaves nothing half-written.
        download_archive(pin["url"], pin["sha256"], vendor_ocr / pin["target"])
    print(f"fetch_binaries: done — {vendor_ocr} is ready.")


def main() -> None:
    plat = host_platform()
    archives = PINS[plat]
    vendor_bin = Path(__file__).resolve().parent / "vendor" / plat / "bin"
    if vendor_up_to_date(vendor_bin, archives):
        print(f"fetch_binaries: {vendor_bin} is up to date, nothing to do.")
    else:
        vendor_bin.mkdir(parents=True, exist_ok=True)
        for archive in archives:
            # Temp file next to the targets so the rename-free write stays on
            # one filesystem, and a hash failure leaves nothing half-extracted.
            with tempfile.NamedTemporaryFile(
                dir=vendor_bin, suffix=".zip", delete=False
            ) as tmp:
                archive_path = Path(tmp.name)
            try:
                download_archive(archive["url"], archive["sha256"], archive_path)
                extract_members(archive_path, archive["members"], vendor_bin)
            finally:
                archive_path.unlink(missing_ok=True)
        print(f"fetch_binaries: done — {vendor_bin} is ready.")
    fetch_ocr_models()


if __name__ == "__main__":
    main()
