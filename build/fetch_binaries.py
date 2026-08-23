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

  llama.cpp (both platforms): ggml-org/llama.cpp GitHub release b10588 (MIT),
    https://github.com/ggml-org/llama.cpp/releases/tag/b10588 — the exact
    build the router-mode gate validation ran against. Assets:
    llama-b10588-bin-macos-arm64.tar.gz (llama-server + its dylib closure,
    Metal + CPU backends) and llama-b10588-bin-win-vulkan-x64.zip
    (llama-server.exe + DLLs; Vulkan for vendor-agnostic GPU, with the
    ggml-cpu-* per-arch DLLs as the no-Vulkan fallback llama.cpp picks at
    runtime). Release assets are permanent; member hashes were computed from
    the archives at pin time.

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
import tarfile
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
        {
            "url": "https://github.com/ggml-org/llama.cpp/releases/download/b10588/llama-b10588-bin-macos-arm64.tar.gz",
            "sha256": "239b46e4b9f537a811fb0c8bc34cd282acee05c4bec49d1d022e66680107b66a",
            "members": {
                "llama-b10588/llama-server": {
                    "target": "llama-server",
                    "sha256": "9f34137ff2559c40a0fe9b130937fd745726130074173665127d863846152198",
                },
                "llama-b10588/libllama-server-impl.dylib": {
                    "target": "libllama-server-impl.dylib",
                    "sha256": "3e1fdf04f0aa51ad6b9f638378e4791a1fd842efeb8e8dc1a6a2ba779b779710",
                },
                "llama-b10588/libllama-common.0.2.0.dylib": {
                    "target": "libllama-common.0.dylib",
                    "sha256": "bb8a1a8a874f66bcd87ddafa9c186ee7137384db1e307a47c12a92b3cee803a1",
                },
                "llama-b10588/libmtmd.0.2.0.dylib": {
                    "target": "libmtmd.0.dylib",
                    "sha256": "63a098e14deeee314ab4a21acf3e6528810b2df7d7c5f1a0bdae2a95ffccd91f",
                },
                "llama-b10588/libllama.0.2.0.dylib": {
                    "target": "libllama.0.dylib",
                    "sha256": "60ef12f0e628dfeaf707ca11900097ff27e609827a7827dceda127569b887a67",
                },
                "llama-b10588/libggml.0.21.0.dylib": {
                    "target": "libggml.0.dylib",
                    "sha256": "ff6664f0778261f54a3bca59a51a05ed5b37aa1a3bafbab8a57a21030fed7e4a",
                },
                "llama-b10588/libggml-cpu.0.21.0.dylib": {
                    "target": "libggml-cpu.0.dylib",
                    "sha256": "af74bad18ff0600e13eca2e01592e66cbf129091f6a00b36b7ad985212dcce55",
                },
                "llama-b10588/libggml-blas.0.21.0.dylib": {
                    "target": "libggml-blas.0.dylib",
                    "sha256": "3bc6de095906be6121fb109d48fd177f90b1aec9e0c51d96a6a65e5c2166a51f",
                },
                "llama-b10588/libggml-metal.0.21.0.dylib": {
                    "target": "libggml-metal.0.dylib",
                    "sha256": "0cf2cf6a5d955524bbcd5087f855d2fbb22ff187bb80aa22f41743fb22f819d1",
                },
                "llama-b10588/libggml-rpc.0.21.0.dylib": {
                    "target": "libggml-rpc.0.dylib",
                    "sha256": "af6a9ab1db29b6e7d5c421eb39ab6f3da2b6e0b07ec3446c012624ef4b0a3059",
                },
                "llama-b10588/libggml-base.0.21.0.dylib": {
                    "target": "libggml-base.0.dylib",
                    "sha256": "1cc45592cb8d243811ff118b23805ad206c94b261fdbd001fee4d441563b4058",
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
        {
            "url": "https://github.com/ggml-org/llama.cpp/releases/download/b10588/llama-b10588-bin-win-vulkan-x64.zip",
            "sha256": "906af65e149b4890d174969875ff6fa2c6e319518cb745b83413706d5fd2d1f5",
            "members": {
                "ggml-base.dll": {
                    "target": "ggml-base.dll",
                    "sha256": "c11a101013510dc4a8dc541d31c8b1e05ff54cbfdad334bf80a46e4ff87a8fd4",
                },
                "ggml-cpu-alderlake.dll": {
                    "target": "ggml-cpu-alderlake.dll",
                    "sha256": "4dbb608492bc0cead7df9bb247b2ecac898ae9b2d1c037276db0fb9753411b57",
                },
                "ggml-cpu-cannonlake.dll": {
                    "target": "ggml-cpu-cannonlake.dll",
                    "sha256": "9bd9097d731f3bcedb9a26d95c4ea33b2388ce9a233215b0501d119980e90ecb",
                },
                "ggml-cpu-cascadelake.dll": {
                    "target": "ggml-cpu-cascadelake.dll",
                    "sha256": "e89212731aceaa4cf099b52dae962071aec2ad7bf2892c5806531a446f9e80ec",
                },
                "ggml-cpu-cooperlake.dll": {
                    "target": "ggml-cpu-cooperlake.dll",
                    "sha256": "953eb64ced873b4016fb832fc4801fe0409aac475cc88c6e09399484303e45fd",
                },
                "ggml-cpu-haswell.dll": {
                    "target": "ggml-cpu-haswell.dll",
                    "sha256": "4dfa0acb6affa1efa556d043e94cd208af29aa4810e63670089e735503de301d",
                },
                "ggml-cpu-icelake.dll": {
                    "target": "ggml-cpu-icelake.dll",
                    "sha256": "d2b89cd49c3e7f90a0c3a6fc4568ae566b638a10d8f26ac72b908bc04306877a",
                },
                "ggml-cpu-ivybridge.dll": {
                    "target": "ggml-cpu-ivybridge.dll",
                    "sha256": "067bd5266190a81430c1444af918c735c287df430e77c53f8f43a35fb8c36203",
                },
                "ggml-cpu-piledriver.dll": {
                    "target": "ggml-cpu-piledriver.dll",
                    "sha256": "691a38b555ed6c03579f5cd8163d51e8577d16e39005d5773fd4e6912fdbe5f4",
                },
                "ggml-cpu-sandybridge.dll": {
                    "target": "ggml-cpu-sandybridge.dll",
                    "sha256": "6754b1e103aa3788b8d3c93caab27abc06719d2d28c8291d1d6597d815766442",
                },
                "ggml-cpu-sapphirerapids.dll": {
                    "target": "ggml-cpu-sapphirerapids.dll",
                    "sha256": "b72b53f24d901ee3df87758f6364fd4337b73d0097202e0f74580862fa980165",
                },
                "ggml-cpu-skylakex.dll": {
                    "target": "ggml-cpu-skylakex.dll",
                    "sha256": "c65fd4ceadc76c28d44319a5a3dd92d4bf3ce2a2dfd2b656cf66b5c029acdda1",
                },
                "ggml-cpu-sse42.dll": {
                    "target": "ggml-cpu-sse42.dll",
                    "sha256": "1024c81019618f86afe2e0cf4d6c2831bb8924ace09412fd1c359b81e6890845",
                },
                "ggml-cpu-x64.dll": {
                    "target": "ggml-cpu-x64.dll",
                    "sha256": "272ed607d294c4c544a42307c8e199e8cfe3a897e1db6e304628e8a836634f0e",
                },
                "ggml-cpu-zen4.dll": {
                    "target": "ggml-cpu-zen4.dll",
                    "sha256": "70c4cb299c8a1119db24173d849b51ab33cf838c98349ad422a799cc719bbd3b",
                },
                "ggml-rpc.dll": {
                    "target": "ggml-rpc.dll",
                    "sha256": "1aa925a4e436e437a4b48c493ef62a0daae299efdf966c8c824a9eeca5e40209",
                },
                "ggml-vulkan.dll": {
                    "target": "ggml-vulkan.dll",
                    "sha256": "0fefae7ed4c19f08788184052578d2e85e092b4653d44393ec5d13bb1ebe4abf",
                },
                "ggml.dll": {
                    "target": "ggml.dll",
                    "sha256": "b7feb6fba7e63377233afd9146b0ca0b1e36369e7a13a12896975e5bfb992f65",
                },
                "libomp.dll": {
                    "target": "libomp.dll",
                    "sha256": "a12116ba72d1d6820407cf30be23da04ce79d6bb8a71a5ee71759c5a1faa6f1c",
                },
                "llama-common.dll": {
                    "target": "llama-common.dll",
                    "sha256": "c94cab2feb7f9a9a5b191619848f7c47f946da88961e0d09917e9e275cd10860",
                },
                "llama-server-impl.dll": {
                    "target": "llama-server-impl.dll",
                    "sha256": "4344686908c4d481349443e5d8643ba532fe4a9c74319ce1d4c621096cd211e5",
                },
                "llama-server.exe": {
                    "target": "llama-server.exe",
                    "sha256": "9a0d241988d0ccc89e7062be4779923f5cee6f6621e9175fe6ed67d2201e0c99",
                },
                "llama.dll": {
                    "target": "llama.dll",
                    "sha256": "f421778948695b1b20b8f091fd92cffabbb3f3e612a08287d4125f63ffbf4b68",
                },
                "mtmd.dll": {
                    "target": "mtmd.dll",
                    "sha256": "d536d0534d73742bfef632a5a78b3f601ee97a081a8d2341ce6bc5dc42c9238e",
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


def _open_member(bundle, member_name: str):
    """One member as a readable file object, from a zip or tar bundle."""
    if isinstance(bundle, zipfile.ZipFile):
        return bundle.open(member_name)
    handle = bundle.extractfile(member_name)
    if handle is None:
        raise SystemExit(
            f"fetch_binaries: {member_name} is not a regular file in the archive"
        )
    return handle


def extract_members(
    archive_path: Path, members: dict[str, dict], vendor_bin: Path
) -> None:
    opener = (
        tarfile.open
        if archive_path.name.endswith((".tar.gz", ".tgz"))
        else zipfile.ZipFile
    )
    with opener(archive_path) as bundle:
        for member_name, member in members.items():
            target = vendor_bin / member["target"]
            digest = hashlib.sha256()
            with _open_member(bundle, member_name) as src, target.open("wb") as out:
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
            suffix = (
                ".tar.gz" if archive["url"].endswith((".tar.gz", ".tgz")) else ".zip"
            )
            with tempfile.NamedTemporaryFile(
                dir=vendor_bin, suffix=suffix, delete=False
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
