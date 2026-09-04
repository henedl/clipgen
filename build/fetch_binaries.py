#!/usr/bin/env python3
"""Fetch the pinned ffmpeg/ffprobe builds and OCR models for the desktop bundle.

Run with `uv run build/fetch_binaries.py` (stdlib-only on purpose: CI runs it
before the project venv matters, and it must never pull dependencies). It
downloads the pinned archives for the host platform, verifies each archive's
SHA256 against PINS below, extracts ffmpeg + ffprobe into
build/vendor/<platform>/bin/, and verifies the extracted binaries' SHA256 too.
It also fetches the pinned RapidOCR recognition models (OCR_MODEL_PINS,
platform-independent) into build/vendor/ocr/ and the speaker-embedding model
(SPEAKER_MODEL_PINS) into build/vendor/speakers/ so the frozen bundle never
downloads a model at runtime. The extracted-file hashes double as the
idempotency check: when the files are already present and hash-clean the
script exits without downloading. clipgen.spec refuses to build if any of
these files are missing.

Provenance (THIRD-PARTY-LICENSES cites this block):
  macos-arm64: Martin Riedl's FFmpeg build server, https://ffmpeg.martin-riedl.de
    Build 1783011502_8.1.2 — FFmpeg 8.1.2 release, GPLv3 build
    (--enable-gpl --enable-version3, with libx264/libvpx/libwebp/libfreetype;
    VideoToolbox enabled by default on macOS). Permanent per-build URLs.
  windows-x64: Gyan Doshi's builds, https://www.gyan.dev/ffmpeg/builds/
    Package ffmpeg-8.1.2-essentials_build.zip — FFmpeg 8.1.2 release, GPLv3
    build (--enable-gpl --enable-version3, with libx264/libvpx/libwebp/
    libfreetype). Versioned packages under /builds/packages/ are permanent;
    the archive hash matches the provider's published .sha256 file. Never pin
    the "release-essentials" / "git-essentials" aliases: those move. (The
    earlier BtbN autobuild pin was dropped because BtbN prunes dated tags
    after ~2 weeks and any edit here rotates the CI cache key.)

  llama.cpp (both platforms): ggml-org/llama.cpp GitHub release b10588 (MIT),
    https://github.com/ggml-org/llama.cpp/releases/tag/b10588 — the exact
    build the router-mode gate validation ran against. Assets:
    llama-b10588-bin-macos-arm64.tar.gz (llama-server + its dylib closure,
    Metal + CPU backends) and llama-b10588-bin-win-vulkan-x64.zip
    (llama-server.exe + DLLs; Vulkan for vendor-agnostic GPU, with the
    ggml-cpu-* per-arch DLLs as the no-Vulkan fallback llama.cpp picks at
    runtime). Release assets are permanent; member hashes were computed from
    the archives at pin time.

  Speaker model: WeSpeaker CAM++ (VoxCeleb, large-margin fine-tune), the
    ONNX export distributed by k2-fsa/sherpa-onnx under the permanent
    `speaker-recongition-models` release tag (sic, upstream spelling),
    Apache-2.0, 29 MB. GitHub publishes no .sha256 for release assets, so
    the pin trusts the first download. speakers.py computes its own fbank
    and reads the model's metadata, so a re-pin only needs a 16 kHz
    wespeaker/3d-speaker export.

Updating the pins: `uv run build/fetch_binaries.py --repin <new-url>` finds
the entry the URL replaces (same host, same platform), downloads it, checks
the archive against the provider's published `<url>.sha256` when one exists,
re-hashes the members the old entry pinned, and rewrites the entry in place.
`--check-urls` probes every pinned URL (the weekly pin-health workflow runs
it). The full procedure — provenance notes, THIRD-PARTY-LICENSES, CI
dispatch, launching the frozen app — is agents/skills/bump-pins/SKILL.md.
"""

import argparse
import hashlib
import platform
import re
import sys
import tarfile
import tempfile
import urllib.error
import urllib.request
import zipfile
from pathlib import Path
from urllib.parse import urlsplit

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
            "url": "https://www.gyan.dev/ffmpeg/builds/packages/ffmpeg-8.1.2-essentials_build.zip",
            "sha256": "db580001caa24ac104c8cb856cd113a87b0a443f7bdf47d8c12b1d740584a2ec",
            "members": {
                "ffmpeg-8.1.2-essentials_build/bin/ffmpeg.exe": {
                    "target": "ffmpeg.exe",
                    "sha256": "1326dde4c84ff1f96fe6b8916c5bed29e163e9b5dccf995f6f3db069d143ec5e",
                },
                "ffmpeg-8.1.2-essentials_build/bin/ffprobe.exe": {
                    "target": "ffprobe.exe",
                    "sha256": "b49ccc7c6547b141ad5a2f6ec69cc04323d7133d7704d70b331b904c63eecb07",
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

# Vendored RapidOCR recognition models; re-derive pins per agents/skills/bump-pins/SKILL.md.
# Target names match screenspace_ocr._vendored_rec_model.
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

# Vendored speaker-embedding model; target name matches speakers.MODEL_FILENAME.
SPEAKER_MODEL_PINS: list[dict] = [
    {
        "url": "https://github.com/k2-fsa/sherpa-onnx/releases/download/speaker-recongition-models/wespeaker_en_voxceleb_CAM%2B%2B_LM.onnx",
        "sha256": "e197af7e9d473030cf486b3124149a19bf37014d0e4485e4c70c483b0ec10cb2",
        "target": "speaker_embed.onnx",
    },
]

_CHUNK = 1024 * 1024
_USER_AGENT = "clipgen-fetch-binaries"


def _request(
    url: str, method: str = "GET", headers: dict[str, str] | None = None
) -> urllib.request.Request:
    # martin-riedl.de returns 403 to urllib's default Python-urllib/3.x agent.
    return urllib.request.Request(
        url, method=method, headers={"User-Agent": _USER_AGENT, **(headers or {})}
    )


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
    try:
        with (
            urllib.request.urlopen(_request(url), timeout=60) as response,
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


def _opener(archive_path: Path):
    if archive_path.name.endswith((".tar.gz", ".tgz")):
        return tarfile.open
    return zipfile.ZipFile


def extract_members(
    archive_path: Path, members: dict[str, dict], vendor_bin: Path
) -> None:
    with _opener(archive_path)(archive_path) as bundle:
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


def fetch_model_pins(pins: list[dict], subdir: str) -> None:
    """Fetch pinned single-file models into build/vendor/<subdir>/."""
    vendor_dir = Path(__file__).resolve().parent / "vendor" / subdir
    stale = [
        pin
        for pin in pins
        if not (vendor_dir / pin["target"]).is_file()
        or file_sha256(vendor_dir / pin["target"]) != pin["sha256"]
    ]
    if not stale:
        print(f"fetch_binaries: {vendor_dir} is up to date, nothing to do.")
        return
    vendor_dir.mkdir(parents=True, exist_ok=True)
    for pin in stale:
        # download_archive hashes plain files too and deletes on mismatch.
        download_archive(pin["url"], pin["sha256"], vendor_dir / pin["target"])
    print(f"fetch_binaries: done — {vendor_dir} is ready.")


def all_pinned_urls() -> list[str]:
    """Every pinned URL: both platforms' archives plus the models."""
    urls = [a["url"] for archives in PINS.values() for a in archives]
    return urls + [pin["url"] for pin in OCR_MODEL_PINS + SPEAKER_MODEL_PINS]


def _probe(url: str) -> int:
    """HTTP status for *url*: HEAD, then a one-byte GET for hosts that reject HEAD."""
    status = 0
    for method, headers in (("HEAD", {}), ("GET", {"Range": "bytes=0-0"})):
        try:
            with urllib.request.urlopen(
                _request(url, method, headers), timeout=30
            ) as response:
                return 200 if response.status in (200, 206) else response.status
        except urllib.error.HTTPError as exc:
            status = exc.code
        except OSError:
            status = 0
    return status


def check_urls() -> None:
    """Probe every pinned URL; exit non-zero when any has gone away."""
    urls = all_pinned_urls()
    gone = []
    for url in urls:
        status = _probe(url)
        print(f"fetch_binaries: {status or 'ERR':>3}  {url}")
        if status != 200:
            gone.append(url)
    if gone:
        raise SystemExit(
            f"fetch_binaries: {len(gone)} of {len(urls)} pinned URLs unreachable:\n  "
            + "\n  ".join(gone)
        )
    print(f"fetch_binaries: all {len(urls)} pinned URLs reachable.")


def published_sha256(url: str) -> str | None:
    """The provider's `<url>.sha256` digest, or None when it publishes none."""
    try:
        with urllib.request.urlopen(_request(url + ".sha256"), timeout=30) as resp:
            text = resp.read(4096).decode("utf-8", "replace")
    except OSError:
        return None
    match = re.search(r"\b[0-9a-f]{64}\b", text)
    return match.group(0) if match else None


_PLATFORM_TOKEN = {"macos-arm64": "macos", "windows-x64": "win"}


def _entry_for(url: str) -> dict:
    """The PINS / OCR entry *url* replaces: same host, then platform or family."""
    host = urlsplit(url).netloc
    name = Path(urlsplit(url).path).name.lower()
    archives = [
        (plat, a)
        for plat, entries in PINS.items()
        for a in entries
        if urlsplit(a["url"]).netloc == host
    ]
    if len(archives) > 1:
        archives = [(p, a) for p, a in archives if _PLATFORM_TOKEN[p] in name]
    models = [
        ("ocr", pin)
        for pin in OCR_MODEL_PINS
        if urlsplit(pin["url"]).netloc == host
        and pin["target"].removesuffix("_rec.onnx") in name
    ]
    models += [
        ("speakers", pin)
        for pin in SPEAKER_MODEL_PINS
        if urlsplit(pin["url"]).netloc == host and "speaker" in name
    ]
    candidates = archives + models
    if len(candidates) != 1:
        raise SystemExit(
            f"fetch_binaries: {url} matches {len(candidates)} pinned entries; "
            "expected exactly one (same host, platform token in the filename)."
        )
    return candidates[0][1]


def _member_target(name: str) -> str:
    """Pinned target for an archive member: basename, dylibs sans minor.patch."""
    base = name.rsplit("/", 1)[-1]
    return re.sub(r"^(.+?\.\d+)\.\d+\.\d+\.dylib$", r"\1.dylib", base)


def _archive_files(archive_path: Path) -> list[str]:
    with _opener(archive_path)(archive_path) as bundle:
        if isinstance(bundle, zipfile.ZipFile):
            return [i.filename for i in bundle.infolist() if not i.is_dir()]
        return [m.name for m in bundle.getmembers() if m.isfile()]


def _rehash_members(archive_path: Path, old_members: dict[str, dict]) -> dict:
    """Re-pin each old member target against the new archive's files."""
    by_target = {_member_target(n): n for n in _archive_files(archive_path)}
    missing = [
        m["target"] for m in old_members.values() if m["target"] not in by_target
    ]
    if missing:
        raise SystemExit(
            f"fetch_binaries: {archive_path.name} lacks pinned members {missing} "
            "— the provider's closure changed; edit the entry by hand."
        )
    members: dict[str, dict] = {}
    with _opener(archive_path)(archive_path) as bundle:
        for old in old_members.values():
            name = by_target[old["target"]]
            digest = hashlib.sha256()
            with _open_member(bundle, name) as src:
                while chunk := src.read(_CHUNK):
                    digest.update(chunk)
            members[name] = {"target": old["target"], "sha256": digest.hexdigest()}
    return members


def _render_entry(entry: dict, indent: int) -> list[str]:
    """The entry as source lines in ruff's layout, so the diff stays minimal."""
    pad = " " * indent
    lines = [
        f"{pad}{{\n",
        f'{pad}    "url": "{entry["url"]}",\n',
        f'{pad}    "sha256": "{entry["sha256"]}",\n',
    ]
    if "members" in entry:
        lines.append(f'{pad}    "members": {{\n')
        for name, member in entry["members"].items():
            lines += [
                f'{pad}        "{name}": {{\n',
                f'{pad}            "target": "{member["target"]}",\n',
                f'{pad}            "sha256": "{member["sha256"]}",\n',
                f"{pad}        }},\n",
            ]
        lines.append(f"{pad}    }},\n")
    else:
        lines.append(f'{pad}    "target": "{entry["target"]}",\n')
    lines.append(f"{pad}}},\n")
    return lines


def _replace_entry(old_url: str, entry: dict) -> None:
    """Rewrite the PINS / OCR entry holding *old_url* in this file."""
    source = Path(__file__)
    lines = source.read_text(encoding="utf-8").splitlines(keepends=True)
    url_line = next(i for i, line in enumerate(lines) if f'"url": "{old_url}"' in line)
    start = url_line - 1
    assert lines[start].strip() == "{", lines[start]
    depth = 0
    for end in range(start, len(lines)):
        depth += lines[end].count("{") - lines[end].count("}")
        if depth == 0:
            break
    indent = len(lines[start]) - len(lines[start].lstrip())
    lines[start : end + 1] = _render_entry(entry, indent)
    source.write_text("".join(lines), encoding="utf-8")


def repin(url: str) -> None:
    """Point the entry *url* replaces at *url*, with freshly computed hashes."""
    old = _entry_for(url)
    print(f"fetch_binaries: re-pinning {old['url']}\n  -> {url}")
    with tempfile.TemporaryDirectory() as tmp:
        archive_path = Path(tmp) / Path(urlsplit(url).path).name
        print(f"fetch_binaries: downloading {url}")
        try:
            with (
                urllib.request.urlopen(_request(url), timeout=60) as response,
                archive_path.open("wb") as out,
            ):
                while chunk := response.read(_CHUNK):
                    out.write(chunk)
        except OSError as exc:
            raise SystemExit(
                f"fetch_binaries: download failed for {url}: {exc}"
            ) from exc
        sha256 = file_sha256(archive_path)
        published = published_sha256(url)
        if published is None:
            print(
                "fetch_binaries: provider publishes no .sha256 — trusting first download"
            )
        elif published != sha256:
            raise SystemExit(
                f"fetch_binaries: {url} does not match the provider's .sha256\n"
                f"  published {published}\n  received  {sha256}"
            )
        else:
            print("fetch_binaries: archive matches the provider's published .sha256")
        entry = {"url": url, "sha256": sha256}
        if "members" in old:
            entry["members"] = _rehash_members(archive_path, old["members"])
        else:
            entry["target"] = old["target"]
    _replace_entry(old["url"], entry)
    print(
        f"fetch_binaries: rewrote {Path(__file__).name}. Next: update the provenance "
        "docstring and build/THIRD-PARTY-LICENSES, clear build/vendor/, rerun "
        "this script, dispatch build-binaries."
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    parser.add_argument(
        "--repin",
        metavar="URL",
        nargs="+",
        help="re-pin the entries these URLs replace",
    )
    parser.add_argument(
        "--check-urls", action="store_true", help="probe every pinned URL and exit"
    )
    args = parser.parse_args()
    if args.check_urls:
        check_urls()
        return
    if args.repin:
        for url in args.repin:
            repin(url)
        return
    fetch_vendor()


def fetch_vendor() -> None:
    plat = host_platform()
    archives = PINS[plat]
    vendor_bin = Path(__file__).resolve().parent / "vendor" / plat / "bin"
    if vendor_up_to_date(vendor_bin, archives):
        print(f"fetch_binaries: {vendor_bin} is up to date, nothing to do.")
    else:
        vendor_bin.mkdir(parents=True, exist_ok=True)
        for archive in archives:
            # Temp file beside the targets keeps the write on one filesystem.
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
    fetch_model_pins(OCR_MODEL_PINS, "ocr")
    fetch_model_pins(SPEAKER_MODEL_PINS, "speakers")


if __name__ == "__main__":
    main()
