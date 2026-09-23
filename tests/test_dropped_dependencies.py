"""Guard against dropped dependencies creeping back in.

imagehash and scikit-image were replaced by in-tree code
(``screenspace_primitives.PHash`` and ``structural_similarity``); scipy and
PyWavelets rode in only as their transitive dependencies (~110 MB unpacked).
Their notice sections are gone from ``build/THIRD-PARTY-LICENSES`` and the
build workflow fails if their directories reappear in the bundle, so a new
import here must be a deliberate decision that restores both — not a habit.
"""

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "source"

DROPPED = ("imagehash", "skimage", "scipy", "pywt")
_IMPORT_RE = re.compile(
    r"^\s*(?:import\s+(?:{names})\b|from\s+(?:{names})[.\s])".format(
        names="|".join(DROPPED)
    ),
    re.MULTILINE,
)


def test_source_does_not_import_dropped_dependencies() -> None:
    offenders = []
    for path in sorted(SOURCE.glob("*.py")):
        match = _IMPORT_RE.search(path.read_text(encoding="utf-8"))
        if match:
            offenders.append(f"{path.name}: {match.group(0).strip()}")
    assert not offenders, (
        "dropped dependencies imported again (see module docstring): "
        + "; ".join(offenders)
    )


def test_windows_installer_deletes_dropped_packages_on_upgrade() -> None:
    """Inno overlays {app}; without InstallDelete, old trees survive upgrades."""
    iss = (ROOT / "build" / "clipgen.iss").read_text(encoding="utf-8")
    assert "[InstallDelete]" in iss
    for name in DROPPED:
        assert f"{{app}}\\lib\\{name}*" in iss, name
