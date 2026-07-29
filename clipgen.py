"""clipgen entry point — `uv run clipgen.py`.

Every module lives in `source/`. This file exists only to put that directory on
`sys.path`, so the flat module names (`config`, `utils`, `cli`, …) keep
resolving exactly as they did when they sat in the repo root, and then to hand
off to `cli.main`.

Two details here are important:

* The path is **absolute and resolved**. `cli.main` calls
  `os.chdir(get_runtime_working_dir())` early on, and many first-party imports
  are deferred until after that point.
* The insert is **skipped when frozen**. A PyInstaller bundle has no `source/`
  next to the executable — the modules live in its archive — so the entry would
  point at a directory that does not exist.
"""

import sys
from pathlib import Path

if not getattr(sys, "frozen", False):
    sys.path.insert(0, str(Path(__file__).resolve().parent / "source"))

if __name__ == "__main__":
    import utils
    from cli import main

    try:
        main()
    except KeyboardInterrupt:
        utils.info_print("Interrupted by user")
        sys.exit(0)
