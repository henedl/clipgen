"""clipgen entry point — `uv run clipgen.py`.

Every module lives in `source/`. This file exists only to put that directory on
`sys.path`, so the flat module names (`config`, `utils`, `cli`, …) keep
resolving exactly as they did when they sat in the repo root, and then to hand
off to `cli.main`.

Three details here are important:

* The path is **absolute and resolved**. `cli.main` calls
  `os.chdir(get_runtime_working_dir())` early on, and many first-party imports
  are deferred until after that point.
* The insert is **skipped when frozen**. A PyInstaller bundle has no `source/`
  next to the executable — the modules live in its archive — so the entry would
  point at a directory that does not exist.
* `multiprocessing.freeze_support()` runs **before any clipgen import**. In a
  frozen app `sys.executable` is the clipgen binary, so anything that touches
  `multiprocessing` (tqdm's RLock inside faster-whisper's transcribe loop, via
  CPython's resource tracker) re-execs it with interpreter-style argv
  (`-B -S -E -s -c '...'`). PyInstaller's multiprocessing runtime hook handles
  that argv shape inside `freeze_support()` and exits — but only if we call it;
  otherwise the child falls through into `cli.main`, where argparse reads `-S`
  as `--severity` and errors out.
"""

import multiprocessing
import sys
from pathlib import Path

if not getattr(sys, "frozen", False):
    sys.path.insert(0, str(Path(__file__).resolve().parent / "source"))

if __name__ == "__main__":
    multiprocessing.freeze_support()

    import utils
    from cli import main

    try:
        main()
    except KeyboardInterrupt:
        utils.info_print("Interrupted by user")
        sys.exit(0)
