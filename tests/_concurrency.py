"""Thread helpers for tests that arrange a race explicitly instead of sleeping.

``CountingLock`` wraps a real lock and counts entry attempts before blocking,
so a test can wait until every competitor sits at the acquisition boundary.
``OneLockDict`` hands that lock out through a module's ``setdefault`` lookup.
``WAIT`` is a failure bound only; no test should reach it on a passing run.
"""

import threading
from collections.abc import Callable
from typing import Any, Self

WAIT = 5.0


class CountingLock:
    """A real lock that records each entry attempt before acquiring."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._cond = threading.Condition()
        self.attempts = 0

    def __enter__(self) -> Self:
        with self._cond:
            self.attempts += 1
            self._cond.notify_all()
        self._lock.acquire()
        return self

    def __exit__(self, *_exc: object) -> None:
        self._lock.release()

    def wait_attempts(self, count: int, timeout: float = WAIT) -> bool:
        """True once *count* callers have reached the lock."""
        with self._cond:
            return self._cond.wait_for(lambda: self.attempts >= count, timeout)


class OneLockDict(dict):
    """A per-key lock table whose ``setdefault`` always yields *lock*."""

    def __init__(self, lock: CountingLock) -> None:
        super().__init__()
        self.lock = lock

    def setdefault(self, key: Any, default: Any = None) -> Any:
        return super().setdefault(key, self.lock)


def start_threads(
    fn: Callable[[], Any], count: int
) -> tuple[list[threading.Thread], list[Any], list[BaseException]]:
    """Start *count* daemon threads running *fn*; collect results and errors."""
    results: list[Any] = []
    errors: list[BaseException] = []

    def target() -> None:
        try:
            results.append(fn())
        except BaseException as exc:
            errors.append(exc)

    threads = [threading.Thread(target=target, daemon=True) for _ in range(count)]
    for thread in threads:
        thread.start()
    return threads, results, errors


def join_all(threads: list[threading.Thread], timeout: float = WAIT) -> None:
    """Join every thread and fail if any is still running."""
    for thread in threads:
        thread.join(timeout)
    alive = [thread.name for thread in threads if thread.is_alive()]
    assert not alive, f"threads still running: {alive}"
