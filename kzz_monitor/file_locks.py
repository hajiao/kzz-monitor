from __future__ import annotations

import threading
import time
from pathlib import Path
from typing import Any, Callable, TypeVar, cast

_LOCKS_GUARD = threading.Lock()
_FILE_LOCKS: dict[str, threading.RLock] = {}
F = TypeVar("F", bound=Callable[..., Any])


def workbook_lock(path: Path) -> threading.RLock:
    key = str(path.resolve()).lower()
    with _LOCKS_GUARD:
        return _FILE_LOCKS.setdefault(key, threading.RLock())


def synchronized_path(function: F) -> F:
    def wrapper(path: Path, *args: Any, **kwargs: Any) -> Any:
        with workbook_lock(path):
            return function(path, *args, **kwargs)
    return cast(F, wrapper)


def replace_with_retry(temporary: Path, target: Path, attempts: int = 5) -> None:
    for attempt in range(attempts):
        try:
            temporary.replace(target)
            return
        except PermissionError:
            if attempt == attempts - 1:
                raise
            time.sleep(0.2 * (attempt + 1))


def cleanup_temporary(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except PermissionError:
        pass
