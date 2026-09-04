from __future__ import annotations

import threading
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
