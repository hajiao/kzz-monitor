from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

APP_NAME = "KzzMonitor"


def is_macos() -> bool:
    return sys.platform == "darwin"


def user_data_dir() -> Path:
    if is_macos():
        return Path.home() / "Library" / "Application Support" / APP_NAME
    if os.name == "nt":
        # Windows 继续使用便携目录，方便连同 Excel 一起复制。
        return executable_dir()
    return Path.home() / ".local" / "share" / APP_NAME


def instance_lock_path() -> Path:
    """返回当前用户全局唯一的锁；即使复制多份程序也只允许运行一个实例。"""
    if os.name == "nt":
        root = Path(os.getenv("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        return root / APP_NAME / "kzz_monitor.lock"
    if is_macos():
        return Path.home() / "Library" / "Application Support" / APP_NAME / "kzz_monitor.lock"
    return Path.home() / ".local" / "share" / APP_NAME / "kzz_monitor.lock"


def show_already_running(message: str) -> None:
    if os.name == "nt":
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, message, "KzzMonitor", 0x40)
    elif is_macos():
        escaped = message.replace("\\", "\\\\").replace('"', '\\"')
        subprocess.run(
            ["osascript", "-e", f'display alert "KzzMonitor" message "{escaped}"'],
            capture_output=True,
            check=False,
        )
    else:
        print(message, file=sys.stderr)


def executable_dir() -> Path:
    if getattr(sys, "frozen", False):
        executable = Path(sys.executable).resolve()
        if is_macos() and ".app" in str(executable):
            # Foo.app/Contents/MacOS/Foo -> 包含 Foo.app 的目录
            return executable.parents[3]
        return executable.parent
    return Path(__file__).resolve().parents[1]


def bundled_resource(name: str) -> Path | None:
    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        candidate = Path(bundle_root) / name
        return candidate if candidate.exists() else None
    candidate = Path(__file__).resolve().parents[1] / name
    return candidate if candidate.exists() else None


def open_path(path: Path) -> None:
    if os.name == "nt":
        os.startfile(path)  # type: ignore[attr-defined]
    elif is_macos():
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])
