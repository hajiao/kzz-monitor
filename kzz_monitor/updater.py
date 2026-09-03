from __future__ import annotations

import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests

from . import __version__


@dataclass(slots=True)
class UpdateInfo:
    version: str
    url: str
    sha256: str
    notes: str = ""


def _version_tuple(value: str) -> tuple[int, ...]:
    parts: list[int] = []
    for part in value.strip().lstrip("v").split("."):
        digits = "".join(char for char in part if char.isdigit())
        parts.append(int(digits or 0))
    return tuple(parts)


def _platform_key() -> str:
    if os.name == "nt":
        return "windows-x64"
    if sys.platform == "darwin":
        machine = platform.machine().lower()
        return "macos-arm64" if machine in {"arm64", "aarch64"} else "macos-x64"
    raise RuntimeError("当前系统暂不支持自动更新")


def _read_location(location: str, timeout: int = 20) -> bytes:
    parsed = urlparse(location)
    if parsed.scheme in {"http", "https"}:
        response = requests.get(location, timeout=timeout)
        response.raise_for_status()
        return response.content
    if parsed.scheme == "file":
        from urllib.request import url2pathname

        return Path(url2pathname(parsed.path)).read_bytes()
    return Path(location).expanduser().read_bytes()


def _resolve_package_url(manifest_location: str, package_location: str) -> str:
    if urlparse(package_location).scheme or Path(package_location).is_absolute():
        return package_location
    if urlparse(manifest_location).scheme in {"http", "https", "file"}:
        return urljoin(manifest_location, package_location)
    return str((Path(manifest_location).expanduser().parent / package_location).resolve())


def check_for_update(manifest_location: str) -> UpdateInfo | None:
    if not manifest_location.strip():
        return None
    manifest = json.loads(_read_location(manifest_location).decode("utf-8-sig"))
    latest = str(manifest["version"])
    if _version_tuple(latest) <= _version_tuple(__version__):
        return None
    artifact = manifest["artifacts"][_platform_key()]
    return UpdateInfo(
        version=latest,
        url=_resolve_package_url(manifest_location, str(artifact["url"])),
        sha256=str(artifact["sha256"]).lower(),
        notes=str(manifest.get("notes", "")),
    )


def _safe_extract(archive: Path, destination: Path) -> None:
    with zipfile.ZipFile(archive) as bundle:
        root = destination.resolve()
        for member in bundle.infolist():
            target = (destination / member.filename).resolve()
            if root not in target.parents and target != root:
                raise ValueError("更新包包含不安全路径")
        bundle.extractall(destination)


def stage_and_launch_update(info: UpdateInfo, application_base: Path) -> None:
    if not getattr(sys, "frozen", False):
        raise RuntimeError("源码运行模式不能自动替换程序，请重新构建")
    staging = Path(tempfile.mkdtemp(prefix="KzzMonitor-update-"))
    archive = staging / "update.zip"
    archive.write_bytes(_read_location(info.url, timeout=120))
    actual_hash = hashlib.sha256(archive.read_bytes()).hexdigest().lower()
    if actual_hash != info.sha256:
        shutil.rmtree(staging, ignore_errors=True)
        raise ValueError("更新包 SHA-256 校验失败，已拒绝安装")
    payload = staging / "payload"
    payload.mkdir()
    if sys.platform == "darwin":
        subprocess.run(["ditto", "-x", "-k", str(archive), str(payload)], check=True)
    else:
        _safe_extract(archive, payload)
    if os.name == "nt":
        source_exe = next(payload.rglob("KzzMonitor.exe"), None)
        if source_exe is None:
            raise ValueError("Windows 更新包缺少 KzzMonitor.exe")
        _launch_windows_replacer(staging, source_exe.parent, application_base)
    elif sys.platform == "darwin":
        source_app = next(payload.rglob("KzzMonitor.app"), None)
        if source_app is None:
            raise ValueError("macOS 更新包缺少 KzzMonitor.app")
        _launch_macos_replacer(staging, source_app)
    else:
        raise RuntimeError("当前系统暂不支持自动更新")


def _launch_windows_replacer(staging: Path, payload: Path, base: Path) -> None:
    script = staging / "install-update.ps1"
    target_exe = base / "KzzMonitor.exe"
    script.write_text(
        "param([int]$ProcessId,[string]$Payload,[string]$TargetDir,[string]$TargetExe)\n"
        "$ErrorActionPreference='Stop'\n"
        "Wait-Process -Id $ProcessId -ErrorAction SilentlyContinue\n"
        "Start-Sleep -Milliseconds 800\n"
        "Copy-Item (Join-Path $Payload 'KzzMonitor.exe') $TargetExe -Force\n"
        "Get-ChildItem $Payload -File | Where-Object Name -ne 'KzzMonitor.exe' | "
        "ForEach-Object { Copy-Item $_.FullName (Join-Path $TargetDir $_.Name) -Force }\n"
        "Start-Process $TargetExe\n"
        "Start-Sleep -Seconds 2\n"
        "Remove-Item (Split-Path $Payload -Parent) -Recurse -Force -ErrorAction SilentlyContinue\n",
        encoding="utf-8-sig",
    )
    subprocess.Popen(
        [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
            "-ProcessId", str(os.getpid()), "-Payload", str(payload), "-TargetDir", str(base),
            "-TargetExe", str(target_exe),
        ],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def _launch_macos_replacer(staging: Path, source_app: Path) -> None:
    executable = Path(sys.executable).resolve()
    target_app = next((parent for parent in executable.parents if parent.suffix == ".app"), None)
    if target_app is None:
        raise RuntimeError("无法定位当前 KzzMonitor.app")
    script = staging / "install-update.sh"
    script.write_text(
        "#!/bin/bash\nset -e\n"
        f"PID={os.getpid()}\n"
        "while kill -0 $PID 2>/dev/null; do sleep 1; done\n"
        f"rm -rf {str(target_app)!r}\n"
        f"cp -R {str(source_app)!r} {str(target_app)!r}\n"
        f"open {str(target_app)!r}\n"
        f"rm -rf {str(staging)!r}\n",
        encoding="utf-8",
    )
    script.chmod(0o700)
    subprocess.Popen(["/bin/bash", str(script)], start_new_session=True)
