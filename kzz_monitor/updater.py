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
    log_path = base / "logs" / "update.log"
    script.write_text(
        "param([int]$ProcessId,[string]$Payload,[string]$TargetDir,[string]$TargetExe,[string]$LogPath)\n"
        "$ErrorActionPreference='Stop'\n"
        "New-Item -ItemType Directory -Force (Split-Path $LogPath) | Out-Null\n"
        "function Write-UpdateLog([string]$Message) { Add-Content -Path $LogPath -Encoding UTF8 "
        "-Value \"$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message\" }\n"
        "try {\n"
        "  Write-UpdateLog \"开始更新，等待旧进程 $ProcessId 退出\"\n"
        "  Wait-Process -Id $ProcessId -ErrorAction SilentlyContinue\n"
        "  for ($i=0; $i -lt 30; $i++) {\n"
        "    $old = Get-CimInstance Win32_Process -Filter \"Name='KzzMonitor.exe'\" -ErrorAction SilentlyContinue | "
        "Where-Object { $_.ExecutablePath -eq $TargetExe }\n"
        "    if (-not $old) { break }; Start-Sleep -Milliseconds 500\n"
        "  }\n"
        "  $copied=$false\n"
        "  for ($i=1; $i -le 20; $i++) {\n"
        "    try { Copy-Item (Join-Path $Payload 'KzzMonitor.exe') $TargetExe -Force; $copied=$true; break } "
        "catch { Write-UpdateLog \"第 $i 次替换失败：$($_.Exception.Message)\"; Start-Sleep -Seconds 1 }\n"
        "  }\n"
        "  if (-not $copied) { throw '20 次尝试后仍无法替换 KzzMonitor.exe' }\n"
        "  Get-ChildItem $Payload -File | Where-Object Name -ne 'KzzMonitor.exe' | ForEach-Object { "
        "try { Copy-Item $_.FullName (Join-Path $TargetDir $_.Name) -Force } catch { "
        "Write-UpdateLog \"手册更新失败但程序继续：$($_.Exception.Message)\" } }\n"
        "  Write-UpdateLog '程序文件替换成功，准备重启'\n"
        "  Start-Sleep -Seconds 2\n"
        "  $started=$false\n"
        "  for ($i=1; $i -le 3; $i++) {\n"
        "    $process=Start-Process -FilePath $TargetExe -WorkingDirectory $TargetDir -PassThru\n"
        "    Start-Sleep -Seconds 5\n"
        "    if (-not $process.HasExited) { $started=$true; Write-UpdateLog \"重启成功，PID=$($process.Id)\"; break }\n"
        "    Write-UpdateLog \"第 $i 次重启后进程提前退出，退出码=$($process.ExitCode)\"\n"
        "    Start-Sleep -Seconds 2\n"
        "  }\n"
        "  if (-not $started) { throw '更新成功，但三次重启均失败；请手动启动 KzzMonitor.exe' }\n"
        "} catch { Write-UpdateLog \"更新失败：$($_.Exception.Message)\" }\n"
        "Start-Sleep -Seconds 2\n"
        "Remove-Item (Split-Path $Payload -Parent) -Recurse -Force -ErrorAction SilentlyContinue\n",
        encoding="utf-8-sig",
    )
    subprocess.Popen(
        [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
            "-ProcessId", str(os.getpid()), "-Payload", str(payload), "-TargetDir", str(base),
            "-TargetExe", str(target_exe), "-LogPath", str(log_path),
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
