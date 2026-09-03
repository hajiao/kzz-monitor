from __future__ import annotations

import base64
import os
import sys
from pathlib import Path
from typing import Any


def save_secret(path: Path, secret: str) -> None:
    if sys.platform == "darwin":
        from .secrets_macos import save_keychain_secret

        save_keychain_secret(secret)
        return
    if os.name != "nt":
        raise RuntimeError("当前系统不支持安全保存 SMTP 授权码")
    _save_windows_secret(path, secret)


def load_secret(path: Path) -> str:
    if sys.platform == "darwin":
        from .secrets_macos import load_keychain_secret

        return load_keychain_secret()
    if os.name != "nt":
        return ""
    return _load_windows_secret(path)


def _windows_api() -> tuple[Any, Any, Any, Any]:
    import ctypes
    from ctypes import wintypes

    class DataBlob(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]

    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    crypt32.CryptProtectData.restype = wintypes.BOOL
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    return ctypes, DataBlob, crypt32, kernel32


def _blob(ctypes: Any, blob_type: Any, data: bytes) -> tuple[Any, Any]:
    buffer = ctypes.create_string_buffer(data)
    return blob_type(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))), buffer


def _save_windows_secret(path: Path, secret: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not secret:
        path.unlink(missing_ok=True)
        return
    ctypes, blob_type, crypt32, kernel32 = _windows_api()
    source, source_buffer = _blob(ctypes, blob_type, secret.encode("utf-8"))
    output = blob_type()
    if not crypt32.CryptProtectData(
        ctypes.byref(source), "KzzMonitor SMTP", None, None, None, 0, ctypes.byref(output)
    ):
        raise ctypes.WinError()
    try:
        path.write_bytes(base64.b64encode(ctypes.string_at(output.pbData, output.cbData)))
    finally:
        kernel32.LocalFree(output.pbData)
        del source_buffer


def _load_windows_secret(path: Path) -> str:
    if not path.exists():
        return ""
    ctypes, blob_type, crypt32, kernel32 = _windows_api()
    source, source_buffer = _blob(ctypes, blob_type, base64.b64decode(path.read_bytes()))
    output = blob_type()
    if not crypt32.CryptUnprotectData(
        ctypes.byref(source), None, None, None, None, 0, ctypes.byref(output)
    ):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(output.pbData, output.cbData).decode("utf-8")
    finally:
        kernel32.LocalFree(output.pbData)
        del source_buffer
