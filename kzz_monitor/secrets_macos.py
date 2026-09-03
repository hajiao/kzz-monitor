from __future__ import annotations

import getpass
import subprocess

SERVICE = "KzzMonitor SMTP"


def save_keychain_secret(secret: str) -> None:
    account = getpass.getuser()
    if not secret:
        subprocess.run(
            ["security", "delete-generic-password", "-s", SERVICE, "-a", account],
            capture_output=True,
            check=False,
        )
        return
    subprocess.run(
        ["security", "add-generic-password", "-U", "-s", SERVICE, "-a", account, "-w", secret],
        capture_output=True,
        text=True,
        check=True,
    )


def load_keychain_secret() -> str:
    result = subprocess.run(
        ["security", "find-generic-password", "-s", SERVICE, "-a", getpass.getuser(), "-w"],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.stdout.strip() if result.returncode == 0 else ""
