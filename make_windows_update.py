from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path

from kzz_monitor import __version__


parser = argparse.ArgumentParser(description="生成不含用户 Excel 的 Windows 在线更新发布目录")
parser.add_argument("--base-url", default="", help="更新目录的 HTTPS 基地址；留空生成相对 URL")
parser.add_argument("--notes", default="功能更新和问题修复")
args = parser.parse_args()

root = Path(__file__).resolve().parent
release = root / "release"
output = root / "updates"
output.mkdir(exist_ok=True)
package_name = f"KzzMonitor-update-v{__version__}-windows-x64.zip"
package = output / package_name
files = [
    "KzzMonitor.exe",
    "README.md",
    "KzzMonitor详细操作手册.md",
    "KzzMonitor详细操作手册.html",
    "KzzMonitor详细操作手册.pdf",
]
with zipfile.ZipFile(package, "w", zipfile.ZIP_DEFLATED) as bundle:
    for name in files:
        bundle.write(release / name, name)
sha256 = hashlib.sha256(package.read_bytes()).hexdigest()
base = args.base_url.rstrip("/")
url = f"{base}/{package_name}" if base else package_name
manifest = {
    "version": __version__,
    "notes": args.notes,
    "artifacts": {
        "windows-x64": {"url": url, "sha256": sha256},
    },
}
(output / "update-manifest.json").write_text(
    json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
)
print(package)
print(output / "update-manifest.json")
print(f"sha256={sha256}")
