#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"

PYTHON_BIN="${PYTHON_BIN:-python3}"
ARCH="$(uname -m)"
if [[ "$ARCH" != "arm64" && "$ARCH" != "x86_64" ]]; then
  echo "Unsupported macOS architecture: $ARCH" >&2
  exit 1
fi

rm -rf .venv-macos build-macos dist-macos release-macos
"$PYTHON_BIN" -m venv .venv-macos
.venv-macos/bin/python -m pip install --upgrade pip setuptools wheel
.venv-macos/bin/python -m pip install '.[build,test]'
.venv-macos/bin/python -m pytest -q
.venv-macos/bin/python generate_template.py

.venv-macos/bin/pyinstaller \
  --noconfirm --clean --windowed --onedir \
  --name KzzMonitor \
  --distpath dist-macos \
  --workpath build-macos \
  --collect-all akshare \
  --collect-all py_mini_racer \
  --collect-all pystray \
  --hidden-import AppKit \
  --add-data "可转债监控.xlsx:." \
  --osx-bundle-identifier "com.local.kzzmonitor" \
  launcher.py

mkdir -p release-macos
cp -R dist-macos/KzzMonitor.app release-macos/
cp 可转债监控.xlsx README-macOS.md KzzMonitor详细操作手册.html KzzMonitor详细操作手册.pdf install_macos_startup.sh uninstall_macos_startup.sh release-macos/
chmod +x release-macos/*.sh
codesign --force --deep --sign - release-macos/KzzMonitor.app

APP_BINARY="release-macos/KzzMonitor.app/Contents/MacOS/KzzMonitor"
"$APP_BINARY" --version-probe
"$APP_BINARY" --data-probe
ditto -c -k --sequesterRsrc --keepParent release-macos "KzzMonitor-macOS-$ARCH.zip"

VERSION="$(.venv-macos/bin/python -c 'from kzz_monitor import __version__; print(__version__)')"
PLATFORM_KEY="macos-x64"
[[ "$ARCH" == "arm64" ]] && PLATFORM_KEY="macos-arm64"
UPDATE_PACKAGE="KzzMonitor-update-v$VERSION-$PLATFORM_KEY.zip"
rm -rf update-macos-payload
mkdir update-macos-payload
cp -R release-macos/KzzMonitor.app update-macos-payload/
cp release-macos/KzzMonitor详细操作手册.html release-macos/KzzMonitor详细操作手册.pdf update-macos-payload/
ditto -c -k --sequesterRsrc --keepParent update-macos-payload "$UPDATE_PACKAGE"
rm -rf update-macos-payload
UPDATE_SHA="$(shasum -a 256 "$UPDATE_PACKAGE" | awk '{print $1}')"
cat > "update-manifest-$PLATFORM_KEY.json" <<EOF
{
  "version": "$VERSION",
  "notes": "KzzMonitor macOS $VERSION",
  "artifacts": {
    "$PLATFORM_KEY": {
      "url": "$UPDATE_PACKAGE",
      "sha256": "$UPDATE_SHA"
    }
  }
}
EOF

echo "Build complete: $(pwd)/release-macos/KzzMonitor.app ($ARCH)"
echo "Portable archive: $(pwd)/KzzMonitor-macOS-$ARCH.zip"
echo "Online update package: $(pwd)/$UPDATE_PACKAGE"
echo "First launch: Control-click KzzMonitor.app, then choose Open."
