#!/bin/bash
set -euo pipefail
LABEL="com.local.kzzmonitor"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
rm -f "$PLIST"
echo "Removed KzzMonitor LaunchAgent."
