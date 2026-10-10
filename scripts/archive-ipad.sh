#!/bin/bash
set -euo pipefail
TASK_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
TASK_OUTPUT="${1:-$TASK_ROOT/safari/build/ipad}"
TASK_BUILD="${BUILD_NUMBER:-1}"
TASK_PROJECT="$TASK_ROOT/safari/ios/NYCU E3 Helper iOS/NYCU E3 Helper iOS.xcodeproj"
# Build outside Desktop/iCloud folders: Finder metadata can invalidate codesigning.
TASK_STAGING="$(mktemp -d "${TMPDIR:-/tmp}/e3-ipad-archive.XXXXXX")"
trap 'rm -rf "$TASK_STAGING"' EXIT
python3 "$TASK_ROOT/scripts/sync-safari.py"
xcodebuild -project "$TASK_PROJECT" -scheme 'NYCU E3 Helper iOS' \
  -configuration Release -destination 'generic/platform=iOS' \
  -archivePath "$TASK_STAGING/NYCU E3 Helper.xcarchive" \
  -derivedDataPath "$TASK_STAGING/DerivedData" \
  CURRENT_PROJECT_VERSION="$TASK_BUILD" -allowProvisioningUpdates archive
xcodebuild -exportArchive -archivePath "$TASK_STAGING/NYCU E3 Helper.xcarchive" \
  -exportPath "$TASK_STAGING/export" \
  -exportOptionsPlist "$TASK_ROOT/safari/ios/ExportOptions.plist" \
  -allowProvisioningUpdates
mkdir -p "$TASK_OUTPUT"
ditto --norsrc "$TASK_STAGING/NYCU E3 Helper.xcarchive" "$TASK_OUTPUT/NYCU E3 Helper.xcarchive"
ditto --norsrc "$TASK_STAGING/export" "$TASK_OUTPUT/export"
