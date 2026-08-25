#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-0.0.0-ci}"
OUTPUT_ROOT="${2:-}"
mounted_device=""

if [[ -z "$OUTPUT_ROOT" ]]; then
  OUTPUT_ROOT="$(mktemp -d)"
  cleanup_root=1
else
  mkdir -p "$OUTPUT_ROOT"
  cleanup_root=0
fi

detach_if_needed() {
  if [[ -n "$mounted_device" ]]; then
    hdiutil detach "$mounted_device" >/dev/null 2>&1 || hdiutil detach -force "$mounted_device" >/dev/null 2>&1 || true
    mounted_device=""
  fi
}

cleanup() {
  detach_if_needed
  if [[ "$cleanup_root" -eq 1 ]]; then
    rm -rf "$OUTPUT_ROOT"
  fi
}
trap cleanup EXIT

create_dmg() {
  local stage="$1"
  local output="$2"
  local volume_name="$3"
  local attempt

  rm -f "$output"
  for attempt in 1 2 3; do
    if hdiutil create -volname "$volume_name" -srcfolder "$stage" -ov -format UDZO "$output" >/dev/null; then
      return 0
    fi
    rm -f "$output"
    # diskimages-helper can take a moment to release the previous image device.
    sleep "$attempt"
  done

  echo "Unable to create DMG after 3 attempts: $output" >&2
  return 1
}

icon="$OUTPUT_ROOT/PhotoOrganizer.icns"
bash Scripts/build_macos_icon.sh src/PhotoOrganizer.App/Assets/app-icon.ico "$icon"
test -s "$icon"

for arch in arm64 x64; do
  rid="osx-$arch"
  publish="$OUTPUT_ROOT/$rid/publish"
  app="$OUTPUT_ROOT/$rid/Photo Organizer.app"
  contents="$app/Contents"
  macos="$contents/MacOS"
  resources="$contents/Resources"
  dmg_stage="$OUTPUT_ROOT/$rid/dmg"
  dmg="$OUTPUT_ROOT/PhotoOrganizer-macOS-$arch-$VERSION.dmg"

  dotnet publish src/PhotoOrganizer.App/PhotoOrganizer.App.csproj \
    --configuration Release \
    --runtime "$rid" \
    --self-contained true \
    -p:Version="$VERSION" \
    -p:ContinuousIntegrationBuild=true \
    --output "$publish"

  for required in PhotoOrganizer PhotoOrganizer.dll PhotoOrganizer.Core.dll PhotoOrganizer.deps.json PhotoOrganizer.runtimeconfig.json; do
    test -s "$publish/$required" || { echo "Missing publish output: $publish/$required" >&2; exit 1; }
  done

  case "$arch" in
    arm64) file "$publish/PhotoOrganizer" | grep -q 'arm64' ;;
    x64) file "$publish/PhotoOrganizer" | grep -Eq 'x86_64|x86-64' ;;
  esac

  rm -rf "$app"
  mkdir -p "$macos" "$resources"
  ditto "$publish" "$macos"
  chmod +x "$macos/PhotoOrganizer"
  cp LICENSE "$resources/LICENSE.txt"
  cp "$icon" "$resources/PhotoOrganizer.icns"

  cat > "$contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleDevelopmentRegion</key><string>en</string>
  <key>CFBundleDisplayName</key><string>Photo Organizer</string>
  <key>CFBundleExecutable</key><string>PhotoOrganizer</string>
  <key>CFBundleIconFile</key><string>PhotoOrganizer.icns</string>
  <key>CFBundleIdentifier</key><string>com.peachgumi.photoorganizer</string>
  <key>CFBundleInfoDictionaryVersion</key><string>6.0</string>
  <key>CFBundleName</key><string>Photo Organizer</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleShortVersionString</key><string>$VERSION</string>
  <key>CFBundleVersion</key><string>1</string>
  <key>LSMinimumSystemVersion</key><string>13.0</string>
  <key>NSHighResolutionCapable</key><true/>
</dict>
</plist>
PLIST

  plutil -lint "$contents/Info.plist"
  [[ "$(/usr/libexec/PlistBuddy -c 'Print :CFBundleExecutable' "$contents/Info.plist")" == "PhotoOrganizer" ]]
  [[ "$(/usr/libexec/PlistBuddy -c 'Print :CFBundleIconFile' "$contents/Info.plist")" == "PhotoOrganizer.icns" ]]
  test -x "$macos/PhotoOrganizer"
  test -s "$resources/PhotoOrganizer.icns"
  test -s "$resources/LICENSE.txt"

  rm -rf "$dmg_stage"
  mkdir -p "$dmg_stage"
  ditto "$app" "$dmg_stage/Photo Organizer.app"
  ln -s /Applications "$dmg_stage/Applications"
  cp LICENSE "$dmg_stage/LICENSE.txt"

  create_dmg "$dmg_stage" "$dmg" "Photo Organizer Smoke $arch"
  test -s "$dmg"
  hdiutil verify "$dmg" >/dev/null

  mount_point="$OUTPUT_ROOT/$rid/mount"
  mkdir -p "$mount_point"
  attach_plist="$OUTPUT_ROOT/$rid/attach.plist"
  hdiutil attach "$dmg" -nobrowse -readonly -mountpoint "$mount_point" -plist > "$attach_plist"
  mounted_device="$(python3 - "$attach_plist" "$mount_point" <<'PY'
import plistlib
import re
import sys
from pathlib import Path

payload = plistlib.loads(Path(sys.argv[1]).read_bytes())
target = str(Path(sys.argv[2]).resolve())
mounted = None
all_devices = []
for entity in payload.get("system-entities", []):
    device = entity.get("dev-entry")
    if device:
        all_devices.append(device)
    mount = entity.get("mount-point")
    if mount and device and str(Path(mount).resolve()) == target:
        mounted = device

if not mounted:
    raise SystemExit("Unable to resolve mounted DMG device from hdiutil plist")

match = re.match(r"^(/dev/disk\d+)", mounted)
whole = match.group(1) if match else mounted
if all_devices and whole not in all_devices:
    # Prefer an explicit whole-disk entity from the attach result when present.
    candidates = [d for d in all_devices if re.fullmatch(r"/dev/disk\d+", d)]
    if candidates:
        whole = candidates[0]
print(whole)
PY
)"
  test -n "$mounted_device"
  test -d "$mount_point/Photo Organizer.app"
  test -L "$mount_point/Applications"
  test -s "$mount_point/LICENSE.txt"
  test -s "$mount_point/Photo Organizer.app/Contents/Resources/PhotoOrganizer.icns"
  hdiutil detach "$mounted_device" >/dev/null
  mounted_device=""
  rmdir "$mount_point" 2>/dev/null || true

  hash="$(shasum -a 256 "$dmg" | awk '{print $1}')"
  [[ "$hash" =~ ^[0-9a-f]{64}$ ]]
  echo "Packaging smoke passed for $rid: $dmg ($hash)"
done
