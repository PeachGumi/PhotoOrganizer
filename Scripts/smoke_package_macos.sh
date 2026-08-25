#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-0.0.0-ci}"
OUTPUT_ROOT="${2:-}"

if [[ -z "$OUTPUT_ROOT" ]]; then
  OUTPUT_ROOT="$(mktemp -d)"
  cleanup_root=1
else
  mkdir -p "$OUTPUT_ROOT"
  cleanup_root=0
fi

cleanup() {
  if [[ "$cleanup_root" -eq 1 ]]; then
    rm -rf "$OUTPUT_ROOT"
  fi
}
trap cleanup EXIT

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

  hdiutil create \
    -volname "Photo Organizer Smoke $arch" \
    -srcfolder "$dmg_stage" \
    -ov -format UDZO "$dmg" >/dev/null
  test -s "$dmg"
  hdiutil verify "$dmg" >/dev/null

  mount_point="$OUTPUT_ROOT/$rid/mount"
  mkdir -p "$mount_point"
  device="$(hdiutil attach "$dmg" -nobrowse -readonly -mountpoint "$mount_point" | awk 'NR==1 {print $1}')"
  test -n "$device"
  test -d "$mount_point/Photo Organizer.app"
  test -L "$mount_point/Applications"
  test -s "$mount_point/LICENSE.txt"
  test -s "$mount_point/Photo Organizer.app/Contents/Resources/PhotoOrganizer.icns"
  hdiutil detach "$device" >/dev/null

  hash="$(shasum -a 256 "$dmg" | awk '{print $1}')"
  [[ "$hash" =~ ^[0-9a-f]{64}$ ]]
  echo "Packaging smoke passed for $rid: $dmg ($hash)"
done
