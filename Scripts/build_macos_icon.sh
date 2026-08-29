#!/usr/bin/env bash
set -euo pipefail

SOURCE_PNG="${1:-src/PhotoOrganizer.App/Assets/photo-organizer-icon.png}"
OUTPUT_ICNS="${2:-PhotoOrganizer.icns}"

if [[ ! -s "$SOURCE_PNG" ]]; then
  echo "Source icon not found or empty: $SOURCE_PNG" >&2
  exit 1
fi

source_width="$(sips -g pixelWidth "$SOURCE_PNG" | awk '/pixelWidth/ { print $2 }')"
source_height="$(sips -g pixelHeight "$SOURCE_PNG" | awk '/pixelHeight/ { print $2 }')"
if [[ ! "$source_width" =~ ^[0-9]+$ || ! "$source_height" =~ ^[0-9]+$ \
      || "$source_width" -lt 1024 || "$source_height" -lt 1024 ]]; then
  echo "Source icon must be at least 1024x1024: ${source_width:-unknown}x${source_height:-unknown}" >&2
  exit 1
fi

icon_workdir="$(mktemp -d)"
trap 'rm -rf "$icon_workdir"' EXIT
source_png="$icon_workdir/source.png"
iconset="$icon_workdir/PhotoOrganizer.iconset"
mkdir -p "$iconset"

sips -z 1024 1024 "$SOURCE_PNG" --out "$source_png" >/dev/null

make_icon() {
  local pixels="$1"
  local name="$2"
  sips -z "$pixels" "$pixels" "$source_png" --out "$iconset/$name" >/dev/null
  test -s "$iconset/$name"
}

make_icon 16 icon_16x16.png
make_icon 32 icon_16x16@2x.png
make_icon 32 icon_32x32.png
make_icon 64 icon_32x32@2x.png
make_icon 128 icon_128x128.png
make_icon 256 icon_128x128@2x.png
make_icon 256 icon_256x256.png
make_icon 512 icon_256x256@2x.png
make_icon 512 icon_512x512.png
cp "$source_png" "$iconset/icon_512x512@2x.png"

mkdir -p "$(dirname "$OUTPUT_ICNS")"
iconutil -c icns "$iconset" -o "$OUTPUT_ICNS"
test -s "$OUTPUT_ICNS"
echo "Created macOS icon: $OUTPUT_ICNS"
