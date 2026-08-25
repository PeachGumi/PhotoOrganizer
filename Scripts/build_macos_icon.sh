#!/usr/bin/env bash
set -euo pipefail

SOURCE_ICO="${1:-src/PhotoOrganizer.App/Assets/app-icon.ico}"
OUTPUT_ICNS="${2:-PhotoOrganizer.icns}"

if [[ ! -s "$SOURCE_ICO" ]]; then
  echo "Source icon not found or empty: $SOURCE_ICO" >&2
  exit 1
fi

workdir="$(mktemp -d)"
trap 'rm -rf "$workdir"' EXIT
source_png="$workdir/source.png"
iconset="$workdir/PhotoOrganizer.iconset"
mkdir -p "$iconset"

# The repository icon is a multi-image ICO whose modern entries are PNG payloads.
# Extract the largest embedded PNG using only Python's standard library so the
# release process does not acquire an extra package-manager dependency.
python3 - "$SOURCE_ICO" "$source_png" <<'PY'
import struct
import sys
from pathlib import Path

source = Path(sys.argv[1])
out = Path(sys.argv[2])
data = source.read_bytes()
if len(data) < 6:
    raise SystemExit("ICO header is truncated")
reserved, kind, count = struct.unpack_from("<HHH", data, 0)
if reserved != 0 or kind != 1 or count < 1:
    raise SystemExit("Invalid ICO header")

png_signature = b"\x89PNG\r\n\x1a\n"
entries = []
for index in range(count):
    offset = 6 + (index * 16)
    if offset + 16 > len(data):
        raise SystemExit("ICO directory is truncated")
    width_raw, height_raw = data[offset], data[offset + 1]
    width = width_raw or 256
    height = height_raw or 256
    size, image_offset = struct.unpack_from("<II", data, offset + 8)
    end = image_offset + size
    if image_offset < 0 or size <= 0 or end > len(data):
        raise SystemExit("ICO image entry is out of bounds")
    payload = data[image_offset:end]
    if payload.startswith(png_signature):
        entries.append((width * height, width, height, payload))

if not entries:
    raise SystemExit("ICO contains no PNG-backed image entry")
_, width, height, payload = max(entries, key=lambda item: item[0])
out.write_bytes(payload)
print(f"Extracted {width}x{height} PNG from {source}")
PY

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
make_icon 1024 icon_512x512@2x.png

mkdir -p "$(dirname "$OUTPUT_ICNS")"
iconutil -c icns "$iconset" -o "$OUTPUT_ICNS"
test -s "$OUTPUT_ICNS"
echo "Created macOS icon: $OUTPUT_ICNS"
