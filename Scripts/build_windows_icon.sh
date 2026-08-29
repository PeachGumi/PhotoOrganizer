#!/usr/bin/env bash
set -euo pipefail

SOURCE_PNG="${1:-src/PhotoOrganizer.App/Assets/photo-organizer-icon.png}"
OUTPUT_ICO="${2:-src/PhotoOrganizer.App/Assets/photo-organizer-icon.ico}"

if [[ ! -s "$SOURCE_PNG" ]]; then
  echo "Source icon not found or empty: $SOURCE_PNG" >&2
  exit 1
fi

icon_workdir="$(mktemp -d)"
trap 'rm -rf "$icon_workdir"' EXIT

for pixels in 16 24 32 48 64 128 256; do
  sips -z "$pixels" "$pixels" "$SOURCE_PNG" --out "$icon_workdir/$pixels.png" >/dev/null
done

mkdir -p "$(dirname "$OUTPUT_ICO")"
python3 - "$icon_workdir" "$OUTPUT_ICO" <<'PY'
import os
import struct
import sys
from pathlib import Path

source = Path(sys.argv[1])
output = Path(sys.argv[2])
sizes = (16, 24, 32, 48, 64, 128, 256)
png_signature = b"\x89PNG\r\n\x1a\n"
payloads: list[tuple[int, bytes]] = []

for pixels in sizes:
    payload = (source / f"{pixels}.png").read_bytes()
    if not payload.startswith(png_signature) or len(payload) < 24:
        raise SystemExit(f"Generated {pixels}px icon is not a PNG")
    width, height = struct.unpack_from(">II", payload, 16)
    if width != pixels or height != pixels:
        raise SystemExit(
            f"Generated icon has wrong dimensions: expected {pixels}x{pixels}, got {width}x{height}"
        )
    payloads.append((pixels, payload))

header_size = 6 + (16 * len(payloads))
offset = header_size
directory = bytearray(struct.pack("<HHH", 0, 1, len(payloads)))
images = bytearray()

for pixels, payload in payloads:
    encoded_size = 0 if pixels == 256 else pixels
    directory.extend(
        struct.pack(
            "<BBBBHHII",
            encoded_size,
            encoded_size,
            0,
            0,
            1,
            32,
            len(payload),
            offset,
        )
    )
    images.extend(payload)
    offset += len(payload)

temporary = output.with_name(f".{output.name}.tmp-{os.getpid()}")
try:
    temporary.write_bytes(directory + images)
    os.replace(temporary, output)
finally:
    temporary.unlink(missing_ok=True)
PY

test -s "$OUTPUT_ICO"
echo "Created Windows icon: $OUTPUT_ICO"
