# Photo Organizer

Photo Organizer is a cross-platform desktop application for safely importing photos and videos from camera media on Windows and macOS.

This repository is the new canonical codebase that consolidates the behavior and lessons learned from:

- `PeachGumi/PhotoOrganizer-win`
- `PeachGumi/PhotoOrganizer-mac`

## Project direction

The long-term architecture is one shared .NET core and one cross-platform desktop UI, with only a small platform-specific layer for Windows and macOS integration.

The macOS implementation is the current reference for data-safety behavior. The Windows implementation is the original product implementation and provides useful Windows-specific behavior and existing C# components.

## Non-negotiable data-safety rules

Photo Organizer must never delete, move, rename, modify, or overwrite source media on the camera card.

A card must never be presented as safe to reuse merely because copying finished. Reuse approval requires a complete supported-media scan and verification that the selected destination contains byte-identical copies of every supported source file.

Supported media scope:

- JPG / JPEG
- configured RAW formats
- MOV / MP4

Unsupported sidecar or miscellaneous files are intentionally outside the import/reuse-verification scope.

## Status

This repository is under active consolidation. Do not treat it as a production release until the migration acceptance criteria are complete.

See `docs/ARCHITECTURE.md`, `docs/DATA_SAFETY.md`, and `docs/MIGRATION_PLAN.md` as they are added.
