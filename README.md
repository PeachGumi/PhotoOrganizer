# Photo Organizer

Photo Organizer is the unified Windows/macOS desktop application for safely importing photos and videos from camera media.

It consolidates the product behavior and lessons learned from the original repositories:

- `PeachGumi/PhotoOrganizer-win` — original Windows implementation and Windows integration reference
- `PeachGumi/PhotoOrganizer-mac` — hardened data-safety reference implementation

## Architecture

The unified product uses:

- .NET 10 LTS
- C# shared Core
- Avalonia desktop UI
- thin Windows/macOS platform adapters for storage detection, mounted-volume identity, startup integration, lifecycle handling, signing and packaging

The import, duplicate detection, copy integrity and SD-reuse safety state machine live in `PhotoOrganizer.Core` and are shared by both operating systems.

```text
src/
├─ PhotoOrganizer.Core/   shared safety and import logic
└─ PhotoOrganizer.App/    shared Avalonia desktop UI

tests/
└─ PhotoOrganizer.Core.Tests/  cross-platform regression tests
```

## Non-negotiable data safety

Photo Organizer must never delete, move, rename, modify, or overwrite source media on the camera card.

Copy completion is **not** equivalent to SD-card reuse approval. Reuse approval requires a fresh complete card scan, mount-session storage-identity checks, and byte-for-byte destination verification with file size and SHA-256.

Supported import/reuse-verification scope:

- JPG / JPEG
- configured RAW formats
- MOV / MP4

Unsupported sidecars and miscellaneous camera files are intentionally outside that scope.

The normative rules are in [`docs/DATA_SAFETY.md`](docs/DATA_SAFETY.md).

## Unified workflow

The Avalonia application provides one Windows/macOS workflow for:

- automatic and manual camera-card detection;
- complete-card scanning even when a nested `DCIM`/`PRIVATE` folder is selected;
- destination and event-name selection plus destination-layout preview;
- RAW extension settings;
- queued second-card handling without changing the active card;
- copy/cancel progress;
- explicit blocked/unverified/verified SD-reuse state;
- platform login-startup integration;
- Windows tray / macOS menu-bar resident operation using the same Avalonia implementation;
- optional background startup independent from login auto-start.

Closing the workflow window while idle hides it instead of stopping camera-card monitoring. Use the tray/menu-bar menu to show the window again or explicitly quit. Graceful quit is rejected during an active import. Starting with `--background`, or enabling the background-start preference, keeps the main window hidden until a camera card is detected or the user opens it from the tray/menu bar.

A green `保存先コピー検証済み — SDカード再利用可能` state is produced only by the shared Core after final post-import verification. It is never persisted across application restarts.

## Development

Required SDK: .NET 10.0.400 or a compatible later feature band selected by `global.json`.

```bash
dotnet restore PhotoOrganizer.slnx
dotnet test tests/PhotoOrganizer.Core.Tests/PhotoOrganizer.Core.Tests.csproj -c Release
dotnet build src/PhotoOrganizer.App/PhotoOrganizer.App.csproj -c Release
```

CI executes shared safety tests and builds the unified desktop app on both Windows and macOS. It also validates the fail-closed release contract. CodeQL and Dependabot are configured in-repository.

## Production release

The release workflow produces signed Windows x64/ARM64 packages and Developer ID signed/notarized macOS Apple Silicon/Intel DMGs from the same version and exact commit. No GitHub Release is created unless both platform signing pipelines succeed and every published artifact passes its SHA-256 verification.

See [`docs/RELEASE.md`](docs/RELEASE.md) for signing setup and [`docs/RELEASE_ACCEPTANCE.md`](docs/RELEASE_ACCEPTANCE.md) for the mandatory real-device acceptance checklist.

## Migration status

This repository is the canonical source for new Windows/macOS development. The shared safety Core, unified production workflow, release pipeline, and shared resident UI are migrated. The legacy applications remain available as references until the unified signed artifacts pass real-device acceptance.

Deliberate behavior differences from the legacy implementations, including removal of unsafe overwrite/timestamp-only duplicate behavior and silent file retry, are documented in [`docs/MIGRATION_PLAN.md`](docs/MIGRATION_PLAN.md).

See also [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md), [`docs/DATA_SAFETY.md`](docs/DATA_SAFETY.md), and [`SECURITY.md`](SECURITY.md).
