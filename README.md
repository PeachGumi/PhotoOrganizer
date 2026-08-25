# Photo Organizer

Photo Organizer is the unified Windows/macOS desktop application for safely importing photos and videos from camera media.

It consolidates the product behavior and lessons learned from the original repositories:

- `PeachGumi/PhotoOrganizer-win` — original Windows implementation and Windows integration reference
- `PeachGumi/PhotoOrganizer-mac` — current data-safety reference implementation

## Architecture

The unified product uses:

- .NET 10 LTS
- C# shared Core
- Avalonia desktop UI
- thin Windows/macOS platform adapters for storage detection, mounted-volume identity, startup integration, lifecycle handling, signing and packaging

Most future fixes should land once in `PhotoOrganizer.Core` and apply to both operating systems.

```text
src/
├─ PhotoOrganizer.Core/   shared safety and import logic
└─ PhotoOrganizer.App/    shared Avalonia desktop UI

tests/
└─ PhotoOrganizer.Core.Tests/  cross-platform regression tests
```

## Non-negotiable data safety

Photo Organizer must never delete, move, rename, modify, or overwrite source media on the camera card.

Copy completion is **not** equivalent to SD-card reuse approval. Reuse approval requires a fresh complete card scan, storage-identity checks, and byte-for-byte destination verification.

Supported import/reuse-verification scope:

- JPG / JPEG
- configured RAW formats
- MOV / MP4

Unsupported sidecars and miscellaneous camera files are intentionally outside that scope.

The normative rules are in [`docs/DATA_SAFETY.md`](docs/DATA_SAFETY.md).

## Development

Required SDK: .NET 10.0.400 or a compatible later feature band selected by `global.json`.

```bash
dotnet restore PhotoOrganizer.slnx
dotnet test tests/PhotoOrganizer.Core.Tests/PhotoOrganizer.Core.Tests.csproj -c Release
dotnet build src/PhotoOrganizer.App/PhotoOrganizer.App.csproj -c Release
```

CI executes shared safety tests and builds the unified desktop app on both Windows and macOS. CodeQL and Dependabot are configured in-repository.

## Migration status

This repository is now the canonical target for new cross-platform development, but it is **not yet a production replacement** for the two legacy applications.

The migration stages and retirement criteria are documented in [`docs/MIGRATION_PLAN.md`](docs/MIGRATION_PLAN.md). Existing repositories remain available as references until real-device and signed-release acceptance is complete.

See also [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) and [`SECURITY.md`](SECURITY.md).
