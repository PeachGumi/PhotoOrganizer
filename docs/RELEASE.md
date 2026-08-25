# Production release

Photo Organizer publishes one version for Windows and macOS from one exact commit. A production release is intentionally fail-closed: if any required signing credential or platform build fails, no GitHub Release is created.

## Canonical artifacts

For version `1.0.0`, the release workflow produces:

- `PhotoOrganizer-Windows-x64-1.0.0.zip`
- `PhotoOrganizer-Windows-arm64-1.0.0.zip`
- `PhotoOrganizer-macOS-arm64-1.0.0.dmg`
- `PhotoOrganizer-macOS-x64-1.0.0.dmg`
- one `.sha256` file for each artifact

Windows packages are self-contained .NET applications. The Photo Organizer executable and application assemblies are Authenticode-signed with SHA-256 and an RFC 3161 SHA-256 timestamp before the ZIP is created.

macOS packages contain a self-contained Avalonia `.app`. Mach-O components and the app are signed with Developer ID Application and Hardened Runtime. The app and DMG are notarized with `notarytool`, stapled, and verified with `codesign`, `stapler`, and Gatekeeper (`spctl`). Apple Silicon and Intel are released separately because a self-contained .NET application contains architecture-specific runtime files; merely using `lipo` on the app host would not make the entire runtime universal.

## Required GitHub Actions secrets

Windows:

- `WINDOWS_CERTIFICATE` — base64 of the production PFX
- `WINDOWS_CERTIFICATE_PASSWORD`

macOS:

- `MACOS_CERTIFICATE` — base64 of the Developer ID `.p12`
- `MACOS_CERTIFICATE_PASSWORD`
- `DEVELOPER_ID_APPLICATION` — exact Developer ID Application signing identity
- `APPLE_ID`
- `APPLE_TEAM_ID`
- `APPLE_APP_SPECIFIC_PASSWORD`

Do not commit any certificate, private key, password, or app-specific password. On a trusted local machine, `bash Scripts/configure_release_secrets.sh` sends values to `gh secret set` through stdin without deliberately printing the values.

## Release flow

1. Ensure `main` CI and CodeQL are green and the release acceptance checklist has no unresolved code blocker.
2. Configure all signing secrets.
3. Create or fast-forward `release/vX.Y.Z` to the exact approved `main` commit, or push tag `vX.Y.Z`.
4. The `Signing preflight` job validates the semantic version and requires every Windows and Apple signing secret before platform release jobs run.
5. Windows and macOS independently run the shared safety tests from the same commit, produce signed packages, verify signatures, and upload short-lived workflow artifacts.
6. Only after both platform jobs succeed does `Publish complete cross-platform release` download the complete set, verify every SHA-256 checksum, and create the GitHub Release.
7. Run and record the real-device acceptance checklist in `docs/RELEASE_ACCEPTANCE.md` before describing the build as production-ready.

A successful copy/build job, an unsigned local build, or a partially successful platform release is never a production release.

## SmartScreen and Gatekeeper

Authenticode signature validity can be verified in CI. Windows SmartScreen reputation is reputation-based and must also be checked on a clean Windows environment as part of release acceptance; valid signing does not guarantee an established SmartScreen reputation for a new certificate/application.

On macOS, CI verifies Developer ID signatures, notarization tickets, stapling, and Gatekeeper assessment. A clean-machine launch remains part of real-device acceptance.
