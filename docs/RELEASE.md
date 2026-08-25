# Production release

Photo Organizer publishes one version for Windows and macOS from one exact `main` commit. Release is intentionally fail-closed: if the protected signing environment, any required signing credential, or either platform build fails, no candidate is published. A signed CI candidate is still **not** a production-ready stable release until the real-device acceptance checklist passes and a separate promotion workflow records that approval.

## Canonical artifacts

For version `1.0.0`, the release workflow produces:

- `PhotoOrganizer-Windows-x64-1.0.0.zip`
- `PhotoOrganizer-Windows-arm64-1.0.0.zip`
- `PhotoOrganizer-macOS-arm64-1.0.0.dmg`
- `PhotoOrganizer-macOS-x64-1.0.0.dmg`
- one `.sha256` file for each artifact

Windows packages are self-contained .NET applications. The Photo Organizer executable and application assemblies are Authenticode-signed with SHA-256 and an RFC 3161 SHA-256 timestamp before the ZIP is created.

macOS packages contain a self-contained Avalonia `.app`. Mach-O components and the app are signed with Developer ID Application and Hardened Runtime. The app and DMG are notarized with `notarytool`, stapled, and verified with `codesign`, `stapler`, and Gatekeeper (`spctl`). Apple Silicon and Intel are released separately because a self-contained .NET application contains architecture-specific runtime files; merely using `lipo` on the app host would not make the entire runtime universal.

## Protected production-signing environment

Production signing credentials are **environment secrets**, not repository-wide Actions secrets. The repository hardening helper creates a GitHub Environment named `production-signing` and resets its deployment policy to exactly one allowed branch: `main`.

Run this first from an authenticated administrator checkout:

```bash
bash Scripts/configure_repository.sh
```

The helper also configures and verifies `main` protection and a read-only default `GITHUB_TOKEN`. If any required hardening check does not match after the API calls, the helper exits non-zero instead of printing a false success message.

GitHub only makes environment secrets available to jobs that reference that environment and satisfy its protection rules. Consequently, a release workflow copied or modified on a release branch/tag cannot obtain the production signing secrets from `production-signing`.

## Required `production-signing` environment secrets

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

Do not commit any certificate, private key, password, or app-specific password. On a trusted local machine, run:

```bash
bash Scripts/configure_release_secrets.sh
```

The script refuses to accept credentials unless `production-signing` already has the verified main-only deployment policy. It sends the values through stdin to `gh secret set --env production-signing`, verifies that every environment-secret name exists, and only then removes legacy repository-level copies of the same signing-secret names. Secret values are not deliberately printed.

If an older setup used repository-level signing secrets, rerunning this helper performs that migration after the protected environment has been configured.

## Stage 1 — signed acceptance candidate

1. Ensure the exact `main` commit has green `required` CI and CodeQL and no unresolved code blocker remains.
2. Run `Scripts/configure_repository.sh` and verify that `main` protection plus the `production-signing` main-only environment policy succeed.
3. Configure all signing credentials with `Scripts/configure_release_secrets.sh`.
4. Manually dispatch `.github/workflows/release.yml` **from `main`** with the desired `vMAJOR.MINOR.PATCH` input. For example:

   ```bash
   gh workflow run release.yml --repo PeachGumi/PhotoOrganizer --ref main -f version=v1.0.0
   ```

5. The unprivileged preflight rejects any ref other than `refs/heads/main` and validates the semantic version.
6. A `production-signing` environment job requires every Windows and Apple credential before signed platform work begins. Windows and macOS signing jobs also reference the protected environment.
7. Windows and macOS independently run the shared safety tests from the same exact main commit, produce signed packages, verify signatures, and upload short-lived workflow artifacts.
8. Only after both platform jobs succeed does the protected publish job download the complete set, verify every SHA-256 checksum, reject an already-existing release **or tag**, and create a new **GitHub Prerelease acceptance candidate** whose tag targets that exact main commit.
9. The candidate is intentionally not marked Latest or stable. Use these exact candidate bytes for clean-machine and real-camera-card acceptance.

The release workflow has no push/tag/release-branch trigger. The candidate tag is created by the final publish job only after the complete signed artifact set has passed verification. The release workflow must never promote its own output to stable.

## Stage 2 — explicit stable promotion

After every applicable item in `docs/RELEASE_ACCEPTANCE.md` has passed for the exact prerelease candidate, run `.github/workflows/promote-release.yml` manually.

The promotion workflow requires:

- `version` — the exact `vMAJOR.MINOR.PATCH` candidate tag;
- `candidate_commit` — the exact 40-character commit SHA recorded during acceptance;
- `acceptance_confirmed=true` — an explicit human confirmation that the checklist passed;
- `evidence` — a durable reference to the recorded acceptance evidence (for example a GitHub issue/comment or controlled test record).

Before promotion it verifies that:

- the release exists and is a published prerelease, not a draft;
- its target commit exactly equals `candidate_commit`;
- all four expected Windows/macOS artifacts are present;
- all four corresponding SHA-256 files are present.

Only then does it clear the prerelease flag, mark the release as Latest, and append the accepted commit and evidence reference to the release notes.

If any of these checks fail, the candidate remains a prerelease and customer-facing stable distribution must not point to it as the approved production release.

## What does not count as production-ready

None of the following is sufficient on its own:

- a successful normal CI run;
- a successful signing/notarization workflow;
- an unsigned local build;
- only one platform succeeding;
- a prerelease candidate that has not completed the real-device checklist;
- a release whose accepted commit or evidence record cannot be identified;
- a release run from any ref other than `main`;
- credentials left only as repository-level secrets rather than in the protected `production-signing` environment.

## SmartScreen and Gatekeeper

Authenticode signature validity can be verified in CI. Windows SmartScreen reputation is reputation-based and must also be checked on a clean Windows environment as part of release acceptance; valid signing does not guarantee an established SmartScreen reputation for a new certificate/application.

On macOS, CI verifies Developer ID signatures, notarization tickets, stapling, and Gatekeeper assessment. A clean-machine launch remains part of real-device acceptance.
