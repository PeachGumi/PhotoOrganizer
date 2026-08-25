# Release acceptance checklist

A workflow success is necessary but not sufficient for a production-ready Photo Organizer release. Complete this checklist against the exact signed **GitHub Prerelease acceptance candidate** and retain evidence (release tag/SHA, OS version, machine architecture, card/filesystem, counts, relevant screenshots/logs, and checksums). Do not promote the candidate to Latest/stable until every applicable item is complete.

## Candidate provenance and signatures

- [ ] Candidate exists as a GitHub Prerelease, not as Latest/stable.
- [ ] Candidate tag/version maps to the intended exact `main` commit and that 40-character SHA is recorded.
- [ ] Windows x64 ZIP SHA-256 matches its published checksum.
- [ ] Windows arm64 ZIP SHA-256 matches its published checksum.
- [ ] macOS arm64 DMG SHA-256 matches its published checksum.
- [ ] macOS x64 DMG SHA-256 matches its published checksum.
- [ ] Windows `PhotoOrganizer.exe`, `PhotoOrganizer.dll`, and `PhotoOrganizer.Core.dll` pass Authenticode verification.
- [ ] macOS `.app` passes `codesign --verify --deep --strict` and `spctl -a -t exec`.
- [ ] macOS DMG passes `stapler validate` and Gatekeeper assessment.
- [ ] No unsigned substitute artifact is present in the Release.

## Clean-machine install, resident mode and launch

- [ ] Windows x64 package launches on a clean supported x64 Windows machine without a .NET runtime preinstall.
- [ ] Windows arm64 package launches on a clean supported ARM64 Windows machine when hardware is available.
- [ ] Windows Explorer, taskbar, and the main application window show the Photo Organizer product icon rather than a generic executable/window icon.
- [ ] Record SmartScreen behavior; investigate unexpected unsigned/unrecognized-publisher warnings. New-certificate reputation warnings must not be confused with signature failure.
- [ ] macOS arm64 DMG opens and app launches on Apple Silicon with Gatekeeper enabled.
- [ ] macOS x64 DMG opens and app launches on Intel macOS when hardware is available.
- [ ] macOS Finder and Dock show the Photo Organizer product icon from the signed app bundle rather than a generic application icon.
- [ ] Windows shows a usable tray icon/menu and macOS shows a usable menu-bar item/menu.
- [ ] Closing the main window while idle hides it but leaves camera-card monitoring active.
- [ ] Tray/menu-bar `Photo Organizerを表示` restores and activates the workflow window.
- [ ] Tray/menu-bar `終了` terminates the application while idle.
- [ ] During an active import, normal window close, tray/menu-bar Quit, and platform graceful quit requests do not interrupt the import.
- [ ] Enabling `バックグラウンド（トレイ/メニューバー）から開始` causes the next ordinary launch to start resident without flashing the main window.
- [ ] Launching with `--background` also starts resident without flashing the main window.
- [ ] An already-mounted or newly inserted valid camera card while resident automatically brings the workflow window forward.
- [ ] Disabling background start causes the next ordinary launch to show the main window.
- [ ] Login-startup enable/disable survives a real logout/login cycle on each tested platform.
- [ ] When login auto-start and background start are both enabled, the registered startup command launches with `--background`; disabling background start rewrites the registration without the flag.
- [ ] Application restart always returns SD reuse safety to an unverified state; a previous green state is never restored from settings.

## Real camera-card data safety

Use disposable/test media with independent ground-truth copies. Do not begin acceptance with irreplaceable photographs.

- [ ] Insert a camera-formatted card containing JPG/JPEG, configured RAW, and MOV/MP4; automatic detection scans the complete card.
- [ ] Manually select a nested folder such as `DCIM/100NIKON`; the app expands to and verifies the complete camera-card root.
- [ ] Arbitrary non-camera folders are rejected as camera-card roots.
- [ ] Unsupported XMP/XML/TXT sidecars are not copied and do not block supported-media reuse verification.
- [ ] A zero-byte supported JPG/RAW/video makes the scan incomplete/blocked.
- [ ] Destination on the camera card, or a parent/child path overlapping the card, is rejected before copying.
- [ ] Destination on another folder of the same physical volume is rejected.
- [ ] After import, every source file remains byte-identical and at the same source path; the app has not deleted, moved, renamed, modified, or overwritten any source media.
- [ ] Existing destination file with the same name but different bytes is preserved; imported data receives `_2`, `_3`, etc.
- [ ] Same-name identical bytes are treated as an already-backed-up duplicate and no unnecessary collision copy is created.
- [ ] Reusing a camera filename with the same size but different bytes is never mistaken for an old import.
- [ ] If a supported file copy fails, the run remains blocked rather than silently declaring success; rerunning safely skips files already proven byte-identical and retries the remaining work under a fresh transaction.
- [ ] Copy completion does not show the green reuse state before the post-import rescan and SHA-256 verification finish.
- [ ] Green `保存先コピー検証済み — SDカード再利用可能` appears only after every currently supported file on the card has an independent destination size+SHA-256 match.

## Repeated shooting workflow

Run at least three real cycles using the same card:

1. shoot/test files → insert → import → wait for verified green;
2. reuse/format card in camera as intended → shoot again → import;
3. repeat once more, including filename-number reuse if the camera permits it.

For every cycle:

- [ ] Historical files still present on the card and already byte-identical in the destination are skipped safely.
- [ ] Only genuinely new bytes consume new destination space.
- [ ] Duplicate-only import does not create an empty event folder.
- [ ] Event date is based on pending/new media when present.
- [ ] Final verification covers all supported media currently visible on the card.

## Failure and race scenarios

- [ ] Remove the source card during scan: scan/import is blocked and no reuse-safe state appears.
- [ ] Remove the source card during copy: operation fails closed; no reuse-safe state appears.
- [ ] Remove the destination during copy/final verification: operation fails closed; no reuse-safe state appears.
- [ ] Cancel during copy: source remains untouched and reuse remains blocked/unverified.
- [ ] Attempt normal application close during an active import: close is prevented until processing/cancellation resolves.
- [ ] Attempt tray/menu-bar Quit during an active import: shutdown is rejected and the workflow window is shown.
- [ ] Mount a different card/device at the same path after a successful scan: the old scan session is rejected.
- [ ] Unmount and remount the same card in the same app session: the previous mount-session approval is invalidated and a fresh scan is required.
- [ ] Insert a second camera card while the first is being scanned/imported: active target does not switch and the second card is queued.
- [ ] Remove a queued second card before it becomes active: it is removed from the queue and is not scanned from stale state.
- [ ] Add a new supported file to the card between initial scan and final verification: final rescan includes it, and reuse cannot become safe unless that file also exists byte-identically in the destination.
- [ ] Make an initially scanned supported file disappear before final verification: reuse is blocked.
- [ ] Force-kill the process during import, relaunch it, and confirm no previous reuse-safe state is restored.

## Stable promotion record

After every applicable check above passes:

- [ ] Store a durable evidence reference containing the tested tag, exact candidate commit, OS/hardware/card details and outcomes.
- [ ] Run `Promote accepted release` with the exact candidate tag and 40-character accepted commit SHA.
- [ ] Set `acceptance_confirmed=true` only after the checklist is complete.
- [ ] Supply the durable evidence reference to the promotion workflow.
- [ ] Confirm the workflow refuses a wrong commit, a non-prerelease release, or an incomplete artifact set.
- [ ] Confirm successful promotion clears the Prerelease flag, marks the same release Latest/stable, preserves the exact artifacts, and appends the accepted commit/evidence reference to its release notes.

## Retirement gate for legacy repositories

Do not archive or label `PhotoOrganizer-win` / `PhotoOrganizer-mac` as superseded for production until:

- [ ] all applicable acceptance checks above are recorded for a signed unified candidate;
- [ ] that exact candidate has been promoted through the explicit acceptance workflow;
- [ ] no open P0/P1 migration defect remains;
- [ ] the chosen customer download/distribution path points to the promoted stable release;
- [ ] rollback instructions point to the last known-good legacy release if the unified version has a release-blocking regression.
