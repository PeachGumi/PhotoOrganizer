# Release acceptance checklist

A workflow success is necessary but not sufficient for a production-ready Photo Organizer release. Complete this checklist against the exact signed artifacts from the candidate GitHub Release and retain evidence (release tag/SHA, OS version, machine architecture, card/filesystem, counts, relevant screenshots/logs, and checksums).

## Build provenance and signatures

- [ ] Candidate tag/version maps to the intended `main` commit.
- [ ] Windows x64 ZIP SHA-256 matches its published checksum.
- [ ] Windows arm64 ZIP SHA-256 matches its published checksum.
- [ ] macOS arm64 DMG SHA-256 matches its published checksum.
- [ ] macOS x64 DMG SHA-256 matches its published checksum.
- [ ] Windows `PhotoOrganizer.exe`, `PhotoOrganizer.dll`, and `PhotoOrganizer.Core.dll` pass Authenticode verification.
- [ ] macOS `.app` passes `codesign --verify --deep --strict` and `spctl -a -t exec`.
- [ ] macOS DMG passes `stapler validate` and Gatekeeper assessment.
- [ ] No unsigned substitute artifact is present in the Release.

## Clean-machine install and launch

- [ ] Windows x64 package launches on a clean supported x64 Windows machine without a .NET runtime preinstall.
- [ ] Windows arm64 package launches on a clean supported ARM64 Windows machine when hardware is available.
- [ ] Record SmartScreen behavior; investigate unexpected unsigned/unrecognized-publisher warnings. New-certificate reputation warnings must not be confused with signature failure.
- [ ] macOS arm64 DMG opens and app launches on Apple Silicon with Gatekeeper enabled.
- [ ] macOS x64 DMG opens and app launches on Intel macOS when hardware is available.
- [ ] Login-startup enable/disable survives a real logout/login cycle on each tested platform.
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
- [ ] Mount a different card/device at the same path after a successful scan: the old scan session is rejected.
- [ ] Unmount and remount the same card in the same app session: the previous mount-session approval is invalidated and a fresh scan is required.
- [ ] Insert a second camera card while the first is being scanned/imported: active target does not switch and the second card is queued.
- [ ] Remove a queued second card before it becomes active: it is removed from the queue and is not scanned from stale state.
- [ ] Add a new supported file to the card between initial scan and final verification: final rescan includes it, and reuse cannot become safe unless that file also exists byte-identically in the destination.
- [ ] Make an initially scanned supported file disappear before final verification: reuse is blocked.

## Retirement gate for legacy repositories

Do not archive or label `PhotoOrganizer-win` / `PhotoOrganizer-mac` as superseded for production until:

- [ ] all applicable acceptance checks above are recorded for a signed unified candidate;
- [ ] no open P0/P1 migration defect remains;
- [ ] the chosen customer download/distribution path is documented;
- [ ] rollback instructions point to the last known-good legacy release if the unified version has a release-blocking regression.
