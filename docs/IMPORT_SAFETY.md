# Import and camera-card reuse safety

This document defines the cross-platform import state machine. It applies equally to Windows and macOS and belongs in `PhotoOrganizer.Core`, not in either platform UI.

## Safety invariant

**Copy completion is not camera-card reuse approval.**

The only transition to `SafeToReuse` is the complete sequence below. Every error, cancellation, storage replacement, incomplete scan, missing file, or verification failure ends in `Blocked`.

## Sequence

1. Resolve a manual or automatic selection to the complete mounted camera-card root containing `DCIM` or `PRIVATE`.
2. Capture the card's process-local mount-session identity.
3. Scan the whole card for supported media. Supported media is JPG/JPEG, configured RAW, MOV, and MP4.
4. Require an error-free scan and re-check the same source mount session after scanning.
5. Capture the destination mount-session identity and reject the source volume, parent/child overlap, or unavailable identity.
6. Search the existing destination library for byte-identical prior imports using file size and SHA-256. This is only a skip optimization; it is not the final reuse proof.
7. Compute the event date from files that still need copying, falling back to all supported files for a duplicate-only run. The earliest detected capture date is used.
8. Preflight free space for pending bytes when capacity is available.
9. Before every copy, re-check both source and destination mount-session identities from one fresh mounted-volume snapshot.
10. Copy with small bounded concurrency, but keep each file inside its own hidden `.partial-*` transaction: hash the exact source bytes read while copying, flush once, freshly verify temporary size and SHA-256, re-hash the source, finalize without overwrite, durably synchronize, and freshly verify the final file again.
11. Report copy processing complete. The card is **still not approved for reuse**.
12. Re-check both storage identities.
13. Re-scan the complete card.
14. Require the rescan to be complete and non-empty, and require every supported path observed before import to still be present.
15. Freshly verify every supported file currently visible on the card against the destination library by size and SHA-256.
16. Re-check both storage identities after hashing and durability synchronization.
17. Re-scan the complete card once more and require the supported-file set to be identical to the set that was just verified. This catches supported media added or removed while final SHA-256/durability verification was running.
18. Re-check both storage identities after that final consistency scan.
19. Only then transition to `SafeToReuse`.

Duplicate lookup and final verification may process a small bounded number of source files concurrently. Every final source-to-destination proof still performs its own durability synchronization and ordered fresh post-durability destination/source hashes before it can count as verified.

A new supported file that appears on the card after copying is intentionally included by the post-copy rescan. If no byte-identical destination copy exists, reuse approval is blocked. A supported file that appears or disappears while final verification is running is caught by the final consistency scan and also blocks reuse approval.

## Filesystem boundaries

Source scans, destination-library searches, and final verification never descend into another mounted volume nested under the selected root. Reparse points/symlinks and hidden directories are skipped. This prevents a camera card mounted below a destination tree from accidentally proving its own backup.

## Source immutability

The application never deletes, moves, renames, overwrites, formats, or otherwise modifies source camera media. Cleanup is limited to application-created destination-side temporary files that were never finalized.

## Duplicate-only runs

If every supported source file already has an independent byte-identical copy in the destination library, no new event directory is created. Fresh post-import and post-verification scans plus fresh final verification are still required before reuse approval.

## Cancellation and application exit

Cancellation, exceptions, or process termination before the final consistency scan and verification sequence completes cannot produce or preserve a `SafeToReuse` state. The UI must present these cases as blocked/not verified.
