# Data Safety Contract

This document is normative. A platform implementation must not weaken these rules.

## Supported source media

Only the following files are in import and SD-reuse-verification scope:

- JPG / JPEG
- configured RAW formats
- MOV / MP4

Unsupported files such as XMP, XML, TXT, camera databases, and sidecars are not imported and do not block reuse approval.

## Source immutability

Photo Organizer must never delete, move, rename, modify, truncate, replace, or overwrite source media.

The application must not offer an operation that formats or erases the camera card.

## Complete scan requirement

A scan is complete only if every directory and relevant file in the selected camera-card scope was enumerated without an I/O or metadata error.

Any scan error is fail-closed. Supported zero-byte media is an error. An incomplete scan must never start an import or produce a reusable-card approval.

Manual selection of a nested camera directory must be expanded to the safe camera-card root. An arbitrary folder must not be accepted as equivalent to a complete card scan.

## Copy transaction

For every new source file:

1. Read source size and SHA-256.
2. Never delete or overwrite an existing destination file.
3. If an existing candidate has identical size and SHA-256, treat it as an already-imported duplicate.
4. If a name collides with different bytes, select `_2`, `_3`, and so on.
5. Copy to an app-created hidden `.partial-*` path.
6. Flush the temporary destination.
7. Verify temporary size and SHA-256.
8. Re-read the source SHA-256 and size to ensure the source did not change during copy.
9. Move the verified temporary file to an unused final path without overwrite.
10. Verify final size and SHA-256 again.

Only app-created never-finalized `.partial-*` files may be automatically deleted by copy cleanup. Existing library files and finalized copies must never be deleted as part of collision handling or error recovery.

## SD-card reuse approval

`copy completed` and `SD card may be reused` are different states.

Reuse approval requires all of the following after copy processing:

- the same selected source storage is still mounted;
- the same selected destination storage is still mounted;
- source and destination storage identities still match the identities captured for this operation;
- a fresh complete scan of the whole safe camera-card scope succeeds;
- all supported source media expected from the initial scan is still present;
- every currently supported source file is non-zero and readable;
- an independent destination file with equal size and SHA-256 exists for every supported source file;
- a source path itself is never accepted as proof of an independent destination copy;
- storage identity is checked again after hashing before approval is displayed.

If any condition cannot be proved, the UI must remain blocked/not-verified.

## Storage replacement

Path strings alone are not storage identity. Unmount/remount, same-letter replacement on Windows, or same-mount-path replacement on macOS must invalidate a previous scan and approval.

Platform adapters are responsible for a mount-session storage identity that fails closed when unavailable.

## Multi-card behavior

A newly detected second card must never replace the currently selected/processing card. It may be queued and scanned only after the current operation is safely completed or reset.

## Test policy

Every safety defect must receive a regression test in the shared Core whenever the behavior is platform-independent. Platform-specific storage and lifecycle defects must receive tests in their platform adapter project plus real-device release acceptance where simulation is insufficient.
