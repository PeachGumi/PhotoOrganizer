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

Hidden or dot-prefixed directories are not an exception: supported JPG/JPEG, configured RAW, MOV, or MP4 media inside them remains in scope and must be discovered. Unsupported hidden sidecars/system metadata remain outside import scope. Reparse points/symlinks and nested mounted volumes are excluded so scanning cannot escape the selected physical card.

Any scan error is fail-closed. Supported zero-byte media is an error. An incomplete scan must never start an import or produce a reusable-card approval.

Manual selection of a nested camera directory must be expanded to the safe camera-card root. An arbitrary folder must not be accepted as equivalent to a complete card scan.

## Destination path and device independence

A destination used for copy, duplicate lookup, or final reuse verification must not pass through a user-controlled symbolic link, junction, or other reparse point. A platform may recognize a fixed OS-owned compatibility alias only when its resolved target is explicitly verified against the expected system path; all other aliases fail closed.

A lexical path alias must never count as an independent backup. In particular, a destination path that resolves back onto the camera card must not be allowed to write data or prove that the source bytes have been backed up.

Source and destination must resolve to different mounted-volume identities **and different physical storage-device identities**. Two partitions/volumes on one physical SD, USB drive, or disk are not independent backup locations. If the physical-device identity of either side cannot be established, import/reuse approval fails closed.

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
- source and destination volume identities, physical-device identities, and mount-session identities still match the identities captured for this operation;
- source and destination remain on different physical storage devices;
- the destination path remains free of user-controlled symlink/junction/reparse aliases;
- a fresh complete scan of the whole safe camera-card scope succeeds;
- all supported source media expected from the initial scan is still present;
- every currently supported source file is non-zero and readable;
- an independent destination file with equal size and SHA-256 exists for every supported source file;
- a source path itself, including the same bytes reached through a path alias, is never accepted as proof of an independent destination copy;
- storage identity is checked again after hashing before approval is displayed.

If any condition cannot be proved, the UI must remain blocked/not-verified.

## Storage replacement

Path strings alone are not storage identity. Unmount/remount, same-letter replacement on Windows, same-mount-path replacement on macOS, or a physical-device mapping change must invalidate a previous scan and approval.

Platform adapters are responsible for mounted-volume, physical-device, and process-local mount-session identities that fail closed when required identity cannot be established.

## Multi-card behavior

A newly detected second card must never replace the currently selected/processing card. It may be queued and scanned only after the current operation is safely completed or reset.

## Test policy

Every safety defect must receive a regression test in the shared Core whenever the behavior is platform-independent. Platform-specific storage and lifecycle defects must receive tests in their platform adapter project plus real-device release acceptance where simulation is insufficient.
