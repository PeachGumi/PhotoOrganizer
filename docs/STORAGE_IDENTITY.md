# Storage identity

Photo Organizer must never treat a path string, drive letter, mount point, persistent UUID, volume GUID, or physical-disk number as sufficient proof that the currently mounted storage is the same mount session that was scanned earlier.

The unified app therefore separates three concepts:

1. **OS volume fingerprint** — a Windows volume GUID or, on macOS, a persistent filesystem/partition UUID reported by `diskutil info -plist` when available. macOS falls back to the current BSD `DeviceIdentifier` only when no persistent UUID is exposed.
2. **Physical storage fingerprint** — the Windows physical-disk mapping or macOS `ParentWholeDisk`/whole-disk identifier. Source and destination must resolve to different physical storage devices.
3. **Process-local mount session ID** — a random identifier created when a mounted volume is first observed. It is discarded on removal or fingerprint/physical-device change and is never persisted.

A safety approval can only remain valid while the volume fingerprint, physical storage fingerprint, and current mount-session ID match. Removal followed by reinsertion creates a new session ID even for the same physical card when the removal is observed.

## Event sources

Windows uses `Win32_VolumeChangeEvent` with a one-second enumeration fallback. macOS watches `/Volumes` with `FileSystemWatcher` and also uses the same periodic fallback. Removal events explicitly invalidate the old session before refreshing current mounts.

On macOS each enumeration uses one bounded `diskutil info -plist` operation per mounted volume to obtain both the volume and whole-disk identities. Redirected output is drained asynchronously and a timed-out child process is killed; storage identity collection must not block indefinitely on a failing filesystem. The persistent UUID is a fingerprint only: it never substitutes for the process-local mount-session ID.

## Fail-closed rules

- If the platform volume fingerprint or required physical-device fingerprint cannot be obtained, no reusable-card approval can be issued.
- Persistent identifiers are never accepted without the process-local session ID.
- Different partitions/volumes on the same physical device are not independent backup locations.
- Camera-card manual selection must resolve to a non-system mounted volume root that itself contains `DCIM` or `PRIVATE`.
- Selecting `DCIM/100NIKON` or another child expands to the complete mounted camera-card root.
- An ordinary folder is never accepted merely because its path resembles camera media.

Real-device acceptance must still exercise removal/reinsertion, same drive-letter or mount-path replacement, source removal during import, destination removal during verification, and physically distinct source/destination detection on both supported operating systems.
