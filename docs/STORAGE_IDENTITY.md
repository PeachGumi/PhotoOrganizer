# Storage identity

Photo Organizer must never treat a path string, drive letter, mount point, persistent UUID, or volume GUID as sufficient proof that the currently mounted storage is the same mount session that was scanned earlier.

The unified app therefore separates two concepts:

1. **OS volume fingerprint** — Windows volume GUID or macOS mounted-device fingerprint. This distinguishes different backing volumes when available.
2. **Process-local mount session ID** — a random identifier created when a mounted volume is first observed. It is discarded on removal or fingerprint change and is never persisted.

A safety approval can only remain valid while both the fingerprint and the current mount-session ID match. Removal followed by reinsertion creates a new session ID even for the same physical card.

## Event sources

Windows uses `Win32_VolumeChangeEvent` with a one-second enumeration fallback. macOS watches `/Volumes` with `FileSystemWatcher` and also uses the same periodic fallback. Removal events explicitly invalidate the old session before refreshing current mounts.

## Fail-closed rules

- If the platform fingerprint cannot be obtained, no storage identity snapshot is issued.
- Persistent identifiers are never accepted without the process-local session ID.
- Camera-card manual selection must resolve to a non-system mounted volume root that itself contains `DCIM` or `PRIVATE`.
- Selecting `DCIM/100NIKON` or another child expands to the complete mounted camera-card root.
- An ordinary folder is never accepted merely because its path resembles camera media.

Real-device acceptance must still exercise removal/reinsertion, same drive-letter or mount-path replacement, source removal during import, and destination removal during verification on both supported operating systems.
