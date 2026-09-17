# Storage identity

Photo Organizer must never treat a path string, drive letter, mount point, persistent UUID, volume GUID, or physical-disk number as sufficient proof that the currently mounted storage is the same mount session that was scanned earlier.

The unified app therefore separates three concepts:

1. **OS volume fingerprint** — a Windows volume GUID or, on macOS, a persistent filesystem/partition UUID reported by Disk Arbitration when available. macOS falls back to the current BSD device name only when no persistent UUID is exposed.
2. **Physical storage fingerprint** — the Windows physical-disk mapping with a confirmed physical `MSFT_Disk.BusType`, or macOS physical whole-disk identifier established through IOKit ancestry. Source and destination must resolve to different physical storage devices. VHD/virtual disks, Storage Spaces, unknown buses, and RAID/iSCSI mappings without a single physical-device proof fail closed on Windows. An APFS container's synthesized whole-disk identifier is not a physical-device identity.
3. **Process-local mount session ID** — a random identifier created when a mounted volume is first observed. It is discarded on removal or fingerprint/physical-device change and is never persisted.

A safety approval can only remain valid while the volume fingerprint, physical storage fingerprint, and current mount-session ID match. Removal followed by reinsertion creates a new session ID even for the same physical card when the removal is observed.

## Event sources

Windows uses `Win32_VolumeChangeEvent` with a one-second enumeration fallback. macOS watches `/Volumes` with `FileSystemWatcher` and also uses the same periodic fallback. Removal events explicitly invalidate the old session before refreshing current mounts.

On macOS each enumeration queries Disk Arbitration and IOKit directly, without subprocesses or a successful-identity cache. IOKit ancestry is followed through APFS containers to the physical block device. Ambiguous parent chains, composited APFS storage, and file-backed or RAM virtual devices cannot establish an independent physical fingerprint. The persistent UUID is a fingerprint only: it never substitutes for the process-local mount-session ID.

Native identity tests compare the result against fresh `diskutil` information, following `APFSPhysicalStores` to the actual backing disk rather than using the synthesized container's `ParentWholeDisk`. A disposable disk-image test checks that a volume identity alone does not establish physical independence.

The platform definitions are documented in Microsoft's [MSFT_Disk reference](https://learn.microsoft.com/en-us/windows-hardware/drivers/storage/msft-disk) and Apple's [IOStorage protocol characteristics](https://github.com/apple-oss-distributions/IOStorageFamily/blob/main/IOStorageProtocolCharacteristics.h). Apple's `Virtual Interface` interconnect explicitly includes file-backed and RAM storage.

## Fail-closed rules

- If the platform volume fingerprint or required physical-device fingerprint cannot be obtained, no reusable-card approval can be issued.
- Persistent identifiers are never accepted without the process-local session ID.
- Different partitions/volumes on the same physical device are not independent backup locations.
- Camera-card manual selection must resolve to a non-system mounted volume root that itself contains `DCIM` or `PRIVATE`.
- Selecting `DCIM/100NIKON` or another child expands to the complete mounted camera-card root.
- An ordinary folder is never accepted merely because its path resembles camera media.

Real-device acceptance must still exercise removal/reinsertion, same drive-letter or mount-path replacement, source removal during import, destination removal during verification, and physically distinct source/destination detection on both supported operating systems.
