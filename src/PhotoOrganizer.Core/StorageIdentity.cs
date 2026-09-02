namespace PhotoOrganizer.Core;

public sealed record MountedVolumeInfo(
    string RootPath,
    string Fingerprint,
    bool IsRemovable,
    bool IsSystem,
    string? PhysicalDeviceFingerprint = null);

public sealed record StorageSessionIdentity(
    string RootPath,
    string Fingerprint,
    Guid SessionId,
    string? PhysicalDeviceFingerprint = null);

public interface IStorageVolumeProvider
{
    StringComparison PathComparison { get; }

    IReadOnlyList<MountedVolumeInfo> GetMountedVolumes();

    MountedVolumeInfo? ResolveVolumeForPath(string path);
}

/// <summary>
/// Tracks a process-local identity for each currently mounted filesystem volume.
/// A persistent OS volume identifier and physical-device identifier are fingerprints;
/// neither authorizes reuse by itself. Once a mount disappears or either identity
/// changes, its session id is discarded permanently.
/// </summary>
public sealed class StorageSessionTracker
{
    private sealed record SessionEntry(
        string Fingerprint,
        string? PhysicalDeviceFingerprint,
        Guid SessionId);

    private readonly IStorageVolumeProvider _provider;
    private readonly object _refreshGate = new();
    private readonly object _gate = new();
    private readonly Dictionary<string, SessionEntry> _sessions;

    public StorageSessionTracker(IStorageVolumeProvider provider)
    {
        _provider = provider;
        _sessions = new Dictionary<string, SessionEntry>(
            provider.PathComparison == StringComparison.OrdinalIgnoreCase
                ? StringComparer.OrdinalIgnoreCase
                : StringComparer.Ordinal);
    }

    public StringComparison PathComparison => _provider.PathComparison;

    public event Action<string>? VolumeMounted;
    public event Action<string>? VolumeRemoved;

    public void Refresh() => _ = RefreshAndGetSnapshot();

    private Dictionary<string, MountedVolumeInfo> RefreshAndGetSnapshot()
    {
        List<string> removed = [];
        List<string> added = [];
        Dictionary<string, MountedVolumeInfo> mounted;

        // Provider enumeration and session replacement are one serialized refresh.
        // Otherwise an older, slower enumeration can complete after a newer one and
        // temporarily resurrect the identity of a volume that has already changed.
        lock (_refreshGate)
        {
            var comparer = _provider.PathComparison == StringComparison.OrdinalIgnoreCase
                ? StringComparer.OrdinalIgnoreCase
                : StringComparer.Ordinal;
            mounted = _provider.GetMountedVolumes()
                .Where(v => !string.IsNullOrWhiteSpace(v.RootPath) && !string.IsNullOrWhiteSpace(v.Fingerprint))
                .Select(v => v with { RootPath = PathSafety.Normalize(v.RootPath) })
                .GroupBy(v => v.RootPath, comparer)
                .Where(group => group
                    .Select(volume => (volume.Fingerprint, volume.PhysicalDeviceFingerprint))
                    .Distinct()
                    .Count() == 1)
                .ToDictionary(group => group.Key, group => group.First(), comparer);

            lock (_gate)
            {
                foreach (var existing in _sessions.Keys.ToList())
                {
                    if (!mounted.TryGetValue(existing, out var current)
                        || !string.Equals(_sessions[existing].Fingerprint, current.Fingerprint, StringComparison.Ordinal)
                        || !string.Equals(
                            _sessions[existing].PhysicalDeviceFingerprint,
                            current.PhysicalDeviceFingerprint,
                            StringComparison.Ordinal))
                    {
                        _sessions.Remove(existing);
                        removed.Add(existing);
                    }
                }

                foreach (var volume in mounted.Values)
                {
                    if (_sessions.ContainsKey(volume.RootPath)) continue;
                    _sessions[volume.RootPath] = new SessionEntry(
                        volume.Fingerprint,
                        volume.PhysicalDeviceFingerprint,
                        Guid.NewGuid());
                    added.Add(volume.RootPath);
                }
            }
        }

        foreach (var root in removed) VolumeRemoved?.Invoke(root);
        foreach (var root in added) VolumeMounted?.Invoke(root);
        return mounted;
    }

    public void MarkRemoved(string rootPath)
    {
        var normalized = PathSafety.Normalize(rootPath);
        var removed = false;
        lock (_refreshGate)
        {
            lock (_gate)
            {
                removed = _sessions.Remove(normalized);
            }
        }

        if (removed) VolumeRemoved?.Invoke(normalized);
    }

    public StorageSessionIdentity? Capture(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;
        if (!PathSafety.TryValidateDirectFilesystemPath(path, out _)) return null;

        string normalizedPath;
        try
        {
            normalizedPath = PathSafety.Normalize(path);
        }
        catch
        {
            return null;
        }

        // Resolve the path from the exact same mounted-volume snapshot that refreshed
        // the process-local sessions. This avoids both a second expensive platform
        // enumeration and a split-brain result from two different mount snapshots.
        var mounted = RefreshAndGetSnapshot();
        var volume = mounted.Values
            .Where(v => PathSafety.IsSameOrDescendant(normalizedPath, v.RootPath, _provider.PathComparison))
            .OrderByDescending(v => v.RootPath.Length)
            .FirstOrDefault();
        if (volume is null || string.IsNullOrWhiteSpace(volume.Fingerprint)) return null;

        var root = volume.RootPath;
        lock (_gate)
        {
            if (!_sessions.TryGetValue(root, out var entry)) return null;
            if (!string.Equals(entry.Fingerprint, volume.Fingerprint, StringComparison.Ordinal)) return null;
            if (!string.Equals(
                    entry.PhysicalDeviceFingerprint,
                    volume.PhysicalDeviceFingerprint,
                    StringComparison.Ordinal)) return null;

            return new StorageSessionIdentity(
                root,
                entry.Fingerprint,
                entry.SessionId,
                entry.PhysicalDeviceFingerprint);
        }
    }

    public bool Matches(StorageSessionIdentity identity, string path)
    {
        var current = Capture(path);
        return current is not null
            && current.SessionId == identity.SessionId
            && string.Equals(current.Fingerprint, identity.Fingerprint, StringComparison.Ordinal)
            && string.Equals(
                current.PhysicalDeviceFingerprint,
                identity.PhysicalDeviceFingerprint,
                StringComparison.Ordinal)
            && string.Equals(current.RootPath, identity.RootPath, _provider.PathComparison);
    }

    public bool MatchesPair(
        StorageSessionIdentity? firstIdentity,
        string? firstPath,
        StorageSessionIdentity? secondIdentity,
        string? secondPath)
    {
        if (firstIdentity is null || secondIdentity is null
            || !TryNormalizeDirectPath(firstPath, out var normalizedFirstPath)
            || !TryNormalizeDirectPath(secondPath, out var normalizedSecondPath))
        {
            return false;
        }

        // Resolve both paths from the one mounted-volume snapshot used to refresh
        // the process-local sessions. Calling Capture twice here would enumerate
        // platform storage state twice and could observe different mount states.
        var mounted = RefreshAndGetSnapshot();
        var currentFirst = CaptureFromSnapshot(mounted, normalizedFirstPath);
        var currentSecond = CaptureFromSnapshot(mounted, normalizedSecondPath);

        return currentFirst is not null
            && currentSecond is not null
            && MatchesIdentity(firstIdentity, currentFirst)
            && MatchesIdentity(secondIdentity, currentSecond);
    }

    private StorageSessionIdentity? CaptureFromSnapshot(
        IReadOnlyDictionary<string, MountedVolumeInfo> mounted,
        string normalizedPath)
    {
        var volume = mounted.Values
            .Where(v => PathSafety.IsSameOrDescendant(normalizedPath, v.RootPath, _provider.PathComparison))
            .OrderByDescending(v => v.RootPath.Length)
            .FirstOrDefault();
        if (volume is null || string.IsNullOrWhiteSpace(volume.Fingerprint)) return null;

        var root = volume.RootPath;
        lock (_gate)
        {
            if (!_sessions.TryGetValue(root, out var entry)) return null;
            if (!string.Equals(entry.Fingerprint, volume.Fingerprint, StringComparison.Ordinal)) return null;
            if (!string.Equals(
                    entry.PhysicalDeviceFingerprint,
                    volume.PhysicalDeviceFingerprint,
                    StringComparison.Ordinal))
            {
                return null;
            }

            return new StorageSessionIdentity(
                root,
                entry.Fingerprint,
                entry.SessionId,
                entry.PhysicalDeviceFingerprint);
        }
    }

    private bool MatchesIdentity(StorageSessionIdentity expected, StorageSessionIdentity current) =>
        current.SessionId == expected.SessionId
        && string.Equals(current.Fingerprint, expected.Fingerprint, StringComparison.Ordinal)
        && string.Equals(
            current.PhysicalDeviceFingerprint,
            expected.PhysicalDeviceFingerprint,
            StringComparison.Ordinal)
        && string.Equals(current.RootPath, expected.RootPath, _provider.PathComparison);

    private static bool TryNormalizeDirectPath(string? path, out string normalized)
    {
        normalized = string.Empty;
        if (string.IsNullOrWhiteSpace(path)
            || !PathSafety.TryValidateDirectFilesystemPath(path, out _))
        {
            return false;
        }

        try
        {
            normalized = PathSafety.Normalize(path);
            return !string.IsNullOrWhiteSpace(normalized);
        }
        catch
        {
            return false;
        }
    }
}
