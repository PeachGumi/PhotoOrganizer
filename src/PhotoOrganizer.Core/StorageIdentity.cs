namespace PhotoOrganizer.Core;

public sealed record MountedVolumeInfo(
    string RootPath,
    string Fingerprint,
    bool IsRemovable,
    bool IsSystem);

public sealed record StorageSessionIdentity(
    string RootPath,
    string Fingerprint,
    Guid SessionId);

public interface IStorageVolumeProvider
{
    StringComparison PathComparison { get; }

    IReadOnlyList<MountedVolumeInfo> GetMountedVolumes();

    MountedVolumeInfo? ResolveVolumeForPath(string path);
}

public static class PathSafety
{
    public static string Normalize(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;

        var full = Path.GetFullPath(path);
        var root = Path.GetPathRoot(full);
        if (!string.IsNullOrEmpty(root) && string.Equals(full, root, StringComparison.OrdinalIgnoreCase))
        {
            return root;
        }

        return full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    public static bool IsSameOrDescendant(string path, string ancestor, StringComparison comparison)
    {
        var normalizedPath = Normalize(path);
        var normalizedAncestor = Normalize(ancestor);
        if (string.Equals(normalizedPath, normalizedAncestor, comparison)) return true;

        var prefix = normalizedAncestor.EndsWith(Path.DirectorySeparatorChar)
            || normalizedAncestor.EndsWith(Path.AltDirectorySeparatorChar)
            ? normalizedAncestor
            : normalizedAncestor + Path.DirectorySeparatorChar;

        return normalizedPath.StartsWith(prefix, comparison);
    }
}

/// <summary>
/// Tracks a process-local identity for each currently mounted filesystem volume.
/// A persistent OS volume identifier is only a fingerprint; it never authorizes reuse by itself.
/// Once a mount disappears or its fingerprint changes, its session id is discarded permanently.
/// </summary>
public sealed class StorageSessionTracker
{
    private sealed record SessionEntry(string Fingerprint, Guid SessionId);

    private readonly IStorageVolumeProvider _provider;
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

    public void Refresh()
    {
        var mounted = _provider.GetMountedVolumes()
            .Where(v => !string.IsNullOrWhiteSpace(v.RootPath) && !string.IsNullOrWhiteSpace(v.Fingerprint))
            .Select(v => v with { RootPath = PathSafety.Normalize(v.RootPath) })
            .ToDictionary(
                v => v.RootPath,
                v => v,
                _provider.PathComparison == StringComparison.OrdinalIgnoreCase
                    ? StringComparer.OrdinalIgnoreCase
                    : StringComparer.Ordinal);

        List<string> removed = [];
        List<string> added = [];

        lock (_gate)
        {
            foreach (var existing in _sessions.Keys.ToList())
            {
                if (!mounted.TryGetValue(existing, out var current)
                    || !string.Equals(_sessions[existing].Fingerprint, current.Fingerprint, StringComparison.Ordinal))
                {
                    _sessions.Remove(existing);
                    removed.Add(existing);
                }
            }

            foreach (var volume in mounted.Values)
            {
                if (_sessions.ContainsKey(volume.RootPath)) continue;
                _sessions[volume.RootPath] = new SessionEntry(volume.Fingerprint, Guid.NewGuid());
                added.Add(volume.RootPath);
            }
        }

        foreach (var root in removed) VolumeRemoved?.Invoke(root);
        foreach (var root in added) VolumeMounted?.Invoke(root);
    }

    public void MarkRemoved(string rootPath)
    {
        var normalized = PathSafety.Normalize(rootPath);
        var removed = false;
        lock (_gate)
        {
            removed = _sessions.Remove(normalized);
        }

        if (removed) VolumeRemoved?.Invoke(normalized);
    }

    public StorageSessionIdentity? Capture(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;

        string physicalPath;
        try
        {
            physicalPath = PhysicalPathResolver.Resolve(path);
        }
        catch
        {
            // An unresolved redirect/ancestor cannot prove which volume will
            // actually receive or supply bytes, so identity must fail closed.
            return null;
        }

        Refresh();

        var volume = _provider.ResolveVolumeForPath(physicalPath);
        if (volume is null || string.IsNullOrWhiteSpace(volume.Fingerprint)) return null;

        var root = PathSafety.Normalize(volume.RootPath);
        lock (_gate)
        {
            if (!_sessions.TryGetValue(root, out var entry)) return null;
            if (!string.Equals(entry.Fingerprint, volume.Fingerprint, StringComparison.Ordinal)) return null;
            return new StorageSessionIdentity(root, entry.Fingerprint, entry.SessionId);
        }
    }

    public bool Matches(StorageSessionIdentity identity, string path)
    {
        var current = Capture(path);
        return current is not null
            && current.SessionId == identity.SessionId
            && string.Equals(current.Fingerprint, identity.Fingerprint, StringComparison.Ordinal)
            && string.Equals(current.RootPath, identity.RootPath, _provider.PathComparison);
    }
}
