namespace PhotoOrganizer.Core;

internal sealed class VolumeTraversalGuard
{
    private readonly string[] _nestedMountRoots;
    private readonly StringComparison _comparison;

    private VolumeTraversalGuard(string[] nestedMountRoots, StringComparison comparison)
    {
        _nestedMountRoots = nestedMountRoots;
        _comparison = comparison;
    }

    public static VolumeTraversalGuard Create(string traversalRoot, IStorageVolumeProvider? provider)
    {
        if (provider is null) return new VolumeTraversalGuard([], OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);

        var root = PathSafety.Normalize(traversalRoot);
        var nested = provider.GetMountedVolumes()
            .Select(v => PathSafety.Normalize(v.RootPath))
            .Where(mount => !string.Equals(mount, root, provider.PathComparison))
            .Where(mount => PathSafety.IsSameOrDescendant(mount, root, provider.PathComparison))
            .Distinct(provider.PathComparison == StringComparison.OrdinalIgnoreCase
                ? StringComparer.OrdinalIgnoreCase
                : StringComparer.Ordinal)
            .OrderByDescending(path => path.Length)
            .ToArray();

        return new VolumeTraversalGuard(nested, provider.PathComparison);
    }

    public bool IsNestedMountedVolume(string path)
    {
        var normalized = PathSafety.Normalize(path);
        return _nestedMountRoots.Any(root =>
            string.Equals(normalized, root, _comparison)
            || PathSafety.IsSameOrDescendant(normalized, root, _comparison));
    }
}
