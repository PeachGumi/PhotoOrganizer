namespace PhotoOrganizer.Core;

public sealed class CameraCardRootResolver
{
    private readonly IStorageVolumeProvider _provider;

    public CameraCardRootResolver(IStorageVolumeProvider provider)
    {
        _provider = provider;
    }

    public string? Resolve(string selectedPath)
    {
        if (string.IsNullOrWhiteSpace(selectedPath)) return null;

        string selected;
        try
        {
            // A manually chosen alias/symlink must be evaluated by where it
            // physically lands, not by the user-facing path string.
            selected = PhysicalPathResolver.Resolve(selectedPath);
        }
        catch
        {
            return null;
        }

        if (!Directory.Exists(selected)) return null;

        var volume = _provider.GetMountedVolumes()
            .Where(v => !v.IsSystem)
            .Where(v => PathSafety.IsSameOrDescendant(selected, v.RootPath, _provider.PathComparison))
            .OrderByDescending(v => PathSafety.Normalize(v.RootPath).Length)
            .FirstOrDefault();

        if (volume is null) return null;

        var root = PathSafety.Normalize(volume.RootPath);
        if (!Directory.Exists(root)) return null;

        var hasDcim = Directory.Exists(Path.Combine(root, "DCIM"));
        var hasPrivate = Directory.Exists(Path.Combine(root, "PRIVATE"));
        if (!hasDcim && !hasPrivate) return null;

        return root;
    }

    public IReadOnlyList<string> GetCandidateRoots()
    {
        return _provider.GetMountedVolumes()
            .Where(v => !v.IsSystem)
            .Select(v => PathSafety.Normalize(v.RootPath))
            .Where(root => Directory.Exists(Path.Combine(root, "DCIM"))
                || Directory.Exists(Path.Combine(root, "PRIVATE")))
            .Distinct(_provider.PathComparison == StringComparison.OrdinalIgnoreCase
                ? StringComparer.OrdinalIgnoreCase
                : StringComparer.Ordinal)
            .OrderBy(root => root, _provider.PathComparison == StringComparison.OrdinalIgnoreCase
                ? StringComparer.OrdinalIgnoreCase
                : StringComparer.Ordinal)
            .ToList();
    }
}
