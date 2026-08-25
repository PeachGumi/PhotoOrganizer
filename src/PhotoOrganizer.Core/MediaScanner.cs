namespace PhotoOrganizer.Core;

public sealed record MediaScanResult(IReadOnlyList<string> Files, IReadOnlyList<string> Errors)
{
    public bool IsComplete => Errors.Count == 0;
}

public sealed class MediaScanner
{
    private readonly MediaClassifier _classifier;
    private readonly IStorageVolumeProvider? _volumeProvider;

    public MediaScanner(MediaClassifier classifier, IStorageVolumeProvider? volumeProvider = null)
    {
        _classifier = classifier;
        _volumeProvider = volumeProvider;
    }

    public MediaScanResult Scan(string root)
    {
        var files = new List<string>();
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
        {
            return new MediaScanResult([], ["Scan root does not exist or is not a directory."]);
        }

        var scanRoot = Path.GetFullPath(root);
        VolumeTraversalGuard guard;
        try
        {
            guard = VolumeTraversalGuard.Create(scanRoot, _volumeProvider);
        }
        catch (Exception ex)
        {
            return new MediaScanResult([], [$"Unable to establish scan volume boundary: {ex.Message}"]);
        }

        var stack = new Stack<string>();
        stack.Push(scanRoot);

        while (stack.Count > 0)
        {
            var current = stack.Pop();
            FileSystemInfo[] entries;

            try
            {
                entries = new DirectoryInfo(current).EnumerateFileSystemInfos().ToArray();
            }
            catch (Exception ex)
            {
                errors.Add($"{current}: {ex.Message}");
                continue;
            }

            foreach (var entry in entries)
            {
                try
                {
                    var attributes = entry.Attributes;
                    if ((attributes & FileAttributes.ReparsePoint) != 0)
                    {
                        continue;
                    }

                    if ((attributes & FileAttributes.Directory) != 0)
                    {
                        // Hidden/dot-prefixed directories are still part of the
                        // camera-card scope. A supported JPG/RAW/video inside one
                        // would otherwise be silently omitted and could be lost when
                        // the user reformats the card after a false reuse approval.
                        // Only reparse points and nested mounted volumes are excluded.
                        if (guard.IsNestedMountedVolume(entry.FullName))
                        {
                            continue;
                        }

                        stack.Push(entry.FullName);
                        continue;
                    }

                    if (entry is not FileInfo file || !_classifier.IsSupported(file.FullName))
                    {
                        continue;
                    }

                    files.Add(file.FullName);
                    if (file.Length <= 0)
                    {
                        errors.Add($"{file.FullName}: supported media is zero bytes.");
                    }
                }
                catch (Exception ex)
                {
                    errors.Add($"{entry.FullName}: {ex.Message}");
                }
            }
        }

        files.Sort(PathComparer.Instance);
        return new MediaScanResult(files, errors);
    }
}

internal sealed class PathComparer : IComparer<string>, IEqualityComparer<string>
{
    public static PathComparer Instance { get; } = new();

    private static StringComparer Comparer => OperatingSystem.IsWindows()
        ? StringComparer.OrdinalIgnoreCase
        : StringComparer.Ordinal;

    public int Compare(string? x, string? y) => Comparer.Compare(x, y);
    public bool Equals(string? x, string? y) => Comparer.Equals(x, y);
    public int GetHashCode(string obj) => Comparer.GetHashCode(obj);
}
