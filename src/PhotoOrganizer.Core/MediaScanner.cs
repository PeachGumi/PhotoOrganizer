namespace PhotoOrganizer.Core;

public sealed record MediaScanResult(IReadOnlyList<string> Files, IReadOnlyList<string> Errors)
{
    public bool IsComplete => Errors.Count == 0;
}

public sealed class MediaScanner(MediaClassifier classifier)
{
    public MediaScanResult Scan(string root)
    {
        var files = new List<string>();
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
        {
            return new MediaScanResult([], ["Scan root does not exist or is not a directory."]);
        }

        var stack = new Stack<string>();
        stack.Push(Path.GetFullPath(root));

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
                        if ((attributes & FileAttributes.Hidden) != 0 || entry.Name.StartsWith('.'))
                        {
                            continue;
                        }

                        stack.Push(entry.FullName);
                        continue;
                    }

                    if (entry is not FileInfo file || !classifier.IsSupported(file.FullName))
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
