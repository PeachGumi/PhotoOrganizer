namespace PhotoOrganizer.Core;

internal sealed record DestinationFileIndex(
    Dictionary<long, List<string>> FilesBySize,
    IReadOnlyList<string> Errors);

internal static class DestinationFileIndexer
{
    public static DestinationFileIndex Build(
        string destinationRoot,
        IStorageVolumeProvider? volumeProvider,
        bool requireExistingRoot,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var filesBySize = new Dictionary<long, List<string>>();
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(destinationRoot))
        {
            errors.Add(requireExistingRoot
                ? "Destination root does not exist or is not a directory."
                : "Destination root is empty.");
            return new DestinationFileIndex(filesBySize, errors);
        }

        if (!PathSafety.TryValidateDirectFilesystemPath(destinationRoot, out var pathError))
        {
            errors.Add($"Destination root is not a direct filesystem path: {pathError}");
            return new DestinationFileIndex(filesBySize, errors);
        }

        if (!Directory.Exists(destinationRoot))
        {
            if (requireExistingRoot)
            {
                errors.Add("Destination root does not exist or is not a directory.");
            }
            return new DestinationFileIndex(filesBySize, errors);
        }

        var root = Path.GetFullPath(destinationRoot);
        VolumeTraversalGuard guard;
        try
        {
            guard = VolumeTraversalGuard.Create(root, volumeProvider);
        }
        catch (Exception ex)
        {
            errors.Add($"Unable to establish destination volume boundary: {ex.Message}");
            return new DestinationFileIndex(filesBySize, errors);
        }

        var stack = new Stack<string>();
        stack.Push(root);

        while (stack.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
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
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    var attributes = entry.Attributes;
                    if ((attributes & FileAttributes.ReparsePoint) != 0) continue;

                    if ((attributes & FileAttributes.Directory) != 0)
                    {
                        if ((attributes & FileAttributes.Hidden) != 0 || entry.Name.StartsWith('.')) continue;
                        if (guard.IsNestedMountedVolume(entry.FullName)) continue;
                        stack.Push(entry.FullName);
                        continue;
                    }

                    if (entry is not FileInfo file
                        || file.Length <= 0
                        || file.Name.StartsWith(".partial-", StringComparison.Ordinal)
                        || file.Name.Contains(".partial-", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (!filesBySize.TryGetValue(file.Length, out var candidates))
                    {
                        candidates = [];
                        filesBySize[file.Length] = candidates;
                    }
                    candidates.Add(file.FullName);
                }
                catch (Exception ex)
                {
                    errors.Add($"{entry.FullName}: {ex.Message}");
                }
            }
        }

        return new DestinationFileIndex(filesBySize, errors);
    }
}
