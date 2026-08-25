namespace PhotoOrganizer.Core;

public sealed record BackupLookupResult(
    IReadOnlySet<string> MatchedSources,
    IReadOnlyList<string> Errors);

public sealed class DestinationLibrary
{
    public async Task<BackupLookupResult> FindVerifiedBackupsAsync(
        IEnumerable<string> sourceFiles,
        string destinationRoot,
        CancellationToken cancellationToken = default)
    {
        var index = BuildIndex(destinationRoot);
        var matched = new HashSet<string>(PathComparer.Instance);
        var errors = index.Errors.ToList();
        var destinationHashCache = new Dictionary<string, string>(PathComparer.Instance);

        foreach (var source in sourceFiles.Select(Path.GetFullPath).Distinct(PathComparer.Instance))
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var sourceInfo = new FileInfo(source);
                if (!sourceInfo.Exists || sourceInfo.Length <= 0) continue;
                if (!index.FilesBySize.TryGetValue(sourceInfo.Length, out var candidates)) continue;

                // Do not hash the source at all when no same-size destination candidate exists.
                var sourceHash = await Hashing.Sha256Async(source, cancellationToken).ConfigureAwait(false);

                foreach (var candidate in candidates)
                {
                    if (PathComparer.Instance.Equals(Path.GetFullPath(candidate), source)) continue;

                    if (!destinationHashCache.TryGetValue(candidate, out var candidateHash))
                    {
                        try
                        {
                            candidateHash = await Hashing.Sha256Async(candidate, cancellationToken).ConfigureAwait(false);
                            destinationHashCache[candidate] = candidateHash;
                        }
                        catch (Exception ex) when (ex is not OperationCanceledException)
                        {
                            errors.Add($"{candidate}: {ex.Message}");
                            continue;
                        }
                    }

                    if (string.Equals(sourceHash, candidateHash, StringComparison.Ordinal))
                    {
                        matched.Add(source);
                        break;
                    }
                }
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                errors.Add($"{source}: {ex.Message}");
            }
        }

        return new BackupLookupResult(matched, errors);
    }

    internal static DestinationIndex BuildIndex(string destinationRoot)
    {
        var filesBySize = new Dictionary<long, List<string>>();
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(destinationRoot))
        {
            errors.Add("Destination root is empty.");
            return new DestinationIndex(filesBySize, errors);
        }

        if (!Directory.Exists(destinationRoot))
        {
            return new DestinationIndex(filesBySize, errors);
        }

        var stack = new Stack<string>();
        stack.Push(Path.GetFullPath(destinationRoot));

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
                    if ((attributes & FileAttributes.ReparsePoint) != 0) continue;

                    if ((attributes & FileAttributes.Directory) != 0)
                    {
                        if ((attributes & FileAttributes.Hidden) != 0 || entry.Name.StartsWith('.')) continue;
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

        return new DestinationIndex(filesBySize, errors);
    }

    internal sealed record DestinationIndex(
        Dictionary<long, List<string>> FilesBySize,
        IReadOnlyList<string> Errors);
}
