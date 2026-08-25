namespace PhotoOrganizer.Core;

public sealed record FormatVerificationResult(
    int Total,
    int Verified,
    IReadOnlyList<string> UnverifiedFiles,
    IReadOnlyList<string> Errors)
{
    public bool IsSafe => Total > 0 && Verified == Total && UnverifiedFiles.Count == 0 && Errors.Count == 0;
}

public sealed class FormatSafetyVerifier(MediaClassifier classifier)
{
    public async Task<FormatVerificationResult> VerifyAsync(
        IEnumerable<string> sourceFiles,
        string destinationRoot,
        CancellationToken cancellationToken = default)
    {
        var supported = sourceFiles
            .Where(classifier.IsSupported)
            .Select(Path.GetFullPath)
            .Distinct(PathComparer.Instance)
            .OrderBy(path => path, PathComparer.Instance)
            .ToArray();

        if (supported.Length == 0)
        {
            return new FormatVerificationResult(0, 0, [], ["No supported media exists to verify."]);
        }

        var destinationIndex = BuildDestinationIndex(destinationRoot);
        if (destinationIndex.Errors.Count > 0)
        {
            return new FormatVerificationResult(
                supported.Length,
                0,
                supported,
                destinationIndex.Errors);
        }

        var unverified = new List<string>();
        var errors = new List<string>();
        var verified = 0;

        foreach (var source in supported)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var sourceInfo = new FileInfo(source);
                if (!sourceInfo.Exists || sourceInfo.Length <= 0)
                {
                    unverified.Add(source);
                    continue;
                }

                if (!destinationIndex.FilesBySize.TryGetValue(sourceInfo.Length, out var candidates))
                {
                    unverified.Add(source);
                    continue;
                }

                var sourceHash = await Hashing.Sha256Async(source, cancellationToken).ConfigureAwait(false);
                var matched = false;

                foreach (var candidate in candidates)
                {
                    if (PathComparer.Instance.Equals(Path.GetFullPath(candidate), source))
                    {
                        // A source file can never prove that an independent destination copy exists.
                        continue;
                    }

                    try
                    {
                        var destinationHash = await Hashing.Sha256Async(candidate, cancellationToken).ConfigureAwait(false);
                        if (string.Equals(sourceHash, destinationHash, StringComparison.Ordinal))
                        {
                            matched = true;
                            break;
                        }
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        errors.Add($"{candidate}: {ex.Message}");
                    }
                }

                if (matched)
                {
                    verified++;
                }
                else
                {
                    unverified.Add(source);
                }
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                unverified.Add(source);
                errors.Add($"{source}: {ex.Message}");
            }
        }

        return new FormatVerificationResult(supported.Length, verified, unverified, errors);
    }

    private static DestinationIndex BuildDestinationIndex(string destinationRoot)
    {
        var index = new Dictionary<long, List<string>>();
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(destinationRoot) || !Directory.Exists(destinationRoot))
        {
            errors.Add("Destination root does not exist or is not a directory.");
            return new DestinationIndex(index, errors);
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

                    if (entry is not FileInfo file || file.Length <= 0 || file.Name.StartsWith(".partial-", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (!index.TryGetValue(file.Length, out var candidates))
                    {
                        candidates = [];
                        index[file.Length] = candidates;
                    }

                    candidates.Add(file.FullName);
                }
                catch (Exception ex)
                {
                    errors.Add($"{entry.FullName}: {ex.Message}");
                }
            }
        }

        return new DestinationIndex(index, errors);
    }

    private sealed record DestinationIndex(
        Dictionary<long, List<string>> FilesBySize,
        IReadOnlyList<string> Errors);
}
