namespace PhotoOrganizer.Core;

public sealed record FormatVerificationResult(
    int Total,
    int Verified,
    IReadOnlyList<string> UnverifiedFiles,
    IReadOnlyList<string> Errors)
{
    public bool IsSafe => Total > 0 && Verified == Total && UnverifiedFiles.Count == 0 && Errors.Count == 0;
}

public sealed class FormatSafetyVerifier
{
    private readonly MediaClassifier _classifier;
    private readonly IStorageVolumeProvider? _volumeProvider;
    private readonly IFileHasher _hasher;
    private readonly IFileDurabilityService _durability;

    public FormatSafetyVerifier(
        MediaClassifier classifier,
        IStorageVolumeProvider? volumeProvider = null,
        IFileHasher? hasher = null,
        IFileDurabilityService? durability = null)
    {
        _classifier = classifier;
        _volumeProvider = volumeProvider;
        _hasher = hasher ?? new Sha256FileHasher();
        _durability = durability ?? new PlatformFileDurabilityService();
    }

    public async Task<FormatVerificationResult> VerifyAsync(
        IEnumerable<string> sourceFiles,
        string destinationRoot,
        CancellationToken cancellationToken = default)
    {
        var supported = sourceFiles
            .Where(_classifier.IsSupported)
            .Select(Path.GetFullPath)
            .Distinct(PathComparer.Instance)
            .OrderBy(path => path, PathComparer.Instance)
            .ToArray();

        if (supported.Length == 0)
        {
            return new FormatVerificationResult(0, 0, [], ["No supported media exists to verify."]);
        }

        var destinationIndex = DestinationFileIndexer.Build(
            destinationRoot,
            _volumeProvider,
            requireExistingRoot: true,
            cancellationToken: cancellationToken);
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

        // This cache is only a cheap within-invocation prefilter. A cached hash may
        // never approve reuse by itself: every source-to-destination match receives a
        // fresh durability sync and post-sync hash below.
        var destinationHashCache = new Dictionary<string, string>(PathComparer.Instance);

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

                if (!PathSafety.TryValidateDirectFilesystemPath(source, out var sourcePathError))
                {
                    unverified.Add(source);
                    errors.Add($"{source}: source path is not direct: {sourcePathError}");
                    continue;
                }

                // No same-size candidate means no byte-identical destination can
                // exist, so do not read/hash the source unnecessarily.
                if (!destinationIndex.FilesBySize.TryGetValue(sourceInfo.Length, out var candidates))
                {
                    unverified.Add(source);
                    continue;
                }

                var sourceHash = await _hasher.Sha256Async(source, cancellationToken).ConfigureAwait(false);
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
                        if (!PathSafety.TryValidateDirectFilesystemPath(candidate, out var candidatePathError))
                        {
                            errors.Add($"{candidate}: destination path is not direct: {candidatePathError}");
                            continue;
                        }

                        if (!destinationHashCache.TryGetValue(candidate, out var destinationHash))
                        {
                            destinationHash = await _hasher.Sha256Async(candidate, cancellationToken).ConfigureAwait(false);
                            destinationHashCache[candidate] = destinationHash;
                        }

                        if (!string.Equals(sourceHash, destinationHash, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        // A cache-readable SHA match is not sufficient for SD reuse.
                        // Synchronize the matched independent destination first.
                        var durability = _durability.EnsureDurable(candidate);
                        if (!durability.Success)
                        {
                            errors.Add($"{candidate}: {durability.Error ?? "durable destination commit failed"}");
                            continue;
                        }

                        if (!PathSafety.TryValidateDirectFilesystemPath(candidate, out candidatePathError))
                        {
                            errors.Add($"{candidate}: destination path changed during durable verification: {candidatePathError}");
                            continue;
                        }

                        // Hash again *after* durable synchronization. Another process
                        // may have replaced or modified the candidate between the first
                        // hash handle closing and the durability handle opening. The
                        // earlier hash is therefore only a prefilter, never final proof.
                        var postDurabilityHash = await _hasher
                            .Sha256Async(candidate, cancellationToken)
                            .ConfigureAwait(false);
                        destinationHashCache[candidate] = postDurabilityHash;

                        if (!string.Equals(sourceHash, postDurabilityHash, StringComparison.Ordinal))
                        {
                            errors.Add($"{candidate}: destination bytes changed during durable verification.");
                            continue;
                        }

                        // Re-read the source after the destination durability proof.
                        // Without this step, media changed in place during the durable
                        // sync could be approved based on an obsolete source hash.
                        if (!PathSafety.TryValidateDirectFilesystemPath(source, out sourcePathError))
                        {
                            errors.Add($"{source}: source path changed during durable verification: {sourcePathError}");
                            continue;
                        }

                        var freshSourceHash = await _hasher
                            .Sha256Async(source, cancellationToken)
                            .ConfigureAwait(false);
                        if (!string.Equals(freshSourceHash, postDurabilityHash, StringComparison.Ordinal))
                        {
                            errors.Add($"{source}: source bytes changed during durable verification.");
                            continue;
                        }

                        matched = true;
                        break;
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
}
