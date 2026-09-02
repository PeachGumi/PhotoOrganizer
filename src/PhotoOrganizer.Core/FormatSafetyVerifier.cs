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
    private readonly int _maxDegreeOfParallelism;

    public FormatSafetyVerifier(
        MediaClassifier classifier,
        IStorageVolumeProvider? volumeProvider = null,
        IFileHasher? hasher = null,
        IFileDurabilityService? durability = null,
        int maxDegreeOfParallelism = 1)
    {
        if (maxDegreeOfParallelism <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(maxDegreeOfParallelism),
                maxDegreeOfParallelism,
                "The maximum degree of parallelism must be positive.");
        }

        _classifier = classifier;
        _volumeProvider = volumeProvider;
        _hasher = hasher ?? new Sha256FileHasher();
        _durability = durability ?? new PlatformFileDurabilityService();
        _maxDegreeOfParallelism = maxDegreeOfParallelism;
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

        var errors = new List<string>();

        // This cache is only a cheap within-invocation prefilter. A cached hash may
        // never approve reuse by itself: every source-to-destination match receives a
        // fresh durability sync and post-sync hash below.
        var destinationHashCache = new AsyncHashCache(_hasher);
        var results = new SourceVerificationResult[supported.Length];

        if (_maxDegreeOfParallelism == 1)
        {
            for (var sourceIndex = 0; sourceIndex < supported.Length; sourceIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                results[sourceIndex] = await VerifySourceAsync(
                    supported[sourceIndex],
                    destinationIndex.FilesBySize,
                    destinationHashCache,
                    cancellationToken).ConfigureAwait(false);
            }
        }
        else
        {
            var options = new ParallelOptions
            {
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = _maxDegreeOfParallelism
            };

            await Parallel.ForEachAsync(
                Enumerable.Range(0, supported.Length),
                options,
                async (sourceIndex, token) =>
                {
                    results[sourceIndex] = await VerifySourceAsync(
                        supported[sourceIndex],
                        destinationIndex.FilesBySize,
                        destinationHashCache,
                        token).ConfigureAwait(false);
                }).ConfigureAwait(false);
        }

        var unverified = new List<string>();
        var verified = 0;
        for (var sourceIndex = 0; sourceIndex < results.Length; sourceIndex++)
        {
            var result = results[sourceIndex];
            if (result.Matched)
            {
                verified++;
            }
            else
            {
                unverified.Add(supported[sourceIndex]);
            }

            errors.AddRange(result.Errors);
        }

        return new FormatVerificationResult(supported.Length, verified, unverified, errors);
    }

    private async Task<SourceVerificationResult> VerifySourceAsync(
        string source,
        Dictionary<long, List<string>> filesBySize,
        AsyncHashCache destinationHashCache,
        CancellationToken cancellationToken)
    {
        var errors = new List<string>();

        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            var sourceInfo = new FileInfo(source);
            if (!sourceInfo.Exists || sourceInfo.Length <= 0)
            {
                return new SourceVerificationResult(false, errors);
            }

            if (!PathSafety.TryValidateDirectFilesystemPath(source, out var sourcePathError))
            {
                errors.Add($"{source}: source path is not direct: {sourcePathError}");
                return new SourceVerificationResult(false, errors);
            }

            // No same-size candidate means no byte-identical destination can
            // exist, so do not read/hash the source unnecessarily.
            if (!filesBySize.TryGetValue(sourceInfo.Length, out var candidates))
            {
                return new SourceVerificationResult(false, errors);
            }

            // Keep this source's proof order strict: the source prehash must complete
            // before any candidate is durable-checked.
            var sourceHash = await _hasher
                .Sha256Async(source, cancellationToken)
                .ConfigureAwait(false);
            var matched = false;

            foreach (var candidate in candidates)
            {
                cancellationToken.ThrowIfCancellationRequested();
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

                    // This cache is only a prefilter. It is shared safely by source
                    // workers, but its lifetime is limited to this verification call.
                    var destinationHash = await destinationHashCache
                        .GetAsync(candidate, cancellationToken)
                        .ConfigureAwait(false);

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

            return new SourceVerificationResult(matched, errors);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            errors.Add($"{source}: {ex.Message}");
            return new SourceVerificationResult(false, errors);
        }
    }

    private sealed record SourceVerificationResult(bool Matched, IReadOnlyList<string> Errors);
}
