namespace PhotoOrganizer.Core;

public sealed record BackupLookupResult(
    IReadOnlySet<string> MatchedSources,
    IReadOnlyList<string> Errors);

public sealed class DestinationLibrary
{
    private readonly IStorageVolumeProvider? _volumeProvider;
    private readonly IFileHasher _hasher;
    private readonly int _maxDegreeOfParallelism;

    public DestinationLibrary(
        IStorageVolumeProvider? volumeProvider = null,
        IFileHasher? hasher = null,
        int maxDegreeOfParallelism = 1)
    {
        if (maxDegreeOfParallelism <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(maxDegreeOfParallelism),
                maxDegreeOfParallelism,
                "The maximum degree of parallelism must be positive.");
        }

        _volumeProvider = volumeProvider;
        _hasher = hasher ?? new Sha256FileHasher();
        _maxDegreeOfParallelism = maxDegreeOfParallelism;
    }

    public async Task<BackupLookupResult> FindVerifiedBackupsAsync(
        IEnumerable<string> sourceFiles,
        string destinationRoot,
        CancellationToken cancellationToken = default)
    {
        var destinationIndex = DestinationFileIndexer.Build(
            destinationRoot,
            _volumeProvider,
            requireExistingRoot: false,
            cancellationToken: cancellationToken);
        var errors = destinationIndex.Errors.ToList();
        var destinationHashCache = new AsyncHashCache(_hasher);
        var sources = sourceFiles
            .Select(Path.GetFullPath)
            .Distinct(PathComparer.Instance)
            .ToArray();
        var results = new LookupResult[sources.Length];

        if (_maxDegreeOfParallelism == 1)
        {
            for (var sourceIndex = 0; sourceIndex < sources.Length; sourceIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                results[sourceIndex] = await LookupSourceAsync(
                    sources[sourceIndex],
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
                Enumerable.Range(0, sources.Length),
                options,
                async (sourceIndex, token) =>
                {
                    results[sourceIndex] = await LookupSourceAsync(
                        sources[sourceIndex],
                        destinationIndex.FilesBySize,
                        destinationHashCache,
                        token).ConfigureAwait(false);
                }).ConfigureAwait(false);
        }

        var matched = new HashSet<string>(PathComparer.Instance);
        for (var sourceIndex = 0; sourceIndex < results.Length; sourceIndex++)
        {
            var result = results[sourceIndex];
            if (result.Matched) matched.Add(sources[sourceIndex]);
            errors.AddRange(result.Errors);
        }

        return new BackupLookupResult(matched, errors);
    }

    private async Task<LookupResult> LookupSourceAsync(
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
                return new LookupResult(false, errors);
            }

            // Size is only a cheap prefilter. If no destination file has the same
            // size there is no possible byte-identical backup, so avoid reading the
            // source solely to compute a SHA-256 that cannot match anything.
            if (!filesBySize.TryGetValue(sourceInfo.Length, out var candidates))
            {
                return new LookupResult(false, errors);
            }

            var sourceHash = await _hasher
                .Sha256Async(source, cancellationToken)
                .ConfigureAwait(false);

            foreach (var candidate in candidates)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (PathComparer.Instance.Equals(Path.GetFullPath(candidate), source)) continue;

                try
                {
                    // A destination file may be compared with several source files in
                    // one lookup. Cache its real SHA-256 only for this operation; the
                    // cache is never persisted and never reused by final reuse approval.
                    var candidateHash = await destinationHashCache
                        .GetAsync(candidate, cancellationToken)
                        .ConfigureAwait(false);

                    if (string.Equals(sourceHash, candidateHash, StringComparison.Ordinal))
                    {
                        return new LookupResult(true, errors);
                    }
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    errors.Add($"{candidate}: {ex.Message}");
                }
            }
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            errors.Add($"{source}: {ex.Message}");
        }

        return new LookupResult(false, errors);
    }

    private sealed record LookupResult(bool Matched, IReadOnlyList<string> Errors);
}
