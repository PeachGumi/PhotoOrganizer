namespace PhotoOrganizer.Core;

public sealed record BackupLookupResult(
    IReadOnlySet<string> MatchedSources,
    IReadOnlyList<string> Errors);

public sealed class DestinationLibrary
{
    private readonly IStorageVolumeProvider? _volumeProvider;
    private readonly IFileHasher _hasher;

    public DestinationLibrary(
        IStorageVolumeProvider? volumeProvider = null,
        IFileHasher? hasher = null)
    {
        _volumeProvider = volumeProvider;
        _hasher = hasher ?? new Sha256FileHasher();
    }

    public async Task<BackupLookupResult> FindVerifiedBackupsAsync(
        IEnumerable<string> sourceFiles,
        string destinationRoot,
        CancellationToken cancellationToken = default)
    {
        var index = DestinationFileIndexer.Build(
            destinationRoot,
            _volumeProvider,
            requireExistingRoot: false,
            cancellationToken: cancellationToken);
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

                // Size is only a cheap prefilter. If no destination file has the same
                // size there is no possible byte-identical backup, so avoid reading the
                // source solely to compute a SHA-256 that cannot match anything.
                if (!index.FilesBySize.TryGetValue(sourceInfo.Length, out var candidates)) continue;

                var sourceHash = await _hasher.Sha256Async(source, cancellationToken).ConfigureAwait(false);

                foreach (var candidate in candidates)
                {
                    if (PathComparer.Instance.Equals(Path.GetFullPath(candidate), source)) continue;

                    // A destination file may be compared with several source files in
                    // one lookup. Cache its real SHA-256 only for this operation; the
                    // cache is never persisted and never reused by final reuse approval.
                    if (!destinationHashCache.TryGetValue(candidate, out var candidateHash))
                    {
                        try
                        {
                            candidateHash = await _hasher.Sha256Async(candidate, cancellationToken).ConfigureAwait(false);
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
}
