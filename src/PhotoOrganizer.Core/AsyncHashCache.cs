using System.Collections.Concurrent;

namespace PhotoOrganizer.Core;

/// <summary>
/// Lazily computes one destination prefilter hash per path for a single lookup
/// or verification call. A Lazy wrapper is required because ConcurrentDictionary
/// value factories can run more than once under contention.
/// </summary>
internal sealed class AsyncHashCache
{
    private readonly IFileHasher _hasher;
    private readonly ConcurrentDictionary<string, Lazy<Task<string>>> _entries =
        new(PathComparer.Instance);

    public AsyncHashCache(IFileHasher hasher)
    {
        _hasher = hasher;
    }

    public Task<string> GetAsync(string path, CancellationToken cancellationToken)
    {
        var entry = _entries.GetOrAdd(
            path,
            candidate => new Lazy<Task<string>>(
                () => _hasher.Sha256Async(candidate, cancellationToken),
                LazyThreadSafetyMode.ExecutionAndPublication));
        return entry.Value;
    }
}
