using System.Collections.Concurrent;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class ParallelHashingOptimizationTests
{
    [TestMethod]
    public async Task DestinationLookup_ParallelSourcesAreBoundedAndSharePrefilterHash()
    {
        using var temp = new TempDirectory();
        var sources = Enumerable.Range(0, 8)
            .Select(index => temp.WriteSource($"source-{index:D2}.jpg", "same-payload"))
            .ToArray();
        var destination = temp.WriteDestination("existing.jpg", "same-payload");
        var hasher = new TrackingHasher(TimeSpan.FromMilliseconds(20));

        var result = await new DestinationLibrary(
                hasher: hasher,
                maxDegreeOfParallelism: 3)
            .FindVerifiedBackupsAsync(sources, temp.Destination);

        Assert.AreEqual(sources.Length, result.MatchedSources.Count);
        Assert.AreEqual(1, hasher.CallCount(destination),
            "The shared destination prefilter must be hashed once per lookup invocation.");
        Assert.IsTrue(hasher.MaxConcurrency >= 2,
            $"The parallel lookup did not overlap hashing (max={hasher.MaxConcurrency}).");
        Assert.IsTrue(hasher.MaxConcurrency <= 3,
            $"Hashing exceeded the configured bound (max={hasher.MaxConcurrency}).");
    }

    [TestMethod]
    public async Task FinalVerifier_ParallelSourcesKeepFreshProofAndReturnSafeResult()
    {
        using var temp = new TempDirectory();
        var sources = Enumerable.Range(0, 8)
            .Select(index => temp.WriteSource($"source-{index:D2}.jpg", "same-payload"))
            .ToArray();
        var destination = temp.WriteDestination("existing.jpg", "same-payload");
        var hasher = new TrackingHasher(TimeSpan.FromMilliseconds(20));
        var durability = new CountingDurabilityService();

        var result = await new FormatSafetyVerifier(
                new MediaClassifier(),
                hasher: hasher,
                durability: durability,
                maxDegreeOfParallelism: 3)
            .VerifyAsync(sources, temp.Destination);

        Assert.IsTrue(result.IsSafe, string.Join(Environment.NewLine, result.Errors));
        Assert.AreEqual(sources.Length, result.Verified);
        Assert.AreEqual(1 + sources.Length, hasher.CallCount(destination),
            "The destination prefilter is shared, while every source receives a fresh post-durability hash.");
        foreach (var source in sources)
        {
            Assert.AreEqual(2, hasher.CallCount(source),
                "Each source must be hashed before matching and again after destination durability.");
        }

        Assert.AreEqual(sources.Length, durability.EnsureDurableCount,
            "Each source match requires an independent fresh durable proof.");
        Assert.IsTrue(hasher.MaxConcurrency >= 2,
            $"The parallel verifier did not overlap hashing (max={hasher.MaxConcurrency}).");
        Assert.IsTrue(hasher.MaxConcurrency <= 3,
            $"Hashing exceeded the configured bound (max={hasher.MaxConcurrency}).");
    }

    private sealed class TrackingHasher(TimeSpan delay) : IFileHasher
    {
        private readonly ConcurrentDictionary<string, int> _calls = new(PathComparerForTests);
        private int _active;
        private int _maxConcurrency;

        public int MaxConcurrency => Volatile.Read(ref _maxConcurrency);

        public int CallCount(string path) =>
            _calls.TryGetValue(Path.GetFullPath(path), out var count) ? count : 0;

        public async Task<string> Sha256Async(
            string path,
            CancellationToken cancellationToken = default)
        {
            var fullPath = Path.GetFullPath(path);
            _calls.AddOrUpdate(fullPath, 1, static (_, count) => count + 1);

            var active = Interlocked.Increment(ref _active);
            while (true)
            {
                var observedMaximum = Volatile.Read(ref _maxConcurrency);
                if (active <= observedMaximum
                    || Interlocked.CompareExchange(ref _maxConcurrency, active, observedMaximum) == observedMaximum)
                {
                    break;
                }
            }

            try
            {
                await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                return await Hashing.Sha256Async(fullPath, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                Interlocked.Decrement(ref _active);
            }
        }

        private static StringComparer PathComparerForTests =>
            OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
    }

    private sealed class CountingDurabilityService : IFileDurabilityService
    {
        private int _ensureDurableCount;

        public int EnsureDurableCount => Volatile.Read(ref _ensureDurableCount);

        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc) =>
            throw new NotSupportedException();

        public DurabilityResult EnsureDurable(string filePath)
        {
            Interlocked.Increment(ref _ensureDurableCount);
            return new DurabilityResult(true);
        }
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Root = Path.Combine(Path.GetTempPath(), "PhotoOrganizerParallelHashTests-" + Guid.NewGuid().ToString("N"));
            Source = Directory.CreateDirectory(Path.Combine(Root, "source")).FullName;
            Destination = Directory.CreateDirectory(Path.Combine(Root, "destination")).FullName;
        }

        public string Root { get; }
        public string Source { get; }
        public string Destination { get; }

        public string WriteSource(string fileName, string content)
        {
            var path = Path.Combine(Source, fileName);
            File.WriteAllText(path, content);
            return path;
        }

        public string WriteDestination(string fileName, string content)
        {
            var path = Path.Combine(Destination, fileName);
            File.WriteAllText(path, content);
            return path;
        }

        public void Dispose()
        {
            try { Directory.Delete(Root, recursive: true); } catch { }
        }
    }
}
