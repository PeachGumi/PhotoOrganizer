using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class IoOptimizationTests
{
    [TestMethod]
    public async Task DuplicateLookupDoesNotHashSourceWithoutSameSizeCandidate()
    {
        using var temp = new TempDirectory();
        var source = Path.Combine(temp.Source, "photo.jpg");
        var differentSize = Path.Combine(temp.Destination, "other.jpg");
        File.WriteAllText(source, "source-bytes");
        File.WriteAllText(differentSize, "different-size-destination-bytes");

        var hasher = new CountingHasher();
        var result = await new DestinationLibrary(hasher: hasher)
            .FindVerifiedBackupsAsync([source], temp.Destination);

        Assert.AreEqual(0, result.MatchedSources.Count);
        Assert.AreEqual(0, hasher.CallCount(source));
        Assert.AreEqual(0L, hasher.BytesRead);
    }

    [TestMethod]
    public async Task DuplicateLookupHashesSharedDestinationCandidateOnlyOncePerOperation()
    {
        using var temp = new TempDirectory();
        var source1 = Path.Combine(temp.Source, "one.jpg");
        var source2 = Path.Combine(temp.Source, "two.jpg");
        var destination = Path.Combine(temp.Destination, "existing.jpg");
        File.WriteAllText(source1, "same-payload");
        File.WriteAllText(source2, "same-payload");
        File.WriteAllText(destination, "same-payload");

        var hasher = new CountingHasher();
        var result = await new DestinationLibrary(hasher: hasher)
            .FindVerifiedBackupsAsync([source1, source2], temp.Destination);

        Assert.AreEqual(2, result.MatchedSources.Count);
        Assert.AreEqual(1, hasher.CallCount(source1));
        Assert.AreEqual(1, hasher.CallCount(source2));
        Assert.AreEqual(1, hasher.CallCount(destination));
    }

    [TestMethod]
    public async Task FinalVerifierDoesNotHashSourceWithoutSameSizeCandidate()
    {
        using var temp = new TempDirectory();
        var source = Path.Combine(temp.Source, "photo.jpg");
        var differentSize = Path.Combine(temp.Destination, "other.jpg");
        File.WriteAllText(source, "source-bytes");
        File.WriteAllText(differentSize, "different-size-destination-bytes");

        var hasher = new CountingHasher();
        var result = await new FormatSafetyVerifier(new MediaClassifier(), hasher: hasher)
            .VerifyAsync([source], temp.Destination);

        Assert.IsFalse(result.IsSafe);
        Assert.AreEqual(0, hasher.CallCount(source));
        Assert.AreEqual(0L, hasher.BytesRead);
    }

    [TestMethod]
    public async Task FinalVerifierCacheIsDiscardedBetweenFreshVerificationRuns()
    {
        using var temp = new TempDirectory();
        var source1 = Path.Combine(temp.Source, "one.jpg");
        var source2 = Path.Combine(temp.Source, "two.jpg");
        var destination = Path.Combine(temp.Destination, "existing.jpg");
        File.WriteAllText(source1, "same-payload");
        File.WriteAllText(source2, "same-payload");
        File.WriteAllText(destination, "same-payload");

        var hasher = new CountingHasher();
        var verifier = new FormatSafetyVerifier(new MediaClassifier(), hasher: hasher);

        var first = await verifier.VerifyAsync([source1, source2], temp.Destination);
        var second = await verifier.VerifyAsync([source1, source2], temp.Destination);

        Assert.IsTrue(first.IsSafe);
        Assert.IsTrue(second.IsSafe);
        Assert.AreEqual(2, hasher.CallCount(destination),
            "Destination bytes must be freshly hashed on every reuse-verification invocation.");
    }

    private sealed class CountingHasher : IFileHasher
    {
        private readonly Dictionary<string, int> _calls = new(StringComparer.OrdinalIgnoreCase);

        public long BytesRead { get; private set; }

        public int CallCount(string path) =>
            _calls.TryGetValue(Path.GetFullPath(path), out var count) ? count : 0;

        public async Task<string> Sha256Async(string path, CancellationToken cancellationToken = default)
        {
            var full = Path.GetFullPath(path);
            _calls[full] = CallCount(full) + 1;
            BytesRead += new FileInfo(full).Length;
            return await Hashing.Sha256Async(full, cancellationToken);
        }
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Root = Path.Combine(Path.GetTempPath(), "PhotoOrganizerIoTests-" + Guid.NewGuid().ToString("N"));
            Source = Directory.CreateDirectory(Path.Combine(Root, "source")).FullName;
            Destination = Directory.CreateDirectory(Path.Combine(Root, "destination")).FullName;
        }

        public string Root { get; }
        public string Source { get; }
        public string Destination { get; }

        public void Dispose()
        {
            try { Directory.Delete(Root, recursive: true); } catch { }
        }
    }
}
