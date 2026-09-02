using System.Security.Cryptography;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class HashingPipelineTests
{
    [TestMethod]
    public async Task DoubleBufferedHasher_MatchesFrameworkAcrossBufferBoundaries()
    {
        using var temp = new TempDirectory();
        var sizes = new[]
        {
            0,
            1,
            (8 * 1024 * 1024) - 1,
            (8 * 1024 * 1024) + 123
        };

        foreach (var size in sizes)
        {
            var payload = new byte[size];
            new Random(size + 17).NextBytes(payload);
            var path = Path.Combine(temp.Path, $"payload-{size}.bin");
            File.WriteAllBytes(path, payload);

            var expected = Convert.ToHexString(SHA256.HashData(payload)).ToLowerInvariant();
            var actual = await Hashing.Sha256Async(path);

            Assert.AreEqual(expected, actual, $"Unexpected SHA-256 for {size} bytes.");
        }
    }

    [TestMethod]
    public async Task DoubleBufferedHasher_HonorsPreCancelledToken()
    {
        using var temp = new TempDirectory();
        var path = Path.Combine(temp.Path, "payload.bin");
        File.WriteAllBytes(path, new byte[1024]);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(
            () => Hashing.Sha256Async(path, cancellation.Token));
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PhotoOrganizerHashPipeline-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
