using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class SafeCopyPipelineOptimizationTests
{
    [TestMethod]
    public async Task ConcurrentSameNameSameBytes_ProducesOneCopyAndOneDuplicateWithoutPartials()
    {
        using var temp = new TempDirectory();
        var sourceA = temp.CreateSource("card-a", "IMG_0001.JPG");
        var sourceB = temp.CreateSource("card-b", "IMG_0001.JPG");
        var payload = CreatePayload('A', 128 * 1024);
        File.WriteAllBytes(sourceA, payload);
        File.WriteAllBytes(sourceB, payload);
        var sourceABefore = File.ReadAllBytes(sourceA);
        var sourceBBefore = File.ReadAllBytes(sourceB);

        var results = await Task.WhenAll(
            new SafeCopyService().CopyAsync(sourceA, temp.Destination),
            new SafeCopyService().CopyAsync(sourceB, temp.Destination));

        Assert.AreEqual(1, results.Count(result => result.Status == CopyStatus.Copied));
        Assert.AreEqual(1, results.Count(result => result.Status == CopyStatus.SkippedDuplicate));
        Assert.IsTrue(results.All(result => result.Error is null), string.Join(Environment.NewLine, results.Select(result => result.Error)));
        Assert.IsTrue(File.Exists(Path.Combine(temp.Destination, "IMG_0001.JPG")));
        Assert.IsFalse(File.Exists(Path.Combine(temp.Destination, "IMG_0001_2.JPG")));
        Assert.AreEqual(1, Directory.EnumerateFiles(temp.Destination).Count());
        Assert.IsFalse(Directory.EnumerateFiles(temp.Destination)
            .Any(path => Path.GetFileName(path).StartsWith(".partial-", StringComparison.Ordinal)));
        CollectionAssert.AreEqual(payload, File.ReadAllBytes(Path.Combine(temp.Destination, "IMG_0001.JPG")));
        CollectionAssert.AreEqual(sourceABefore, File.ReadAllBytes(sourceA));
        CollectionAssert.AreEqual(sourceBBefore, File.ReadAllBytes(sourceB));
    }

    [TestMethod]
    public async Task ConcurrentSameNameDifferentBytes_PreservesBothCopiesWithoutOverwrite()
    {
        using var temp = new TempDirectory();
        var sourceA = temp.CreateSource("card-a", "IMG_0002.JPG");
        var sourceB = temp.CreateSource("card-b", "IMG_0002.JPG");
        var payloadA = CreatePayload('A', 128 * 1024);
        var payloadB = CreatePayload('B', 128 * 1024);
        File.WriteAllBytes(sourceA, payloadA);
        File.WriteAllBytes(sourceB, payloadB);
        var sourceABefore = File.ReadAllBytes(sourceA);
        var sourceBBefore = File.ReadAllBytes(sourceB);

        var results = await Task.WhenAll(
            new SafeCopyService().CopyAsync(sourceA, temp.Destination),
            new SafeCopyService().CopyAsync(sourceB, temp.Destination));

        Assert.IsTrue(results.All(result => result.Status == CopyStatus.Copied),
            string.Join(Environment.NewLine, results.Select(result => result.Error)));
        var destinations = Directory.EnumerateFiles(temp.Destination)
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        CollectionAssert.AreEquivalent(
            new[]
            {
                Path.Combine(temp.Destination, "IMG_0002.JPG"),
                Path.Combine(temp.Destination, "IMG_0002_2.JPG")
            },
            destinations);

        var destinationPayloads = destinations.Select(File.ReadAllBytes).ToList();
        Assert.IsTrue(destinationPayloads.Any(payload => payload.SequenceEqual(payloadA)));
        Assert.IsTrue(destinationPayloads.Any(payload => payload.SequenceEqual(payloadB)));
        Assert.IsFalse(Directory.EnumerateFiles(temp.Destination)
            .Any(path => Path.GetFileName(path).StartsWith(".partial-", StringComparison.Ordinal)));
        CollectionAssert.AreEqual(sourceABefore, File.ReadAllBytes(sourceA));
        CollectionAssert.AreEqual(sourceBBefore, File.ReadAllBytes(sourceB));
    }

    [TestMethod]
    public async Task ExistingCollision_IsNeverOverwrittenAndUsesSuffix()
    {
        using var temp = new TempDirectory();
        var source = temp.CreateSource("card", "IMG_0003.JPG");
        var sourcePayload = CreatePayload('S', 128 * 1024);
        var existingPayload = CreatePayload('E', 128 * 1024);
        File.WriteAllBytes(source, sourcePayload);
        var existing = Path.Combine(temp.Destination, "IMG_0003.JPG");
        File.WriteAllBytes(existing, existingPayload);
        var sourceBefore = File.ReadAllBytes(source);

        var result = await new SafeCopyService().CopyAsync(source, temp.Destination);

        Assert.AreEqual(CopyStatus.Copied, result.Status, result.Error);
        Assert.AreEqual(Path.Combine(temp.Destination, "IMG_0003_2.JPG"), result.DestinationPath);
        CollectionAssert.AreEqual(existingPayload, File.ReadAllBytes(existing));
        CollectionAssert.AreEqual(sourcePayload, File.ReadAllBytes(result.DestinationPath!));
        CollectionAssert.AreEqual(sourceBefore, File.ReadAllBytes(source));
        Assert.IsFalse(Directory.EnumerateFiles(temp.Destination)
            .Any(path => Path.GetFileName(path).StartsWith(".partial-", StringComparison.Ordinal)));
    }

    [TestMethod]
    public async Task FinalizationFailureBeforeMove_CleansOwnedPartialAndPreservesSource()
    {
        using var temp = new TempDirectory();
        var source = temp.CreateSource("card", "IMG_0004.JPG");
        var payload = CreatePayload('F', 128 * 1024);
        File.WriteAllBytes(source, payload);

        var result = await new SafeCopyService(new FailBeforeMoveDurabilityService())
            .CopyAsync(source, temp.Destination);

        Assert.AreEqual(CopyStatus.Failed, result.Status);
        Assert.IsFalse(File.Exists(Path.Combine(temp.Destination, "IMG_0004.JPG")));
        Assert.IsFalse(Directory.EnumerateFiles(temp.Destination)
            .Any(path => Path.GetFileName(path).StartsWith(".partial-", StringComparison.Ordinal)));
        CollectionAssert.AreEqual(payload, File.ReadAllBytes(source));
    }

    [TestMethod]
    public async Task DuplicateDurabilityFailure_CleansTemporaryAndPreservesBothCopies()
    {
        using var temp = new TempDirectory();
        var source = temp.CreateSource("card", "IMG_0005.JPG");
        var payload = CreatePayload('D', 128 * 1024);
        File.WriteAllBytes(source, payload);
        var destination = Path.Combine(temp.Destination, "IMG_0005.JPG");
        File.WriteAllBytes(destination, payload);
        var sourceBefore = File.ReadAllBytes(source);

        var result = await new SafeCopyService(new DuplicateDurabilityFailureService())
            .CopyAsync(source, temp.Destination);

        Assert.AreEqual(CopyStatus.Failed, result.Status);
        Assert.AreEqual(destination, result.DestinationPath);
        StringAssert.Contains(result.Error ?? string.Empty, "simulated duplicate durability failure");
        CollectionAssert.AreEqual(payload, File.ReadAllBytes(destination));
        CollectionAssert.AreEqual(sourceBefore, File.ReadAllBytes(source));
        Assert.IsFalse(Directory.EnumerateFiles(temp.Destination)
            .Any(path => Path.GetFileName(path).StartsWith(".partial-", StringComparison.Ordinal)));
    }

    private static byte[] CreatePayload(char value, int length) =>
        Encoding.ASCII.GetBytes(new string(value, length));

    private sealed class FailBeforeMoveDurabilityService : IFileDurabilityService
    {
        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc) =>
            new(FinalizeFileStatus.Failed, false, "simulated failure before move");

        public DurabilityResult EnsureDurable(string filePath) => new(true);
    }

    private sealed class DuplicateDurabilityFailureService : IFileDurabilityService
    {
        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc) =>
            throw new NotSupportedException();

        public DurabilityResult EnsureDurable(string filePath) =>
            new(false, "simulated duplicate durability failure");
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerSafeCopyPipeline-{Guid.NewGuid():N}");
            Destination = Directory.CreateDirectory(Path.Combine(Root, "destination")).FullName;
        }

        public string Root { get; }
        public string Destination { get; }

        public string CreateSource(string directoryName, string fileName)
        {
            var directory = Directory.CreateDirectory(Path.Combine(Root, directoryName, "DCIM"));
            return Path.Combine(directory.FullName, fileName);
        }

        public void Dispose()
        {
            try { Directory.Delete(Root, recursive: true); } catch { }
        }
    }
}
