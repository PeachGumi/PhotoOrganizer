using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class SafeCopyRaceSafetyTests
{
    [TestMethod]
    public async Task FinalNameClaimedDuringFinalize_PreservesCompetitorAndUsesSuffix()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        File.WriteAllText(source, "irreplaceable-camera-data");

        var result = await new SafeCopyService(new ClaimFirstFinalNameDurabilityService())
            .CopyAsync(source, destinationDirectory);

        Assert.AreEqual(CopyStatus.Copied, result.Status, result.Error);
        Assert.AreEqual("irreplaceable-camera-data", File.ReadAllText(source));
        Assert.AreEqual("competing-user-data", File.ReadAllText(Path.Combine(destinationDirectory, "photo.jpg")));
        Assert.AreEqual("irreplaceable-camera-data", File.ReadAllText(Path.Combine(destinationDirectory, "photo_2.jpg")));
        Assert.IsFalse(Directory.EnumerateFiles(destinationDirectory, ".partial-*", SearchOption.TopDirectoryOnly).Any());
    }

    [TestMethod]
    public async Task FinalizationFailureBeforeMove_CleansOnlyOwnedPartialAndPreservesSource()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        File.WriteAllText(source, "irreplaceable-camera-data");

        var result = await new SafeCopyService(new FailBeforeMoveDurabilityService())
            .CopyAsync(source, destinationDirectory);

        Assert.AreEqual(CopyStatus.Failed, result.Status);
        Assert.AreEqual("irreplaceable-camera-data", File.ReadAllText(source));
        Assert.IsFalse(File.Exists(Path.Combine(destinationDirectory, "photo.jpg")));
        Assert.IsFalse(Directory.EnumerateFiles(destinationDirectory, ".partial-*", SearchOption.TopDirectoryOnly).Any());
    }

    [TestMethod]
    public async Task ExistingByteIdenticalFileWithoutDurabilityProof_IsNotAcceptedAsDuplicate()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        var destination = Path.Combine(destinationDirectory, "photo.jpg");
        File.WriteAllText(source, "same-bytes");
        File.WriteAllText(destination, "same-bytes");

        var result = await new SafeCopyService(new DuplicateDurabilityFailureService())
            .CopyAsync(source, destinationDirectory);

        Assert.AreEqual(CopyStatus.Failed, result.Status);
        Assert.AreEqual(destination, result.DestinationPath);
        Assert.AreEqual("same-bytes", File.ReadAllText(source));
        Assert.AreEqual("same-bytes", File.ReadAllText(destination));
        StringAssert.Contains(result.Error ?? string.Empty, "simulated duplicate durability failure");
    }

    private sealed class ClaimFirstFinalNameDurabilityService : IFileDurabilityService
    {
        private int _calls;

        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc)
        {
            if (Interlocked.Increment(ref _calls) == 1)
            {
                File.WriteAllText(finalPath, "competing-user-data");
                return new FinalizeFileResult(FinalizeFileStatus.DestinationExists, false);
            }

            File.Move(temporaryPath, finalPath, overwrite: false);
            File.SetLastWriteTimeUtc(finalPath, lastWriteUtc);
            return new FinalizeFileResult(FinalizeFileStatus.Committed, true);
        }

        public DurabilityResult EnsureDurable(string filePath) => new(true);
    }

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
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PhotoOrganizerSafeCopyRace-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
