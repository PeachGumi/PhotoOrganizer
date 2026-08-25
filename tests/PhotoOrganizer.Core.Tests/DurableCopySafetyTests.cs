using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class DurableCopySafetyTests
{
    [TestMethod]
    public async Task DurabilityFailureAfterFinalMove_ReturnsFailedWithoutDeletingFinalizedFile()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "DSC_0001.jpg");
        File.WriteAllText(source, "camera-data-that-must-survive");

        var result = await new SafeCopyService(new FailAfterMoveDurabilityService())
            .CopyAsync(source, destinationDirectory);

        Assert.AreEqual(CopyStatus.Failed, result.Status);
        Assert.IsNotNull(result.DestinationPath);
        Assert.IsTrue(File.Exists(result.DestinationPath));
        Assert.AreEqual("camera-data-that-must-survive", File.ReadAllText(result.DestinationPath));
        Assert.IsFalse(Directory.EnumerateFiles(destinationDirectory, ".partial-*", SearchOption.TopDirectoryOnly).Any());
    }

    [TestMethod]
    public async Task ExistingDirectoryWithRequestedFileName_IsPreservedAndUsesCollisionSuffix()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "DSC_0001.jpg");
        var conflictingDirectory = Directory.CreateDirectory(Path.Combine(destinationDirectory, "DSC_0001.jpg")).FullName;
        File.WriteAllText(source, "camera-data");

        var result = await new SafeCopyService().CopyAsync(source, destinationDirectory);

        Assert.AreEqual(CopyStatus.Copied, result.Status, result.Error);
        Assert.IsTrue(Directory.Exists(conflictingDirectory));
        Assert.AreEqual(
            "camera-data",
            File.ReadAllText(Path.Combine(destinationDirectory, "DSC_0001_2.jpg")));
    }

    [TestMethod]
    public async Task FinalReuseVerification_DurabilityFailureLeavesMatchingBytesUnverified()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        var destination = Path.Combine(destinationDirectory, "photo.jpg");
        File.WriteAllText(source, "identical-bytes");
        File.WriteAllText(destination, "identical-bytes");

        var result = await new FormatSafetyVerifier(
                new MediaClassifier(),
                durability: new AlwaysFailDurabilityService())
            .VerifyAsync([source], destinationDirectory);

        Assert.IsFalse(result.IsSafe);
        Assert.AreEqual(0, result.Verified);
        CollectionAssert.Contains(result.UnverifiedFiles.ToList(), source);
        Assert.IsTrue(result.Errors.Any(error => error.Contains("simulated durability failure", StringComparison.Ordinal)));
    }

    [TestMethod]
    public async Task FinalReuseVerification_DestinationMutationDuringDurabilityIsRehashedAndBlocked()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        var destination = Path.Combine(destinationDirectory, "photo.jpg");
        File.WriteAllText(source, "identical-before-durability");
        File.WriteAllText(destination, "identical-before-durability");

        var result = await new FormatSafetyVerifier(
                new MediaClassifier(),
                durability: new MutateDuringDurabilityService())
            .VerifyAsync([source], destinationDirectory);

        Assert.IsFalse(result.IsSafe);
        Assert.AreEqual(0, result.Verified);
        CollectionAssert.Contains(result.UnverifiedFiles.ToList(), source);
        Assert.AreEqual("mutated-during-durability", File.ReadAllText(destination));
        Assert.IsTrue(result.Errors.Any(error =>
            error.Contains("changed during durable verification", StringComparison.Ordinal)));
    }

    private sealed class FailAfterMoveDurabilityService : IFileDurabilityService
    {
        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc)
        {
            File.Move(temporaryPath, finalPath, overwrite: false);
            File.SetLastWriteTimeUtc(finalPath, lastWriteUtc);
            return new FinalizeFileResult(
                FinalizeFileStatus.Failed,
                FinalPathCreated: true,
                "simulated durability failure after final move");
        }

        public DurabilityResult EnsureDurable(string filePath) =>
            new(false, "simulated durability failure");
    }

    private sealed class AlwaysFailDurabilityService : IFileDurabilityService
    {
        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc) =>
            throw new NotSupportedException();

        public DurabilityResult EnsureDurable(string filePath) =>
            new(false, "simulated durability failure");
    }

    private sealed class MutateDuringDurabilityService : IFileDurabilityService
    {
        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc) =>
            throw new NotSupportedException();

        public DurabilityResult EnsureDurable(string filePath)
        {
            File.WriteAllText(filePath, "mutated-during-durability");
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
            stream.Flush(flushToDisk: true);
            return new DurabilityResult(true);
        }
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PhotoOrganizerDurabilityTests-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try
            {
                Directory.Delete(Path, recursive: true);
            }
            catch
            {
                // Test cleanup only. Never used for application/user data.
            }
        }
    }
}
