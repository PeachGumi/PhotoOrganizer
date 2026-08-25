using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class DataSafetyTests
{
    [TestMethod]
    public void StandardFormatsCannotBeReclassifiedAsRaw()
    {
        var classifier = new MediaClassifier([".jpg", ".mp4", ".nef"]);

        Assert.AreEqual(MediaKind.Jpeg, classifier.Classify("DSC_0001.JPG"));
        Assert.AreEqual(MediaKind.Video, classifier.Classify("clip.mp4"));
        Assert.AreEqual(MediaKind.Raw, classifier.Classify("DSC_0001.NEF"));
    }

    [TestMethod]
    public void ScannerMarksZeroByteSupportedMediaIncomplete()
    {
        using var temp = new TempDirectory();
        File.WriteAllBytes(Path.Combine(temp.Path, "broken.jpg"), []);
        File.WriteAllText(Path.Combine(temp.Path, "sidecar.xmp"), "ignored");

        var result = new MediaScanner(new MediaClassifier()).Scan(temp.Path);

        Assert.IsFalse(result.IsComplete);
        Assert.AreEqual(1, result.Files.Count);
        StringAssert.Contains(result.Errors.Single(), "zero bytes");
    }

    [TestMethod]
    public async Task ExistingDifferentDestinationIsNeverOverwritten()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "DSC_0001.jpg");
        var existing = Path.Combine(destinationDirectory, "DSC_0001.jpg");
        File.WriteAllText(source, "new-camera-data");
        File.WriteAllText(existing, "existing-library-data");

        var result = await new SafeCopyService().CopyAsync(source, destinationDirectory);

        Assert.AreEqual(CopyStatus.Copied, result.Status);
        Assert.AreEqual("existing-library-data", File.ReadAllText(existing));
        Assert.AreEqual("new-camera-data", File.ReadAllText(Path.Combine(destinationDirectory, "DSC_0001_2.jpg")));
    }

    [TestMethod]
    public async Task IdenticalExistingBytesAreSkippedWithoutCreatingCollisionCopy()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "DSC_0001.nef");
        var existing = Path.Combine(destinationDirectory, "DSC_0001.nef");
        File.WriteAllText(source, "same-real-bytes");
        File.WriteAllText(existing, "same-real-bytes");

        var result = await new SafeCopyService().CopyAsync(source, destinationDirectory);

        Assert.AreEqual(CopyStatus.SkippedDuplicate, result.Status);
        Assert.AreEqual(existing, result.DestinationPath);
        Assert.IsFalse(File.Exists(Path.Combine(destinationDirectory, "DSC_0001_2.nef")));
    }

    [TestMethod]
    public async Task ZeroByteSupportedMediaCannotBeCopiedAsSuccess()
    {
        using var temp = new TempDirectory();
        var source = Path.Combine(temp.Path, "empty.jpg");
        var destination = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        File.WriteAllBytes(source, []);

        var result = await new SafeCopyService().CopyAsync(source, destination);

        Assert.AreEqual(CopyStatus.Failed, result.Status);
        Assert.AreEqual(0, Directory.EnumerateFiles(destination).Count());
    }

    [TestMethod]
    public async Task FreshByteVerificationApprovesIndependentMatchingCopy()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        File.WriteAllText(source, "verified-data");
        File.WriteAllText(Path.Combine(destinationDirectory, "renamed-copy.jpg"), "verified-data");

        var result = await new FormatSafetyVerifier(new MediaClassifier())
            .VerifyAsync([source], destinationDirectory);

        Assert.IsTrue(result.IsSafe);
        Assert.AreEqual(1, result.Total);
        Assert.AreEqual(1, result.Verified);
    }

    [TestMethod]
    public async Task SameNameAndSizeWithDifferentBytesDoesNotVerify()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "source")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        File.WriteAllText(source, "AAAA");
        File.WriteAllText(Path.Combine(destinationDirectory, "photo.jpg"), "BBBB");

        var result = await new FormatSafetyVerifier(new MediaClassifier())
            .VerifyAsync([source], destinationDirectory);

        Assert.IsFalse(result.IsSafe);
        CollectionAssert.Contains(result.UnverifiedFiles.ToList(), source);
    }

    [TestMethod]
    public async Task SourcePathItselfCanNeverProveBackupExists()
    {
        using var temp = new TempDirectory();
        var source = Path.Combine(temp.Path, "photo.jpg");
        File.WriteAllText(source, "camera-data");

        var result = await new FormatSafetyVerifier(new MediaClassifier())
            .VerifyAsync([source], temp.Path);

        Assert.IsFalse(result.IsSafe);
        Assert.AreEqual(0, result.Verified);
    }

    [TestMethod]
    public async Task UnsupportedFilesDoNotCreateFalseSafeState()
    {
        using var temp = new TempDirectory();
        var source = Path.Combine(temp.Path, "metadata.xmp");
        File.WriteAllText(source, "sidecar");

        var result = await new FormatSafetyVerifier(new MediaClassifier())
            .VerifyAsync([source], temp.Path);

        Assert.IsFalse(result.IsSafe);
        Assert.AreEqual(0, result.Total);
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PhotoOrganizerTests-{Guid.NewGuid():N}");
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
