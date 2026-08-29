using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class PathAliasSafetyTests
{
    [TestMethod]
    public void DirectFilesystemPath_IsAccepted()
    {
        using var temp = new TempDirectory();
        var direct = Directory.CreateDirectory(Path.Combine(temp.Path, "direct", "library")).FullName;

        Assert.IsTrue(
            PathSafety.TryValidateDirectFilesystemPath(direct, out var error),
            error);
    }

    [TestMethod]
    public void DirectorySymlink_IsRejectedAsDestinationPath()
    {
        using var temp = new TempDirectory();
        var target = Directory.CreateDirectory(Path.Combine(temp.Path, "target")).FullName;
        var alias = Path.Combine(temp.Path, "alias");
        CreateDirectorySymlinkOrSkip(alias, target);

        Assert.IsFalse(PathSafety.TryValidateDirectFilesystemPath(alias, out var error));
        StringAssert.Contains(error ?? string.Empty, "reparse");
    }

    [TestMethod]
    public void NestedPathThroughDirectorySymlink_IsRejected()
    {
        using var temp = new TempDirectory();
        var target = Directory.CreateDirectory(Path.Combine(temp.Path, "target")).FullName;
        var alias = Path.Combine(temp.Path, "alias");
        CreateDirectorySymlinkOrSkip(alias, target);

        var nested = Path.Combine(alias, "new-event", "JPG");
        Assert.IsFalse(PathSafety.TryValidateDirectFilesystemPath(nested, out var error));
        Assert.IsFalse(string.IsNullOrWhiteSpace(error));
    }

    [TestMethod]
    public async Task FormatVerifier_NeverAcceptsSourceThroughDestinationSymlink()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "camera")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        File.WriteAllText(source, "irreplaceable-camera-bytes");

        var alias = Path.Combine(temp.Path, "destination-alias");
        CreateDirectorySymlinkOrSkip(alias, sourceDirectory);

        var result = await new FormatSafetyVerifier(new MediaClassifier())
            .VerifyAsync([source], alias);

        Assert.IsFalse(result.IsSafe);
        Assert.AreEqual(0, result.Verified);
        Assert.IsTrue(result.Errors.Count > 0);
    }

    [TestMethod]
    public async Task SafeCopyService_RefusesAliasedDestinationWithoutWritingSource()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "camera")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        File.WriteAllText(source, "irreplaceable-camera-bytes");

        var alias = Path.Combine(temp.Path, "destination-alias");
        CreateDirectorySymlinkOrSkip(alias, sourceDirectory);

        var result = await new SafeCopyService().CopyAsync(source, alias);

        Assert.AreEqual(CopyStatus.Failed, result.Status);
        Assert.AreEqual("irreplaceable-camera-bytes", File.ReadAllText(source));
        Assert.IsFalse(File.Exists(Path.Combine(sourceDirectory, "photo_2.jpg")));
        Assert.IsFalse(Directory.EnumerateFiles(sourceDirectory).Any(path =>
            Path.GetFileName(path).StartsWith(".partial-", StringComparison.Ordinal)));
    }

    [TestMethod]
    public void MediaScanner_RejectsSymlinkRoot()
    {
        using var temp = new TempDirectory();
        var target = Directory.CreateDirectory(Path.Combine(temp.Path, "actual-card")).FullName;
        Directory.CreateDirectory(Path.Combine(target, "DCIM"));
        File.WriteAllText(Path.Combine(target, "DCIM", "photo.jpg"), "camera-data");
        var alias = Path.Combine(temp.Path, "card-alias");
        CreateDirectorySymlinkOrSkip(alias, target);

        var result = new MediaScanner(new MediaClassifier()).Scan(alias);

        Assert.IsFalse(result.IsComplete);
        Assert.AreEqual(0, result.Files.Count);
        Assert.IsTrue(result.Errors.Any(error => error.Contains("direct filesystem path", StringComparison.Ordinal)));
    }

    [TestMethod]
    public async Task SafeCopyService_FileSymlinkCollisionIsNeverAcceptedAsDuplicate()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "camera")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "destination")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        var alias = Path.Combine(destinationDirectory, "photo.jpg");
        File.WriteAllText(source, "camera-data");
        CreateFileSymlinkOrSkip(alias, source);

        var result = await new SafeCopyService().CopyAsync(source, destinationDirectory);

        Assert.AreEqual(CopyStatus.Copied, result.Status, result.Error);
        Assert.AreEqual(Path.Combine(destinationDirectory, "photo_2.jpg"), result.DestinationPath);
        Assert.AreEqual("camera-data", File.ReadAllText(source));
        Assert.AreEqual("camera-data", File.ReadAllText(result.DestinationPath!));
    }

    private static void CreateDirectorySymlinkOrSkip(string alias, string target)
    {
        try
        {
            Directory.CreateSymbolicLink(alias, target);
        }
        catch (Exception ex) when (ex is PlatformNotSupportedException
                                   or UnauthorizedAccessException
                                   or IOException)
        {
            Assert.Inconclusive($"Directory symbolic links are unavailable in this test environment: {ex.Message}");
        }
    }

    private static void CreateFileSymlinkOrSkip(string alias, string target)
    {
        try
        {
            File.CreateSymbolicLink(alias, target);
        }
        catch (Exception ex) when (ex is PlatformNotSupportedException
                                   or UnauthorizedAccessException
                                   or IOException)
        {
            Assert.Inconclusive($"File symbolic links are unavailable in this test environment: {ex.Message}");
        }
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PhotoOrganizerPathAlias-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
