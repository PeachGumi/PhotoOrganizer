using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class HiddenMediaSafetyTests
{
    [TestMethod]
    public void HiddenDirectorySupportedMedia_IsIncludedInCompleteScan()
    {
        using var temp = new TempDirectory();
        var hidden = Directory.CreateDirectory(Path.Combine(temp.Path, ".camera-hidden"));
        var photo = Path.Combine(hidden.FullName, "recoverable.jpg");
        File.WriteAllText(photo, "camera-bytes");

        var result = new MediaScanner(new MediaClassifier()).Scan(temp.Path);

        Assert.IsTrue(result.IsComplete, string.Join(Environment.NewLine, result.Errors));
        CollectionAssert.Contains(result.Files.ToList(), photo);
    }

    [TestMethod]
    public void HiddenDirectoryZeroByteSupportedMedia_BlocksCompleteScan()
    {
        using var temp = new TempDirectory();
        var hidden = Directory.CreateDirectory(Path.Combine(temp.Path, ".camera-hidden"));
        var photo = Path.Combine(hidden.FullName, "broken.nef");
        File.WriteAllBytes(photo, []);

        var result = new MediaScanner(new MediaClassifier()).Scan(temp.Path);

        Assert.IsFalse(result.IsComplete);
        CollectionAssert.Contains(result.Files.ToList(), photo);
        Assert.IsTrue(result.Errors.Any(error => error.Contains("zero bytes", StringComparison.OrdinalIgnoreCase)));
    }

    [TestMethod]
    public void HiddenUnsupportedSidecar_DoesNotBecomeImportScope()
    {
        using var temp = new TempDirectory();
        var hidden = Directory.CreateDirectory(Path.Combine(temp.Path, ".camera-hidden"));
        File.WriteAllText(Path.Combine(hidden.FullName, "metadata.xmp"), "sidecar");
        var photo = Path.Combine(temp.Path, "visible.jpg");
        File.WriteAllText(photo, "camera-bytes");

        var result = new MediaScanner(new MediaClassifier()).Scan(temp.Path);

        Assert.IsTrue(result.IsComplete, string.Join(Environment.NewLine, result.Errors));
        Assert.AreEqual(1, result.Files.Count);
        Assert.AreEqual(photo, result.Files.Single());
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PhotoOrganizerHiddenMedia-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
