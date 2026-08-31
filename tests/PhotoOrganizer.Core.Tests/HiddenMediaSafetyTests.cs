using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class HiddenMediaSafetyTests
{
    [TestMethod]
    public void DotPrefixedDirectory_IsExcludedFromCompleteScan()
    {
        using var temp = new TempDirectory();
        var hidden = Directory.CreateDirectory(Path.Combine(temp.Path, ".Spotlight-V100"));
        File.WriteAllText(Path.Combine(hidden.FullName, "metadata.jpg"), "metadata");
        var photo = Path.Combine(temp.Path, "visible.jpg");
        File.WriteAllText(photo, "camera-bytes");

        var result = new MediaScanner(new MediaClassifier()).Scan(temp.Path);

        Assert.IsTrue(result.IsComplete, string.Join(Environment.NewLine, result.Errors));
        CollectionAssert.AreEqual(new[] { photo }, result.Files.ToArray());
    }

    [TestMethod]
    public void DotPrefixedSupportedFile_IsExcludedFromCompleteScan()
    {
        using var temp = new TempDirectory();
        File.WriteAllText(Path.Combine(temp.Path, "._DSC_0001.JPG"), "apple-double");
        var photo = Path.Combine(temp.Path, "DSC_0001.JPG");
        File.WriteAllText(photo, "camera-bytes");

        var result = new MediaScanner(new MediaClassifier()).Scan(temp.Path);

        Assert.IsTrue(result.IsComplete, string.Join(Environment.NewLine, result.Errors));
        CollectionAssert.AreEqual(new[] { photo }, result.Files.ToArray());
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
