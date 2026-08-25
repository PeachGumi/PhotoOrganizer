using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class VolumeTraversalSafetyTests
{
    [TestMethod]
    public void CameraScan_DoesNotDescendIntoNestedMountedVolume()
    {
        using var temp = new TempDirectory();
        var card = Path.Combine(temp.Path, "card");
        var nested = Path.Combine(card, "OTHER_VOLUME");
        Directory.CreateDirectory(Path.Combine(card, "DCIM"));
        Directory.CreateDirectory(nested);
        var expected = Path.Combine(card, "DCIM", "photo.jpg");
        var excluded = Path.Combine(nested, "other.jpg");
        File.WriteAllText(expected, "card-data");
        File.WriteAllText(excluded, "other-volume-data");

        var provider = new BoundaryProvider([
            new MountedVolumeInfo(card, "card", true, false),
            new MountedVolumeInfo(nested, "nested", true, false)
        ]);

        var result = new MediaScanner(new MediaClassifier(), provider).Scan(card);

        Assert.IsTrue(result.IsComplete);
        Assert.AreEqual(1, result.Files.Count);
        Assert.AreEqual(Path.GetFullPath(expected), result.Files.Single());
    }

    [TestMethod]
    public async Task NestedMountedVolumeCannotProveDestinationBackup()
    {
        using var temp = new TempDirectory();
        var sourceRoot = Path.Combine(temp.Path, "source");
        var destination = Path.Combine(temp.Path, "destination");
        var nested = Path.Combine(destination, "MOUNTED_CAMERA");
        Directory.CreateDirectory(sourceRoot);
        Directory.CreateDirectory(nested);
        var source = Path.Combine(sourceRoot, "photo.jpg");
        var falseProof = Path.Combine(nested, "photo.jpg");
        File.WriteAllText(source, "same-bytes");
        File.WriteAllText(falseProof, "same-bytes");

        var provider = new BoundaryProvider([
            new MountedVolumeInfo(destination, "destination", false, false),
            new MountedVolumeInfo(nested, "nested-camera", true, false)
        ]);

        var result = await new FormatSafetyVerifier(new MediaClassifier(), provider)
            .VerifyAsync([source], destination);

        Assert.IsFalse(result.IsSafe);
        Assert.AreEqual(0, result.Verified);
        CollectionAssert.Contains(result.UnverifiedFiles.ToList(), Path.GetFullPath(source));
    }

    private sealed class BoundaryProvider(IReadOnlyList<MountedVolumeInfo> volumes) : IStorageVolumeProvider
    {
        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes() => volumes;

        public MountedVolumeInfo? ResolveVolumeForPath(string path) => volumes
            .Where(v => PathSafety.IsSameOrDescendant(path, v.RootPath, PathComparison))
            .OrderByDescending(v => PathSafety.Normalize(v.RootPath).Length)
            .FirstOrDefault();
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PhotoOrganizerBoundaryTests-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
