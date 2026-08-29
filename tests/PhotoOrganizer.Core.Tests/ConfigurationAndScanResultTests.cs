using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class ConfigurationAndScanResultTests
{
    [TestMethod]
    public void ParseRawExtensions_NormalizesSeparatorsCaseAndDuplicates()
    {
        var extensions = MediaClassifier.ParseRawExtensions("NEF, .CR3; nef\nDNG  .raf");

        CollectionAssert.AreEqual(
            new[] { ".nef", ".cr3", ".dng", ".raf" },
            extensions);
    }

    [TestMethod]
    public void DefaultRawExtensions_AreAllClassifiedAsRaw()
    {
        var classifier = new MediaClassifier();

        foreach (var extension in MediaClassifier.DefaultRawExtensions)
        {
            Assert.AreEqual(MediaKind.Raw, classifier.Classify("photo" + extension), extension);
        }
    }

    [TestMethod]
    public void EmptyCameraCard_HasTypedNoSupportedMediaReason()
    {
        using var temp = new TempDirectory();
        Directory.CreateDirectory(Path.Combine(temp.Path, "DCIM"));
        var provider = new SingleVolumeProvider(temp.Path, "device-card");
        var tracker = new StorageSessionTracker(provider);
        var coordinator = new ImportCoordinator(
            new MediaClassifier(),
            tracker,
            new CameraCardRootResolver(provider),
            provider);

        var result = coordinator.ScanCard(temp.Path);

        Assert.IsFalse(result.IsReady);
        Assert.IsTrue(result.IsNoSupportedMedia);
        Assert.AreEqual(ScanFailureReason.NoSupportedMedia, result.FailureReason);
    }

    [TestMethod]
    public void MissingPhysicalIdentity_HasTypedFailureReason()
    {
        using var temp = new TempDirectory();
        Directory.CreateDirectory(Path.Combine(temp.Path, "DCIM"));
        File.WriteAllText(Path.Combine(temp.Path, "DCIM", "photo.jpg"), "camera-data");
        var provider = new SingleVolumeProvider(temp.Path, physicalDeviceFingerprint: null);
        var tracker = new StorageSessionTracker(provider);
        var coordinator = new ImportCoordinator(
            new MediaClassifier(),
            tracker,
            new CameraCardRootResolver(provider),
            provider);

        var result = coordinator.ScanCard(temp.Path);

        Assert.IsFalse(result.IsReady);
        Assert.AreEqual(ScanFailureReason.MissingPhysicalDeviceIdentity, result.FailureReason);
    }

    private sealed class SingleVolumeProvider(string root, string? physicalDeviceFingerprint) : IStorageVolumeProvider
    {
        private readonly MountedVolumeInfo _volume = new(
            PathSafety.Normalize(root),
            "test-volume",
            true,
            false,
            physicalDeviceFingerprint);

        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes() => [_volume];

        public MountedVolumeInfo? ResolveVolumeForPath(string path) =>
            PathSafety.IsSameOrDescendant(path, _volume.RootPath, PathComparison) ? _volume : null;
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PhotoOrganizerConfigTests-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
