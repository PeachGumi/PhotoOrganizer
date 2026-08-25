using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class StorageIdentityTests
{
    [TestMethod]
    public void Capture_MatchesWithinSameMountSession()
    {
        using var temp = new TempDirectory();
        var provider = new FakeProvider(temp.Path, "volume-a");
        var tracker = new StorageSessionTracker(provider);

        var snapshot = tracker.Capture(temp.Path);

        Assert.IsNotNull(snapshot);
        Assert.IsTrue(tracker.Matches(snapshot, temp.Path));
        Assert.AreEqual("physical-a", snapshot.PhysicalDeviceFingerprint);
    }

    [TestMethod]
    public void Capture_UsesOneMountedSnapshotAndDoesNotResolveAgain()
    {
        using var temp = new TempDirectory();
        var provider = new FakeProvider(temp.Path, "volume-a");
        var tracker = new StorageSessionTracker(provider);

        var snapshot = tracker.Capture(Path.Combine(temp.Path, "future", "destination"));

        Assert.IsNotNull(snapshot);
        Assert.AreEqual(1, provider.GetMountedVolumesCallCount);
        Assert.AreEqual(0, provider.ResolveVolumeForPathCallCount);

        Assert.IsTrue(tracker.Matches(snapshot, temp.Path));
        Assert.AreEqual(2, provider.GetMountedVolumesCallCount,
            "Matches should perform one fresh mounted-volume snapshot, not two enumerations.");
        Assert.AreEqual(0, provider.ResolveVolumeForPathCallCount);
    }

    [TestMethod]
    public void RemovalAndReinsert_InvalidatesOldSessionEvenForSameFingerprint()
    {
        using var temp = new TempDirectory();
        var provider = new FakeProvider(temp.Path, "volume-a");
        var tracker = new StorageSessionTracker(provider);
        var snapshot = tracker.Capture(temp.Path)!;

        provider.Mounted = false;
        tracker.Refresh();
        provider.Mounted = true;
        tracker.Refresh();

        Assert.IsFalse(tracker.Matches(snapshot, temp.Path));
        var replacement = tracker.Capture(temp.Path);
        Assert.IsNotNull(replacement);
        Assert.AreNotEqual(snapshot.SessionId, replacement.SessionId);
    }

    [TestMethod]
    public void SamePathDifferentVolume_InvalidatesOldSession()
    {
        using var temp = new TempDirectory();
        var provider = new FakeProvider(temp.Path, "volume-a");
        var tracker = new StorageSessionTracker(provider);
        var snapshot = tracker.Capture(temp.Path)!;

        provider.Fingerprint = "volume-b";
        tracker.Refresh();

        Assert.IsFalse(tracker.Matches(snapshot, temp.Path));
    }

    [TestMethod]
    public void SameVolumeFingerprintDifferentPhysicalDevice_InvalidatesOldSession()
    {
        using var temp = new TempDirectory();
        var provider = new FakeProvider(temp.Path, "volume-a", physicalDeviceFingerprint: "physical-a");
        var tracker = new StorageSessionTracker(provider);
        var snapshot = tracker.Capture(temp.Path)!;

        provider.PhysicalDeviceFingerprint = "physical-b";
        tracker.Refresh();

        Assert.IsFalse(tracker.Matches(snapshot, temp.Path));
        var replacement = tracker.Capture(temp.Path);
        Assert.IsNotNull(replacement);
        Assert.AreNotEqual(snapshot.SessionId, replacement.SessionId);
        Assert.AreEqual("physical-b", replacement.PhysicalDeviceFingerprint);
    }

    [TestMethod]
    public void MissingFingerprint_FailsClosed()
    {
        using var temp = new TempDirectory();
        var provider = new FakeProvider(temp.Path, string.Empty);
        var tracker = new StorageSessionTracker(provider);

        Assert.IsNull(tracker.Capture(temp.Path));
    }

    [TestMethod]
    public void NestedDcimSelection_ExpandsToMountedCardRoot()
    {
        using var temp = new TempDirectory();
        var nested = Path.Combine(temp.Path, "DCIM", "100NIKON");
        Directory.CreateDirectory(nested);
        var provider = new FakeProvider(temp.Path, "volume-a", isRemovable: true);
        var resolver = new CameraCardRootResolver(provider);

        Assert.AreEqual(PathSafety.Normalize(temp.Path), resolver.Resolve(nested));
    }

    [TestMethod]
    public void PrivateCameraStructure_IsAccepted()
    {
        using var temp = new TempDirectory();
        var nested = Path.Combine(temp.Path, "PRIVATE", "AVCHD");
        Directory.CreateDirectory(nested);
        var provider = new FakeProvider(temp.Path, "volume-a", isRemovable: true);
        var resolver = new CameraCardRootResolver(provider);

        Assert.AreEqual(PathSafety.Normalize(temp.Path), resolver.Resolve(nested));
    }

    [TestMethod]
    public void ArbitraryMountedFolderWithoutCameraStructure_IsRejected()
    {
        using var temp = new TempDirectory();
        var child = Path.Combine(temp.Path, "Pictures");
        Directory.CreateDirectory(child);
        var provider = new FakeProvider(temp.Path, "volume-a", isRemovable: true);
        var resolver = new CameraCardRootResolver(provider);

        Assert.IsNull(resolver.Resolve(child));
    }

    [TestMethod]
    public void SystemVolume_IsRejectedEvenWhenItContainsDcim()
    {
        using var temp = new TempDirectory();
        Directory.CreateDirectory(Path.Combine(temp.Path, "DCIM"));
        var provider = new FakeProvider(temp.Path, "system", isSystem: true);
        var resolver = new CameraCardRootResolver(provider);

        Assert.IsNull(resolver.Resolve(temp.Path));
    }

    private sealed class FakeProvider : IStorageVolumeProvider
    {
        private readonly string _root;
        private readonly bool _isRemovable;
        private readonly bool _isSystem;

        public FakeProvider(
            string root,
            string fingerprint,
            bool isRemovable = false,
            bool isSystem = false,
            string? physicalDeviceFingerprint = "physical-a")
        {
            _root = PathSafety.Normalize(root);
            Fingerprint = fingerprint;
            PhysicalDeviceFingerprint = physicalDeviceFingerprint;
            _isRemovable = isRemovable;
            _isSystem = isSystem;
        }

        public bool Mounted { get; set; } = true;
        public string Fingerprint { get; set; }
        public string? PhysicalDeviceFingerprint { get; set; }
        public int GetMountedVolumesCallCount { get; private set; }
        public int ResolveVolumeForPathCallCount { get; private set; }
        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes()
        {
            GetMountedVolumesCallCount++;
            if (!Mounted) return [];
            return [new MountedVolumeInfo(
                _root,
                Fingerprint,
                _isRemovable,
                _isSystem,
                PhysicalDeviceFingerprint)];
        }

        public MountedVolumeInfo? ResolveVolumeForPath(string path)
        {
            ResolveVolumeForPathCallCount++;
            if (!Mounted || string.IsNullOrWhiteSpace(Fingerprint)) return null;
            if (!PathSafety.IsSameOrDescendant(path, _root, PathComparison)) return null;
            return new MountedVolumeInfo(
                _root,
                Fingerprint,
                _isRemovable,
                _isSystem,
                PhysicalDeviceFingerprint);
        }
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PhotoOrganizerStorageTests-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
