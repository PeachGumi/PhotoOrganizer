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

        public FakeProvider(string root, string fingerprint, bool isRemovable = false, bool isSystem = false)
        {
            _root = PathSafety.Normalize(root);
            Fingerprint = fingerprint;
            _isRemovable = isRemovable;
            _isSystem = isSystem;
        }

        public bool Mounted { get; set; } = true;
        public string Fingerprint { get; set; }
        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes()
        {
            if (!Mounted) return [];
            return [new MountedVolumeInfo(_root, Fingerprint, _isRemovable, _isSystem)];
        }

        public MountedVolumeInfo? ResolveVolumeForPath(string path)
        {
            if (!Mounted || string.IsNullOrWhiteSpace(Fingerprint)) return null;
            if (!PathSafety.IsSameOrDescendant(path, _root, PathComparison)) return null;
            return new MountedVolumeInfo(_root, Fingerprint, _isRemovable, _isSystem);
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
