using PhotoOrganizer.Core;

namespace PhotoOrganizer.Core.Tests;

public sealed class StorageIdentityTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "photo-organizer-storage-tests-" + Guid.NewGuid().ToString("N"));

    public StorageIdentityTests()
    {
        Directory.CreateDirectory(_root);
    }

    [Fact]
    public void Capture_MatchesWithinSameMountSession()
    {
        var provider = new FakeProvider(_root, "volume-a");
        var tracker = new StorageSessionTracker(provider);

        var snapshot = tracker.Capture(_root);

        Assert.NotNull(snapshot);
        Assert.True(tracker.Matches(snapshot!, _root));
    }

    [Fact]
    public void RemovalAndReinsert_InvalidatesOldSessionEvenForSameFingerprint()
    {
        var provider = new FakeProvider(_root, "volume-a");
        var tracker = new StorageSessionTracker(provider);
        var snapshot = tracker.Capture(_root)!;

        provider.Mounted = false;
        tracker.Refresh();
        provider.Mounted = true;
        tracker.Refresh();

        Assert.False(tracker.Matches(snapshot, _root));
        var replacement = tracker.Capture(_root);
        Assert.NotNull(replacement);
        Assert.NotEqual(snapshot.SessionId, replacement!.SessionId);
    }

    [Fact]
    public void SamePathDifferentVolume_InvalidatesOldSession()
    {
        var provider = new FakeProvider(_root, "volume-a");
        var tracker = new StorageSessionTracker(provider);
        var snapshot = tracker.Capture(_root)!;

        provider.Fingerprint = "volume-b";
        tracker.Refresh();

        Assert.False(tracker.Matches(snapshot, _root));
    }

    [Fact]
    public void MissingFingerprint_FailsClosed()
    {
        var provider = new FakeProvider(_root, string.Empty);
        var tracker = new StorageSessionTracker(provider);

        Assert.Null(tracker.Capture(_root));
    }

    [Fact]
    public void NestedDcimSelection_ExpandsToMountedCardRoot()
    {
        var nested = Path.Combine(_root, "DCIM", "100NIKON");
        Directory.CreateDirectory(nested);
        var provider = new FakeProvider(_root, "volume-a", isRemovable: true);
        var resolver = new CameraCardRootResolver(provider);

        Assert.Equal(PathSafety.Normalize(_root), resolver.Resolve(nested));
    }

    [Fact]
    public void PrivateCameraStructure_IsAccepted()
    {
        var nested = Path.Combine(_root, "PRIVATE", "AVCHD");
        Directory.CreateDirectory(nested);
        var provider = new FakeProvider(_root, "volume-a", isRemovable: true);
        var resolver = new CameraCardRootResolver(provider);

        Assert.Equal(PathSafety.Normalize(_root), resolver.Resolve(nested));
    }

    [Fact]
    public void ArbitraryMountedFolderWithoutCameraStructure_IsRejected()
    {
        var child = Path.Combine(_root, "Pictures");
        Directory.CreateDirectory(child);
        var provider = new FakeProvider(_root, "volume-a", isRemovable: true);
        var resolver = new CameraCardRootResolver(provider);

        Assert.Null(resolver.Resolve(child));
    }

    [Fact]
    public void SystemVolume_IsRejectedEvenWhenItContainsDcim()
    {
        Directory.CreateDirectory(Path.Combine(_root, "DCIM"));
        var provider = new FakeProvider(_root, "system", isSystem: true);
        var resolver = new CameraCardRootResolver(provider);

        Assert.Null(resolver.Resolve(_root));
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
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
}
