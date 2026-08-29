using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class ExtendedRuntimeSafetyTests
{
    [TestMethod]
    public async Task SupportedFileAppearingAfterRescanBeforeVerification_BlocksReuse()
    {
        using var env = RuntimeEnvironment.Create();
        env.AddCameraFile("DCIM/original.jpg", "original-data", new DateTime(2026, 8, 30, 8, 0, 0));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        Assert.IsTrue(scan.IsReady);

        var added = false;
        var progress = new InlineProgress(update =>
        {
            if (!added && update.Phase == ImportProgressPhase.Verifying)
            {
                added = true;
                env.AddCameraFile("DCIM/late.jpg", "late-data", new DateTime(2026, 8, 30, 8, 1, 0));
            }
        });

        var result = await coordinator.ImportAsync(
            scan.Session!,
            env.DestinationRoot,
            "Race",
            progress);

        Assert.IsTrue(added, "The late media mutation was not injected.");
        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        Assert.IsFalse(result.IsSafeToReuse);
    }

    [TestMethod]
    public async Task DestinationParentOfCameraCard_IsRejectedBeforeCopy()
    {
        using var env = RuntimeEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 8, 30));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        Assert.IsTrue(scan.IsReady);

        var result = await coordinator.ImportAsync(scan.Session!, env.Root, "UnsafeParent");

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        StringAssert.Contains(result.Message, "parent/child");
        Assert.AreEqual("camera-data", File.ReadAllText(Path.Combine(env.CardRoot, "DCIM", "photo.jpg")));
    }

    [TestMethod]
    public void NestedDcimSelection_ScansWholeCardIncludingPrivateMedia()
    {
        using var env = RuntimeEnvironment.Create();
        var selected = Directory.CreateDirectory(Path.Combine(env.CardRoot, "DCIM", "100NIKON")).FullName;
        var jpg = env.AddCameraFile("DCIM/100NIKON/photo.jpg", "jpg-data", new DateTime(2026, 8, 30));
        var mov = env.AddCameraFile("PRIVATE/AVCHD/clip.mp4", "video-data", new DateTime(2026, 8, 30));

        var scan = env.CreateCoordinator().ScanCard(selected);

        Assert.IsTrue(scan.IsReady, scan.Message);
        Assert.AreEqual(PathSafety.Normalize(env.CardRoot), scan.Session!.CardRoot);
        CollectionAssert.AreEquivalent(
            new[] { Path.GetFullPath(jpg), Path.GetFullPath(mov) },
            scan.Session.Files.ToArray());
    }

    [TestMethod]
    public void AmbiguousMountedIdentityForSameRoot_FailsClosed()
    {
        using var temp = new TempDirectory();
        var root = Directory.CreateDirectory(Path.Combine(temp.Path, "volume")).FullName;
        var provider = new StaticVolumeProvider(
        [
            new MountedVolumeInfo(root, "fingerprint-a", true, false, "physical-a"),
            new MountedVolumeInfo(root, "fingerprint-b", true, false, "physical-b")
        ]);
        var tracker = new StorageSessionTracker(provider);

        Assert.IsNull(tracker.Capture(root));
    }

    private sealed class InlineProgress(Action<ImportProgress> callback) : IProgress<ImportProgress>
    {
        public void Report(ImportProgress value) => callback(value);
    }

    private sealed class RuntimeEnvironment : IDisposable
    {
        private RuntimeEnvironment(string root, string cardRoot, string destinationRoot, StaticVolumeProvider provider)
        {
            Root = root;
            CardRoot = cardRoot;
            DestinationRoot = destinationRoot;
            Provider = provider;
            Tracker = new StorageSessionTracker(provider);
        }

        public string Root { get; }
        public string CardRoot { get; }
        public string DestinationRoot { get; }
        public StaticVolumeProvider Provider { get; }
        public StorageSessionTracker Tracker { get; }

        public static RuntimeEnvironment Create()
        {
            var root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerExtendedRuntime-{Guid.NewGuid():N}");
            var card = Path.Combine(root, "card");
            var destination = Path.Combine(root, "destination");
            Directory.CreateDirectory(Path.Combine(card, "DCIM"));
            Directory.CreateDirectory(destination);

            var provider = new StaticVolumeProvider(
            [
                new MountedVolumeInfo(card, "camera-volume", true, false, "physical-camera"),
                new MountedVolumeInfo(destination, "destination-volume", false, false, "physical-destination")
            ]);
            return new RuntimeEnvironment(root, card, destination, provider);
        }

        public string AddCameraFile(string relativePath, string content, DateTime timestamp)
        {
            var path = Path.Combine(CardRoot, relativePath.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, content);
            File.SetLastWriteTime(path, timestamp);
            return path;
        }

        public ImportCoordinator CreateCoordinator() => new(
            new MediaClassifier(),
            Tracker,
            new CameraCardRootResolver(Provider),
            Provider);

        public void Dispose()
        {
            try { Directory.Delete(Root, recursive: true); } catch { }
        }
    }

    private sealed class StaticVolumeProvider(IReadOnlyList<MountedVolumeInfo> volumes) : IStorageVolumeProvider
    {
        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes() => volumes;

        public MountedVolumeInfo? ResolveVolumeForPath(string path)
        {
            var normalized = PathSafety.Normalize(path);
            return volumes
                .Where(volume => PathSafety.IsSameOrDescendant(normalized, volume.RootPath, PathComparison))
                .OrderByDescending(volume => PathSafety.Normalize(volume.RootPath).Length)
                .FirstOrDefault();
        }
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PhotoOrganizerAmbiguousIdentity-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
