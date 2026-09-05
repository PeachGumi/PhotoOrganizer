using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class ImportCoordinatorTests
{
    [TestMethod]
    public async Task EventDirectorySymlink_IsRejectedWithoutCreatingDirectoriesOnCard()
    {
        using var env = TestEnvironment.Create();
        var source = env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 1, 2));
        var year = Directory.CreateDirectory(Path.Combine(env.DestinationRoot, "2026")).FullName;
        var link = Path.Combine(year, "2026-01-02_Event");
        try
        {
            Directory.CreateSymbolicLink(link, env.CardRoot);
        }
        catch (Exception ex) when (ex is PlatformNotSupportedException or UnauthorizedAccessException or IOException)
        {
            Assert.Inconclusive($"Symbolic links unavailable: {ex.Message}");
        }
        try
        {
            var coordinator = env.CreateCoordinator();
            var result = await coordinator.ImportAsync(coordinator.ScanCard(env.CardRoot).Session!, env.DestinationRoot, "Event");
            Assert.IsFalse(result.IsSafeToReuse);
            Assert.AreEqual("camera-data", File.ReadAllText(source));
            foreach (var kind in Enum.GetValues<MediaKind>())
            {
                Assert.IsFalse(Directory.Exists(Path.Combine(env.CardRoot, MediaClassifier.FolderName(kind))),
                    "Reject aliased event paths before any directory creation.");
            }
        }
        finally { Directory.Delete(link); }
    }

    [TestMethod]
    public async Task SuccessfulImport_IsSafeOnlyAfterFreshRescanAndByteVerification()
    {
        using var env = TestEnvironment.Create();
        var source = env.AddCameraFile("DCIM/100NIKON/DSC_0001.jpg", "camera-bytes", new DateTime(2026, 1, 2, 10, 0, 0));
        var sourceBefore = File.ReadAllBytes(source);
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(Path.GetDirectoryName(source)!);

        Assert.IsTrue(scan.IsReady);
        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Tokyo");

        Assert.IsTrue(result.IsSafeToReuse, result.Message);
        Assert.AreEqual(1, result.Summary.Copied);
        var copied = Path.Combine(env.DestinationRoot, "2026", "2026-01-02_Tokyo", "JPG", "DSC_0001.jpg");
        Assert.IsTrue(File.Exists(copied));
        CollectionAssert.AreEqual(sourceBefore, File.ReadAllBytes(source));
        CollectionAssert.AreEqual(sourceBefore, File.ReadAllBytes(copied));
    }

    [TestMethod]
    public async Task ExistingDifferentFile_IsNeverOverwrittenAndUsesCollisionSuffix()
    {
        using var env = TestEnvironment.Create();
        var source = env.AddCameraFile("DCIM/DSC_0001.jpg", "NEW1", new DateTime(2026, 2, 3));
        var expectedFolder = Path.Combine(env.DestinationRoot, "2026", "2026-02-03_Event", "JPG");
        Directory.CreateDirectory(expectedFolder);
        var existing = Path.Combine(expectedFolder, "DSC_0001.jpg");
        File.WriteAllText(existing, "OLD1");
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);

        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Event");

        Assert.IsTrue(result.IsSafeToReuse, result.Message);
        Assert.AreEqual("OLD1", File.ReadAllText(existing));
        Assert.AreEqual("NEW1", File.ReadAllText(Path.Combine(expectedFolder, "DSC_0001_2.jpg")));
        Assert.AreEqual("NEW1", File.ReadAllText(source));
    }

    [TestMethod]
    public async Task ConcurrentSameNameCopies_PreserveBothSourcesAndFinalizeWithoutPartials()
    {
        using var env = TestEnvironment.Create();
        var first = env.AddCameraFile("DCIM/CARD_A/photo.jpg", "first-camera-payload", new DateTime(2026, 2, 4));
        var second = env.AddCameraFile("DCIM/CARD_B/photo.jpg", "second-camera-payload", new DateTime(2026, 2, 4));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);

        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Parallel");

        Assert.IsTrue(
            result.IsSafeToReuse,
            $"{result.Message} Errors: {string.Join(" | ", result.Summary.Errors)}");
        Assert.AreEqual(2, result.Summary.Copied);
        Assert.AreEqual("first-camera-payload", File.ReadAllText(first));
        Assert.AreEqual("second-camera-payload", File.ReadAllText(second));

        var destinationFolder = Path.Combine(
            env.DestinationRoot,
            "2026",
            "2026-02-04_Parallel",
            "JPG");
        var copiedPayloads = Directory
            .EnumerateFiles(destinationFolder, "photo*.jpg", SearchOption.TopDirectoryOnly)
            .Select(File.ReadAllText)
            .ToArray();
        CollectionAssert.AreEquivalent(
            new[] { "first-camera-payload", "second-camera-payload" },
            copiedPayloads);
        Assert.IsFalse(Directory
            .EnumerateFiles(destinationFolder, ".partial-*", SearchOption.TopDirectoryOnly)
            .Any());
    }

    [TestMethod]
    public async Task DuplicateOnlySecondRun_DoesNotCreateEmptyEventFolder()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "same-data", new DateTime(2026, 3, 4));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        var first = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "First");
        Assert.IsTrue(first.IsSafeToReuse);

        var secondScan = coordinator.ScanCard(env.CardRoot);
        var second = await coordinator.ImportAsync(secondScan.Session!, env.DestinationRoot, "Second");

        Assert.IsTrue(second.IsSafeToReuse, second.Message);
        Assert.AreEqual(0, second.Summary.Copied);
        Assert.AreEqual(1, second.Summary.SkippedAlreadyBackedUp);
        Assert.IsTrue(Directory.Exists(second.Summary.BasePath), "Completion must point to an existing directory.");
        Assert.AreEqual(env.DestinationRoot, second.Summary.BasePath);
        Assert.IsFalse(Directory.Exists(Path.Combine(env.DestinationRoot, "2026", "2026-03-04_Second")));
    }

    [TestMethod]
    public async Task DestinationInsideCard_IsRejectedBeforeCopy()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 4, 5));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        var invalidDestination = Path.Combine(env.CardRoot, "Backup");

        var result = await coordinator.ImportAsync(scan.Session!, invalidDestination, "Event");

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        Assert.IsFalse(Directory.Exists(invalidDestination));
    }

    [TestMethod]
    public async Task DifferentVolumeOnSamePhysicalDevice_IsRejectedBeforeCopy()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 4, 6));
        env.Provider.SetPhysicalDeviceFingerprint(env.DestinationRoot, "physical-camera-card");
        env.Tracker.Refresh();
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);

        Assert.IsTrue(scan.IsReady);
        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Event");

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        StringAssert.Contains(result.Message, "physical storage device");
        Assert.IsFalse(Directory.Exists(Path.Combine(env.DestinationRoot, "2026")));
    }

    [TestMethod]
    public void MissingCameraPhysicalDeviceIdentity_BlocksScan()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 4, 7));
        env.Provider.SetPhysicalDeviceFingerprint(env.CardRoot, null);
        env.Tracker.Refresh();

        var scan = env.CreateCoordinator().ScanCard(env.CardRoot);

        Assert.AreEqual(ImportSafetyStatus.Blocked, scan.Status);
        Assert.IsFalse(scan.IsReady);
        StringAssert.Contains(scan.Message, "physical-device identity");
    }

    [TestMethod]
    public async Task MissingDestinationPhysicalDeviceIdentity_BlocksBeforeCopy()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 4, 8));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        env.Provider.SetPhysicalDeviceFingerprint(env.DestinationRoot, null);
        env.Tracker.Refresh();

        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Event");

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        StringAssert.Contains(result.Message, "Destination physical-device identity is unavailable");
        Assert.IsFalse(Directory.Exists(Path.Combine(env.DestinationRoot, "2026")));
    }

    [TestMethod]
    public async Task RemovalAndReinsertAfterScan_InvalidatesSession()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 5, 6));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        Assert.IsTrue(scan.IsReady);

        env.Provider.SetMounted(env.CardRoot, false);
        env.Tracker.Refresh();
        env.Provider.SetMounted(env.CardRoot, true);
        env.Tracker.Refresh();

        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Event");

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        StringAssert.Contains(result.Message, "changed");
    }

    [TestMethod]
    public async Task SamePathReplacementAfterScan_IsBlocked()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 6, 7));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);

        env.Provider.SetFingerprint(env.CardRoot, "replacement-card");
        env.Tracker.Refresh();
        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Event");

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        Assert.IsFalse(result.IsSafeToReuse);
    }

    [TestMethod]
    public async Task NewSupportedFileAppearingAfterCopy_BlocksReuseEvenThoughCopyCompleted()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "original-data", new DateTime(2026, 7, 8));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        var sawCopyComplete = false;
        var progress = new InlineProgress(update =>
        {
            if (update.Phase == ImportProgressPhase.Copying && update.Message.Contains("complete", StringComparison.OrdinalIgnoreCase))
            {
                sawCopyComplete = true;
            }

            if (update.Phase == ImportProgressPhase.Rescanning)
            {
                env.AddCameraFile("DCIM/new-photo.jpg", "new-after-copy", new DateTime(2026, 7, 8));
            }
        });

        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Event", progress);

        Assert.IsTrue(sawCopyComplete);
        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        Assert.IsFalse(result.IsSafeToReuse);
        Assert.IsNotNull(result.Verification);
        Assert.AreEqual(2, result.Verification.Total);
        Assert.AreEqual(1, result.Verification.Verified);
    }

    [TestMethod]
    public async Task InitiallyObservedFileMissingAtRescan_BlocksReuse()
    {
        using var env = TestEnvironment.Create();
        var source = env.AddCameraFile("DCIM/photo.jpg", "original-data", new DateTime(2026, 8, 9));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        var progress = new InlineProgress(update =>
        {
            if (update.Phase == ImportProgressPhase.Rescanning && File.Exists(source))
            {
                File.Delete(source);
            }
        });

        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Event", progress);

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        Assert.IsFalse(result.IsSafeToReuse);
    }

    [TestMethod]
    public async Task DestinationReplacementDuringCopy_BlocksFinalApproval()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "original-data", new DateTime(2026, 9, 10));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        var changed = false;
        var progress = new InlineProgress(update =>
        {
            if (!changed && update.Phase == ImportProgressPhase.Copying)
            {
                changed = true;
                env.Provider.SetFingerprint(env.DestinationRoot, "replacement-destination");
                env.Tracker.Refresh();
            }
        });

        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Event", progress);

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        Assert.IsFalse(result.IsSafeToReuse);
    }

    [TestMethod]
    public async Task CancellationBeforePostImportRescan_BlocksReusePromptly()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 9, 11));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        using var cancellation = new CancellationTokenSource();
        var progress = new InlineProgress(update =>
        {
            if (update.Phase == ImportProgressPhase.Rescanning)
            {
                cancellation.Cancel();
            }
        });

        var result = await coordinator.ImportAsync(
            scan.Session!,
            env.DestinationRoot,
            "Event",
            progress,
            cancellation.Token);

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        Assert.IsFalse(result.IsSafeToReuse);
        StringAssert.Contains(result.Message, "cancelled");
    }

    [TestMethod]
    public async Task SessionContainingFileOutsideCard_IsRejectedBeforeCopy()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 9, 12));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        var outside = Path.Combine(env.DestinationRoot, "outside.jpg");
        File.WriteAllText(outside, "unrelated-data");
        var tamperedSession = scan.Session! with { Files = [outside] };

        var result = await coordinator.ImportAsync(
            tamperedSession,
            env.DestinationRoot,
            "Event");

        Assert.AreEqual(ImportSafetyStatus.Blocked, result.Status);
        Assert.IsFalse(result.IsSafeToReuse);
        StringAssert.Contains(result.Message, "outside");
        Assert.IsFalse(Directory.Exists(Path.Combine(env.DestinationRoot, "2026")));
    }

    [TestMethod]
    public async Task UnsupportedSidecars_DoNotBlockSupportedMediaReuseApproval()
    {
        using var env = TestEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-data", new DateTime(2026, 10, 11));
        env.AddCameraFile("DCIM/photo.xmp", "sidecar-data", new DateTime(2026, 10, 11));
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);

        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Event");

        Assert.IsTrue(result.IsSafeToReuse, result.Message);
        Assert.AreEqual(1, result.Summary.TotalSupported);
    }

    private sealed class InlineProgress(Action<ImportProgress> callback) : IProgress<ImportProgress>
    {
        public void Report(ImportProgress value) => callback(value);
    }

    private sealed class TestEnvironment : IDisposable
    {
        private readonly string _root;

        private TestEnvironment(string root, string cardRoot, string destinationRoot, FakeVolumeProvider provider)
        {
            _root = root;
            CardRoot = cardRoot;
            DestinationRoot = destinationRoot;
            Provider = provider;
            Tracker = new StorageSessionTracker(provider);
        }

        public string CardRoot { get; }
        public string DestinationRoot { get; }
        public FakeVolumeProvider Provider { get; }
        public StorageSessionTracker Tracker { get; }

        public static TestEnvironment Create()
        {
            var root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerImportTests-{Guid.NewGuid():N}");
            var card = Path.Combine(root, "card");
            var destination = Path.Combine(root, "destination");
            Directory.CreateDirectory(Path.Combine(card, "DCIM"));
            Directory.CreateDirectory(destination);
            var provider = new FakeVolumeProvider([
                new MutableVolume(card, "camera-card", true, false, true, "physical-camera-card"),
                new MutableVolume(destination, "destination-volume", false, false, true, "physical-destination")
            ]);
            return new TestEnvironment(root, card, destination, provider);
        }

        public string AddCameraFile(string relativePath, string content, DateTime timestamp)
        {
            var path = Path.Combine(CardRoot, relativePath.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, content);
            File.SetLastWriteTime(path, timestamp);
            return path;
        }

        public ImportCoordinator CreateCoordinator()
        {
            var classifier = new MediaClassifier();
            return new ImportCoordinator(
                classifier,
                Tracker,
                new CameraCardRootResolver(Provider),
                Provider);
        }

        public void Dispose()
        {
            try { Directory.Delete(_root, recursive: true); } catch { }
        }
    }

    private sealed record MutableVolume(
        string RootPath,
        string Fingerprint,
        bool IsRemovable,
        bool IsSystem,
        bool Mounted,
        string? PhysicalDeviceFingerprint);

    private sealed class FakeVolumeProvider : IStorageVolumeProvider
    {
        private readonly List<MutableVolume> _volumes;

        public FakeVolumeProvider(IEnumerable<MutableVolume> volumes)
        {
            _volumes = volumes
                .Select(v => v with { RootPath = PathSafety.Normalize(v.RootPath) })
                .ToList();
        }

        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes() => _volumes
            .Where(v => v.Mounted)
            .Select(v => new MountedVolumeInfo(
                v.RootPath,
                v.Fingerprint,
                v.IsRemovable,
                v.IsSystem,
                v.PhysicalDeviceFingerprint))
            .ToArray();

        public MountedVolumeInfo? ResolveVolumeForPath(string path)
        {
            var current = PathSafety.Normalize(path);
            while (!File.Exists(current) && !Directory.Exists(current))
            {
                var parent = Directory.GetParent(current)?.FullName;
                if (string.IsNullOrWhiteSpace(parent) || string.Equals(parent, current, PathComparison)) break;
                current = parent;
            }

            var volume = _volumes
                .Where(v => v.Mounted)
                .Where(v => PathSafety.IsSameOrDescendant(current, v.RootPath, PathComparison))
                .OrderByDescending(v => v.RootPath.Length)
                .FirstOrDefault();

            return volume is null
                ? null
                : new MountedVolumeInfo(
                    volume.RootPath,
                    volume.Fingerprint,
                    volume.IsRemovable,
                    volume.IsSystem,
                    volume.PhysicalDeviceFingerprint);
        }

        public void SetMounted(string root, bool mounted)
        {
            var normalized = PathSafety.Normalize(root);
            var index = _volumes.FindIndex(v => string.Equals(v.RootPath, normalized, PathComparison));
            if (index >= 0) _volumes[index] = _volumes[index] with { Mounted = mounted };
        }

        public void SetFingerprint(string root, string fingerprint)
        {
            var normalized = PathSafety.Normalize(root);
            var index = _volumes.FindIndex(v => string.Equals(v.RootPath, normalized, PathComparison));
            if (index >= 0) _volumes[index] = _volumes[index] with { Fingerprint = fingerprint };
        }

        public void SetPhysicalDeviceFingerprint(string root, string? fingerprint)
        {
            var normalized = PathSafety.Normalize(root);
            var index = _volumes.FindIndex(v => string.Equals(v.RootPath, normalized, PathComparison));
            if (index >= 0) _volumes[index] = _volumes[index] with { PhysicalDeviceFingerprint = fingerprint };
        }
    }
}
