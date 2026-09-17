using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

/// <summary>
/// A destination that cannot flush to its physical media (SMB/NFS) must still produce a
/// verified import, but the weaker durability level has to travel with the result so the
/// UI can say what was actually proven.
/// </summary>
[TestClass]
public sealed class DurabilityLevelSafetyTests
{
    [TestMethod]
    public async Task Import_FlushOnlyDestination_ApprovesReuseWithDocumentedCaveat()
    {
        using var env = CopyEnvironment.Create();
        var source = env.AddCameraFile("DCIM/photo.jpg", "camera-bytes", new DateTime(2026, 9, 17, 10, 0, 0));
        env.SeedDestinationEntry("keep-1.jpg");
        env.SeedDestinationEntry("keep-2.jpg");
        env.SeedDestinationEntry("keep-3.jpg");
        var coordinator = env.CreateCoordinator(new FlushOnlyDurabilityService());
        var scan = coordinator.ScanCard(env.CardRoot);
        Assert.IsTrue(scan.IsReady, scan.Message);

        var progress = new CollectedProgress();
        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Zoo", progress);

        Assert.AreEqual(ImportSafetyStatus.SafeToReuse, result.Status, result.Message);
        Assert.AreEqual(1, result.Summary.Copied);
        Assert.AreEqual(0, result.Summary.Failed);
        Assert.IsNotNull(result.Verification);
        Assert.IsTrue(
            result.Verification!.FullSyncUnsupported,
            "A flush-only destination must be reported as lacking a full media sync.");
        Assert.IsTrue(
            result.Summary.Warnings.Any(warning => warning.Contains("F_FULLFSYNC", StringComparison.Ordinal)),
            string.Join(" | ", result.Summary.Warnings));
        Assert.AreEqual("camera-bytes", File.ReadAllText(source));

        // The pre-copy destination index must announce how far it has scanned instead of
        // leaving the import silent for minutes.
        Assert.IsTrue(
            progress.MaxCheckingDestinationEntries >= 3,
            $"Expected the destination scan to report its entry count, got {progress.MaxCheckingDestinationEntries}.");
    }

    [TestMethod]
    public async Task Import_FullSyncDestination_ApprovesReuseWithoutCaveat()
    {
        using var env = CopyEnvironment.Create();
        env.AddCameraFile("DCIM/photo.jpg", "camera-bytes", new DateTime(2026, 9, 17, 10, 0, 0));
        var coordinator = env.CreateCoordinator(new FullSyncDurabilityService());
        var scan = coordinator.ScanCard(env.CardRoot);
        Assert.IsTrue(scan.IsReady, scan.Message);

        var result = await coordinator.ImportAsync(scan.Session!, env.DestinationRoot, "Zoo");

        Assert.AreEqual(ImportSafetyStatus.SafeToReuse, result.Status, result.Message);
        Assert.IsFalse(result.Verification!.FullSyncUnsupported);
        Assert.IsFalse(
            result.Summary.Warnings.Any(warning => warning.Contains("F_FULLFSYNC", StringComparison.Ordinal)),
            string.Join(" | ", result.Summary.Warnings));
    }

    private sealed class CollectedProgress : IProgress<ImportProgress>
    {
        public List<ImportProgressPhase> Phases { get; } = [];
        public int MaxCheckingDestinationEntries { get; private set; }

        public void Report(ImportProgress value)
        {
            Phases.Add(value.Phase);
            if (value.Phase == ImportProgressPhase.CheckingDestination)
            {
                MaxCheckingDestinationEntries = Math.Max(MaxCheckingDestinationEntries, value.Current);
            }
        }
    }

    private abstract class RealMoveDurabilityService : IFileDurabilityService
    {
        protected abstract DurabilityLevel Level { get; }

        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc)
        {
            File.Move(temporaryPath, finalPath, overwrite: false);
            File.SetLastWriteTimeUtc(finalPath, lastWriteUtc);
            return new FinalizeFileResult(FinalizeFileStatus.Committed, FinalPathCreated: true);
        }

        public DurabilityResult EnsureDurable(string filePath) => new(true, null, Level);
    }

    private sealed class FlushOnlyDurabilityService : RealMoveDurabilityService
    {
        protected override DurabilityLevel Level => DurabilityLevel.FlushOnly;
    }

    private sealed class FullSyncDurabilityService : RealMoveDurabilityService
    {
        protected override DurabilityLevel Level => DurabilityLevel.FullSync;
    }

    private sealed class CopyEnvironment : IDisposable
    {
        private CopyEnvironment(string root, string cardRoot, string destinationRoot, StaticVolumeProvider provider)
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

        public static CopyEnvironment Create()
        {
            var root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerDurabilityLevel-{Guid.NewGuid():N}");
            var card = Path.Combine(root, "card");
            var destination = Path.Combine(root, "destination");
            Directory.CreateDirectory(Path.Combine(card, "DCIM"));
            Directory.CreateDirectory(destination);

            var provider = new StaticVolumeProvider(
            [
                new MountedVolumeInfo(card, "camera-volume", true, false, "physical-camera"),
                new MountedVolumeInfo(destination, "destination-volume", false, false, "physical-destination")
            ]);
            return new CopyEnvironment(root, card, destination, provider);
        }

        public string AddCameraFile(string relativePath, string content, DateTime timestamp)
        {
            var path = Path.Combine(CardRoot, relativePath.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, content);
            File.SetLastWriteTime(path, timestamp);
            return path;
        }

        public void SeedDestinationEntry(string fileName) =>
            File.WriteAllText(Path.Combine(DestinationRoot, fileName), $"existing-{fileName}");

        public ImportCoordinator CreateCoordinator(IFileDurabilityService durability) => new(
            new MediaClassifier(),
            Tracker,
            new CameraCardRootResolver(Provider),
            Provider,
            durability);

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
}
