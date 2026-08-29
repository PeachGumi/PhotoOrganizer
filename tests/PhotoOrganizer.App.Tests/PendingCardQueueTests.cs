using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class PendingCardQueueTests
{
    [TestMethod]
    public async Task EmptyAutoDetectedScan_ContinuesToCardQueuedWhileScanning()
    {
        using var temp = new TempDirectory();
        var emptyCard = Directory.CreateDirectory(Path.Combine(temp.Path, "01-empty-card")).FullName;
        var validCard = Directory.CreateDirectory(Path.Combine(temp.Path, "02-valid-card")).FullName;
        Directory.CreateDirectory(Path.Combine(emptyCard, "DCIM"));
        var validDcim = Directory.CreateDirectory(Path.Combine(validCard, "DCIM")).FullName;
        File.WriteAllText(Path.Combine(validDcim, "photo.jpg"), "camera-data");

        using var provider = new BlockingVolumeProvider(
        [
            new MountedVolumeInfo(emptyCard, "volume-empty", true, false, "device-empty"),
            new MountedVolumeInfo(validCard, "volume-valid", true, false, "device-valid")
        ]);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        using var viewModel = new MainWindowViewModel(provider, sessions, roots);

        var firstScan = viewModel.ScanCardAsync(emptyCard, autoDetected: true);
        Assert.IsTrue(
            provider.FirstEnumerationEntered.Wait(TimeSpan.FromSeconds(5)),
            "The first scan did not enter volume enumeration in time.");
        Assert.IsTrue(viewModel.IsBusy);

        var queuedResult = await viewModel.ScanCardAsync(validCard, autoDetected: true);
        Assert.IsNull(queuedResult);
        Assert.AreEqual(1, viewModel.PendingSdCount);

        provider.ReleaseFirstEnumeration.Set();
        var firstResult = await firstScan.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.IsNotNull(firstResult);
        Assert.IsTrue(firstResult.IsNoSupportedMedia);
        Assert.AreEqual(PathSafety.Normalize(validCard), viewModel.SelectedSdPath);
        Assert.AreEqual("RAW:0 / JPG:1 / MP4:0", viewModel.CountLabel);
        Assert.AreEqual("取り込み準備完了", viewModel.ProgressLabel);
        Assert.AreEqual(0, viewModel.PendingSdCount);
    }

    [TestMethod]
    public async Task SafetyFailure_DoesNotAutoAdvanceToQueuedCard()
    {
        using var temp = new TempDirectory();
        var brokenCard = Directory.CreateDirectory(Path.Combine(temp.Path, "01-broken-card")).FullName;
        var validCard = Directory.CreateDirectory(Path.Combine(temp.Path, "02-valid-card")).FullName;
        var brokenDcim = Directory.CreateDirectory(Path.Combine(brokenCard, "DCIM")).FullName;
        var validDcim = Directory.CreateDirectory(Path.Combine(validCard, "DCIM")).FullName;
        File.WriteAllText(Path.Combine(brokenDcim, "broken.jpg"), "camera-data");
        File.WriteAllText(Path.Combine(validDcim, "photo.jpg"), "camera-data");

        using var provider = new BlockingVolumeProvider(
        [
            new MountedVolumeInfo(brokenCard, "volume-broken", true, false, null),
            new MountedVolumeInfo(validCard, "volume-valid", true, false, "device-valid")
        ]);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        using var viewModel = new MainWindowViewModel(provider, sessions, roots);

        var firstScan = viewModel.ScanCardAsync(brokenCard, autoDetected: true);
        Assert.IsTrue(
            provider.FirstEnumerationEntered.Wait(TimeSpan.FromSeconds(5)),
            "The first scan did not enter volume enumeration in time.");

        Assert.IsNull(await viewModel.ScanCardAsync(validCard, autoDetected: true));
        Assert.AreEqual(1, viewModel.PendingSdCount);

        provider.ReleaseFirstEnumeration.Set();
        var firstResult = await firstScan.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.IsNotNull(firstResult);
        Assert.AreEqual(ScanFailureReason.MissingPhysicalDeviceIdentity, firstResult.FailureReason);
        Assert.AreEqual(string.Empty, viewModel.SelectedSdPath);
        Assert.AreEqual("スキャン失敗", viewModel.ProgressLabel);
        Assert.AreEqual(1, viewModel.PendingSdCount);
        StringAssert.Contains(viewModel.SafetyDetail, "physical-device identity");
    }

    [TestMethod]
    public async Task QueueDrain_StopsOnSafetyFailureAndLeavesLaterCardPending()
    {
        using var temp = new TempDirectory();
        var emptyCard = Directory.CreateDirectory(Path.Combine(temp.Path, "01-empty-card")).FullName;
        var brokenCard = Directory.CreateDirectory(Path.Combine(temp.Path, "02-broken-card")).FullName;
        var validCard = Directory.CreateDirectory(Path.Combine(temp.Path, "03-valid-card")).FullName;
        Directory.CreateDirectory(Path.Combine(emptyCard, "DCIM"));
        var brokenDcim = Directory.CreateDirectory(Path.Combine(brokenCard, "DCIM")).FullName;
        var validDcim = Directory.CreateDirectory(Path.Combine(validCard, "DCIM")).FullName;
        File.WriteAllText(Path.Combine(brokenDcim, "broken.jpg"), "camera-data");
        File.WriteAllText(Path.Combine(validDcim, "photo.jpg"), "camera-data");

        using var provider = new BlockingVolumeProvider(
        [
            new MountedVolumeInfo(emptyCard, "volume-empty", true, false, "device-empty"),
            new MountedVolumeInfo(brokenCard, "volume-broken", true, false, null),
            new MountedVolumeInfo(validCard, "volume-valid", true, false, "device-valid")
        ]);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        using var viewModel = new MainWindowViewModel(provider, sessions, roots);

        var firstScan = viewModel.ScanCardAsync(emptyCard, autoDetected: true);
        Assert.IsTrue(
            provider.FirstEnumerationEntered.Wait(TimeSpan.FromSeconds(5)),
            "The first scan did not enter volume enumeration in time.");

        Assert.IsNull(await viewModel.ScanCardAsync(brokenCard, autoDetected: true));
        Assert.IsNull(await viewModel.ScanCardAsync(validCard, autoDetected: true));
        Assert.AreEqual(2, viewModel.PendingSdCount);

        provider.ReleaseFirstEnumeration.Set();
        var firstResult = await firstScan.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.IsNotNull(firstResult);
        Assert.IsTrue(firstResult.IsNoSupportedMedia);
        Assert.AreEqual(string.Empty, viewModel.SelectedSdPath);
        Assert.AreEqual("スキャン失敗", viewModel.ProgressLabel);
        Assert.AreEqual(1, viewModel.PendingSdCount);
        StringAssert.Contains(viewModel.SafetyDetail, "physical-device identity");
    }

    private sealed class BlockingVolumeProvider(IReadOnlyList<MountedVolumeInfo> volumes)
        : IStorageVolumeProvider, IDisposable
    {
        private int _blockFirstEnumeration = 1;

        public ManualResetEventSlim FirstEnumerationEntered { get; } = new(false);
        public ManualResetEventSlim ReleaseFirstEnumeration { get; } = new(false);

        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes()
        {
            if (Interlocked.Exchange(ref _blockFirstEnumeration, 0) == 1)
            {
                FirstEnumerationEntered.Set();
                if (!ReleaseFirstEnumeration.Wait(TimeSpan.FromSeconds(10)))
                {
                    throw new TimeoutException("Test did not release the first volume enumeration.");
                }
            }

            return volumes;
        }

        public MountedVolumeInfo? ResolveVolumeForPath(string path)
        {
            var normalized = PathSafety.Normalize(path);
            return volumes
                .Where(volume => PathSafety.IsSameOrDescendant(normalized, volume.RootPath, PathComparison))
                .OrderByDescending(volume => PathSafety.Normalize(volume.RootPath).Length)
                .FirstOrDefault();
        }

        public void Dispose()
        {
            FirstEnumerationEntered.Dispose();
            ReleaseFirstEnumeration.Dispose();
        }
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PhotoOrganizerPendingCards-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
