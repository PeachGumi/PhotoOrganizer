using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class StartupCardDiscoveryTests
{
    [TestMethod]
    public async Task InitializeAsync_SkipsEmptyFirstCardAndScansNextCandidate()
    {
        using var temp = new TempDirectory();
        var emptyCard = Directory.CreateDirectory(Path.Combine(temp.Path, "01-empty-card")).FullName;
        var validCard = Directory.CreateDirectory(Path.Combine(temp.Path, "02-valid-card")).FullName;
        Directory.CreateDirectory(Path.Combine(emptyCard, "DCIM"));
        var validDcim = Directory.CreateDirectory(Path.Combine(validCard, "DCIM")).FullName;
        File.WriteAllText(Path.Combine(validDcim, "photo.jpg"), "camera-data");

        var provider = new FakeVolumeProvider(
        [
            new MountedVolumeInfo(emptyCard, "volume-empty", true, false, "device-empty"),
            new MountedVolumeInfo(validCard, "volume-valid", true, false, "device-valid")
        ]);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        using var viewModel = new MainWindowViewModel(provider, sessions, roots);

        await StartupCardDiscovery.InitializeAsync(viewModel, roots);

        Assert.AreEqual(PathSafety.Normalize(validCard), viewModel.SelectedSdPath);
        Assert.AreEqual("RAW:0 / JPG:1 / MP4:0", viewModel.CountLabel);
        Assert.AreEqual("取り込み準備完了", viewModel.ProgressLabel);
        Assert.AreEqual(0, viewModel.PendingSdCount);
    }

    [TestMethod]
    public async Task InitializeAsync_StopsOnRealScanFailureInsteadOfSkippingToAnotherCard()
    {
        using var temp = new TempDirectory();
        var brokenCard = Directory.CreateDirectory(Path.Combine(temp.Path, "01-broken-card")).FullName;
        var validCard = Directory.CreateDirectory(Path.Combine(temp.Path, "02-valid-card")).FullName;
        Directory.CreateDirectory(Path.Combine(brokenCard, "DCIM"));
        var validDcim = Directory.CreateDirectory(Path.Combine(validCard, "DCIM")).FullName;
        File.WriteAllText(Path.Combine(validDcim, "photo.jpg"), "camera-data");

        var provider = new FakeVolumeProvider(
        [
            // Missing physical-device identity is a safety failure, not a benign empty-card case.
            new MountedVolumeInfo(brokenCard, "volume-broken", true, false, null),
            new MountedVolumeInfo(validCard, "volume-valid", true, false, "device-valid")
        ]);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        using var viewModel = new MainWindowViewModel(provider, sessions, roots);

        await StartupCardDiscovery.InitializeAsync(viewModel, roots);

        Assert.AreEqual(string.Empty, viewModel.SelectedSdPath);
        Assert.AreEqual("スキャン失敗", viewModel.ProgressLabel);
        StringAssert.Contains(viewModel.SafetyDetail, "physical-device identity");
    }

    private sealed class FakeVolumeProvider(IReadOnlyList<MountedVolumeInfo> volumes) : IStorageVolumeProvider
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
                $"PhotoOrganizerStartup-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
