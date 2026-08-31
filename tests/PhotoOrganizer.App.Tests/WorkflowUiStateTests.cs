using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App.Tests;

[TestClass]
[DoNotParallelize]
public sealed class WorkflowUiStateTests
{
    [TestMethod]
    public void InitialState_DoesNotShowSdReuseSafety()
    {
        using var temp = new TempDirectory();
        var provider = new TestVolumeProvider(temp.Path);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        using var viewModel = new MainWindowViewModel(provider, sessions, roots);

        Assert.IsFalse(viewModel.ShowSafetyPanel);
        Assert.IsFalse(viewModel.HasSelectedSd);
        Assert.AreEqual("未選択", viewModel.SelectedSdDisplay);
        Assert.AreEqual("SDカードを選択してください", viewModel.WorkflowHeadline);
    }

    [TestMethod]
    public async Task AutoDetectedEmptyCard_HidesSdReuseSafetyAfterBenignScan()
    {
        using var temp = new TempDirectory();
        Directory.CreateDirectory(Path.Combine(temp.Path, "DCIM"));

        var provider = new TestVolumeProvider(temp.Path);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        using var viewModel = new MainWindowViewModel(provider, sessions, roots);

        var result = await viewModel.ScanCardAsync(temp.Path, autoDetected: true);

        Assert.IsNotNull(result);
        Assert.IsTrue(result.IsNoSupportedMedia);
        Assert.IsFalse(viewModel.ShowSafetyPanel);
        Assert.IsFalse(viewModel.HasSelectedSd);
        Assert.AreEqual("SDカードを選択してください", viewModel.WorkflowHeadline);
    }

    [TestMethod]
    public async Task ValidCardScan_ShowsReuseSafetyOnlyAfterCardContextExists()
    {
        using var temp = new TempDirectory();
        var mediaDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "DCIM", "100CAM")).FullName;
        await File.WriteAllBytesAsync(Path.Combine(mediaDirectory, "photo.jpg"), [1, 2, 3, 4]);

        var provider = new TestVolumeProvider(temp.Path);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        using var viewModel = new MainWindowViewModel(provider, sessions, roots);

        Assert.IsFalse(viewModel.ShowSafetyPanel);

        var result = await viewModel.ScanCardAsync(temp.Path, autoDetected: true);

        Assert.IsNotNull(result);
        Assert.IsTrue(result.IsReady);
        Assert.IsTrue(viewModel.HasSelectedSd);
        Assert.IsTrue(viewModel.ShowSafetyPanel);
        StringAssert.Contains(viewModel.SafetyHeadline, "未検証");
        Assert.AreNotEqual("SDカードを選択してください", viewModel.WorkflowHeadline);
    }

    [TestMethod]
    public async Task FailedCardScan_KeepsUiContextWithoutClaimingValidatedSelection()
    {
        using var temp = new TempDirectory();
        var mediaDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "DCIM", "100CAM")).FullName;
        await File.WriteAllBytesAsync(Path.Combine(mediaDirectory, "photo.jpg"), [1, 2, 3, 4]);

        var provider = new TestVolumeProvider(temp.Path, hasPhysicalDeviceIdentity: false);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        using var viewModel = new MainWindowViewModel(provider, sessions, roots);

        var result = await viewModel.ScanCardAsync(temp.Path);

        Assert.IsNotNull(result);
        Assert.IsFalse(result.IsReady);
        Assert.AreEqual(ScanFailureReason.MissingPhysicalDeviceIdentity, result.FailureReason);
        Assert.AreEqual(string.Empty, viewModel.SelectedSdPath);
        Assert.IsTrue(viewModel.HasSelectedSd);
        Assert.AreNotEqual("未選択", viewModel.SelectedSdDisplay);
        Assert.IsTrue(viewModel.ShowSafetyPanel);
        Assert.AreEqual("SDカードを再選択してください", viewModel.WorkflowHeadline);
        Assert.IsFalse(viewModel.CanImport);
    }

    private sealed class TestVolumeProvider : IStorageVolumeProvider
    {
        private readonly MountedVolumeInfo _volume;

        public TestVolumeProvider(string root, bool hasPhysicalDeviceIdentity = true)
        {
            var normalized = PathSafety.Normalize(root);
            _volume = new MountedVolumeInfo(
                normalized,
                "test-volume",
                IsRemovable: true,
                IsSystem: false,
                PhysicalDeviceFingerprint: hasPhysicalDeviceIdentity ? "test-device" : null);
        }

        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes() => [_volume];

        public MountedVolumeInfo? ResolveVolumeForPath(string path)
        {
            var normalized = PathSafety.Normalize(path);
            return PathSafety.IsSameOrDescendant(normalized, _volume.RootPath, PathComparison)
                ? _volume
                : null;
        }
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = Directory.CreateTempSubdirectory("photo-organizer-ui-").FullName;
        }

        public string Path { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
            }
            catch
            {
                // Best-effort test cleanup only.
            }
        }
    }
}
