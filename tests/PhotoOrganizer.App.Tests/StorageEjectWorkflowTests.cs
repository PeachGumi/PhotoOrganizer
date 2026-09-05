using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App.Tests;

[TestClass]
[DoNotParallelize]
public sealed class StorageEjectWorkflowTests
{
    [TestMethod]
    public async Task VerifiedImport_AllowsEjectAndClearsSelectedCardAfterSuccess()
    {
        using var card = new TempDirectory("photo-organizer-eject-card-");
        using var destination = new TempDirectory("photo-organizer-eject-destination-");
        var mediaDirectory = Directory.CreateDirectory(Path.Combine(card.Path, "DCIM", "100CAM")).FullName;
        await File.WriteAllBytesAsync(Path.Combine(mediaDirectory, "photo.jpg"), [1, 2, 3, 4]);

        var provider = new TestVolumeProvider(card.Path, destination.Path);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        var eject = new RecordingEjectService(StorageEjectResult.Succeeded("取り出しました"));
        using var viewModel = new MainWindowViewModel(provider, sessions, roots, eject, new TestPreferencesStore());

        var scan = await viewModel.ScanCardAsync(card.Path);
        Assert.IsNotNull(scan);
        Assert.IsTrue(scan.IsReady);
        viewModel.SetDestinationFromPicker(destination.Path);
        viewModel.EventName = "撮影";
        await viewModel.StartImportAsync();

        Assert.IsTrue(viewModel.IsSafeToReuseCurrentCard);
        Assert.IsTrue(viewModel.CanEjectSelectedSd);

        await viewModel.EjectSelectedSdAsync();

        Assert.AreEqual(card.Path, eject.RequestedRoot);
        Assert.IsFalse(viewModel.HasSelectedSd);
        Assert.IsFalse(viewModel.CanEjectSelectedSd);
        Assert.AreEqual("SDカードを安全に取り出しました", viewModel.ProgressLabel);
    }

    [TestMethod]
    public async Task EjectFailure_PreservesVerifiedCardAndExplainsRetry()
    {
        using var card = new TempDirectory("photo-organizer-eject-card-");
        using var destination = new TempDirectory("photo-organizer-eject-destination-");
        var mediaDirectory = Directory.CreateDirectory(Path.Combine(card.Path, "DCIM", "100CAM")).FullName;
        await File.WriteAllBytesAsync(Path.Combine(mediaDirectory, "photo.jpg"), [1, 2, 3, 4]);

        var provider = new TestVolumeProvider(card.Path, destination.Path);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        var eject = new RecordingEjectService(StorageEjectResult.Failed("別のアプリが使用中です"));
        using var viewModel = new MainWindowViewModel(provider, sessions, roots, eject, new TestPreferencesStore());

        await viewModel.ScanCardAsync(card.Path);
        viewModel.SetDestinationFromPicker(destination.Path);
        viewModel.EventName = "撮影";
        await viewModel.StartImportAsync();
        await viewModel.EjectSelectedSdAsync();

        Assert.IsTrue(viewModel.HasSelectedSd);
        Assert.IsTrue(viewModel.IsSafeToReuseCurrentCard);
        Assert.IsTrue(viewModel.CanEjectSelectedSd);
        Assert.AreEqual("SDカードを取り出せませんでした", viewModel.SafetyHeadline);
        StringAssert.Contains(viewModel.SafetyDetail, "別のアプリが使用中です");
    }

    [TestMethod]
    public async Task ChangedPhysicalIdentity_BlocksEjectWithoutCallingPlatform()
    {
        using var card = new TempDirectory("photo-organizer-eject-card-");
        using var destination = new TempDirectory("photo-organizer-eject-destination-");
        var mediaDirectory = Directory.CreateDirectory(Path.Combine(card.Path, "DCIM", "100CAM")).FullName;
        await File.WriteAllBytesAsync(Path.Combine(mediaDirectory, "photo.jpg"), [1, 2, 3, 4]);

        var provider = new TestVolumeProvider(card.Path, destination.Path);
        var sessions = new StorageSessionTracker(provider);
        var roots = new CameraCardRootResolver(provider);
        var eject = new RecordingEjectService(StorageEjectResult.Succeeded("取り出しました"));
        using var viewModel = new MainWindowViewModel(provider, sessions, roots, eject, new TestPreferencesStore());

        await viewModel.ScanCardAsync(card.Path);
        viewModel.SetDestinationFromPicker(destination.Path);
        viewModel.EventName = "撮影";
        await viewModel.StartImportAsync();
        provider.ChangeCardIdentity();

        await viewModel.EjectSelectedSdAsync();

        Assert.IsNull(eject.RequestedRoot);
        Assert.IsFalse(viewModel.IsSafeToReuseCurrentCard);
        Assert.IsFalse(viewModel.CanEjectSelectedSd);
        StringAssert.Contains(viewModel.SafetyDetail, "物理デバイス情報が変化");
    }

    private sealed class RecordingEjectService(StorageEjectResult result) : IStorageEjectService
    {
        public bool IsSupported => true;
        public string? RequestedRoot { get; private set; }

        public StorageEjectResult Eject(MountedVolumeInfo volume)
        {
            RequestedRoot = volume.RootPath;
            return result;
        }
    }

    private sealed class TestVolumeProvider : IStorageVolumeProvider
    {
        private readonly MountedVolumeInfo[] _volumes;

        public TestVolumeProvider(string cardRoot, string destinationRoot)
        {
            _volumes =
            [
                new MountedVolumeInfo(cardRoot, "card-volume", true, false, "card-device"),
                new MountedVolumeInfo(destinationRoot, "destination-volume", false, false, "destination-device")
            ];
        }

        public StringComparison PathComparison => StringComparison.Ordinal;
        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes() => _volumes;

        public void ChangeCardIdentity() =>
            _volumes[0] = _volumes[0] with
            {
                Fingerprint = "replacement-volume",
                PhysicalDeviceFingerprint = "replacement-device"
            };

        public MountedVolumeInfo? ResolveVolumeForPath(string path) => _volumes
            .Where(volume => PathSafety.IsSameOrDescendant(path, volume.RootPath, PathComparison))
            .OrderByDescending(volume => volume.RootPath.Length)
            .FirstOrDefault();
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory(string prefix) =>
            Path = Directory.CreateTempSubdirectory(prefix).FullName;

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
        }
    }
}
