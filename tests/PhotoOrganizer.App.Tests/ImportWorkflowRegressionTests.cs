using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App.Tests;

[TestClass]
[DoNotParallelize]
public sealed class ImportWorkflowRegressionTests
{
    [TestMethod]
    public async Task StartImportAsync_IsBusyBeforeDestinationValidationCompletes()
    {
        using var environment = WorkflowEnvironment.Create();
        using var viewModel = environment.CreateViewModel();
        await environment.PrepareAsync(viewModel);

        environment.Provider.BlockNextEnumeration();
        var firstImport = viewModel.StartImportAsync();
        try
        {
            Assert.IsTrue(
                environment.Provider.EnumerationBlocked.Wait(TimeSpan.FromSeconds(5)),
                "The import did not reach destination validation.");
            Assert.IsTrue(viewModel.IsBusy, "Import must claim busy state before its first await.");

            var secondImport = viewModel.StartImportAsync();
            Assert.IsTrue(
                secondImport.IsCompleted,
                "A second import must be rejected while the first validation is pending.");

            environment.Provider.ReleaseEnumeration.Set();
            await firstImport.WaitAsync(TimeSpan.FromSeconds(15));
            await secondImport;
        }
        finally
        {
            environment.Provider.ReleaseEnumeration.Set();
            try { await firstImport.WaitAsync(TimeSpan.FromSeconds(15)); } catch { }
        }

        Assert.IsFalse(viewModel.IsBusy);
    }

    [TestMethod]
    public async Task StartImportAsync_StopsWhenDestinationChangesDuringValidation()
    {
        using var environment = WorkflowEnvironment.Create();
        using var viewModel = environment.CreateViewModel();
        await environment.PrepareAsync(viewModel);

        environment.Provider.BlockNextEnumeration();
        var import = viewModel.StartImportAsync();
        Assert.IsTrue(
            environment.Provider.EnumerationBlocked.Wait(TimeSpan.FromSeconds(5)),
            "The import did not reach destination validation.");

        viewModel.DestinationPath = environment.DestinationPaths[1];
        environment.Provider.ReleaseEnumeration.Set();
        await import.WaitAsync(TimeSpan.FromSeconds(15));

        Assert.IsFalse(viewModel.IsSafeToReuseCurrentCard);
        Assert.IsFalse(viewModel.HasCompletedImport);
        Assert.AreEqual(string.Empty, viewModel.LastImportBasePath);
        Assert.IsFalse(
            Directory.EnumerateFiles(environment.DestinationPaths[1], "*", SearchOption.AllDirectories).Any(),
            "A changed destination must not receive the stale import.");
    }

    [TestMethod]
    public async Task StartImportAsync_CancelDuringValidationDoesNotStartImport()
    {
        using var environment = WorkflowEnvironment.Create();
        using var viewModel = environment.CreateViewModel();
        await environment.PrepareAsync(viewModel);

        environment.Provider.BlockNextEnumeration();
        var import = viewModel.StartImportAsync();
        Assert.IsTrue(
            environment.Provider.EnumerationBlocked.Wait(TimeSpan.FromSeconds(5)),
            "The import did not reach destination validation.");

        viewModel.CancelImport();
        environment.Provider.ReleaseEnumeration.Set();
        await import.WaitAsync(TimeSpan.FromSeconds(15));

        Assert.IsFalse(viewModel.IsSafeToReuseCurrentCard);
        Assert.IsFalse(viewModel.HasCompletedImport);
        Assert.AreEqual("取り込みキャンセル", viewModel.ProgressLabel);
        Assert.IsFalse(
            Directory.EnumerateFiles(environment.DestinationPaths[0], "*", SearchOption.AllDirectories).Any(),
            "Cancellation during validation must prevent the coordinator from starting.");
    }

    [TestMethod]
    public async Task StartImportAsync_StopsWhenScannedCardIsRemovedDuringValidation()
    {
        using var environment = WorkflowEnvironment.Create();
        using var viewModel = environment.CreateViewModel();
        await environment.PrepareAsync(viewModel);

        environment.Provider.BlockNextEnumeration();
        var import = viewModel.StartImportAsync();
        Assert.IsTrue(
            environment.Provider.EnumerationBlocked.Wait(TimeSpan.FromSeconds(5)),
            "The import did not reach destination validation.");

        environment.Provider.UnmountCard();
        environment.Provider.ReleaseEnumeration.Set();
        await import.WaitAsync(TimeSpan.FromSeconds(15));

        Assert.IsFalse(viewModel.IsSafeToReuseCurrentCard);
        Assert.IsFalse(viewModel.HasCompletedImport);
        Assert.AreEqual("取り込み入力が変更されました", viewModel.ProgressLabel);
    }

    [TestMethod]
    public async Task DestinationPathChange_ClearsReuseApprovalAndCompletionState()
    {
        using var environment = WorkflowEnvironment.Create();
        using var viewModel = environment.CreateViewModel();
        await environment.PrepareAsync(viewModel);

        await viewModel.StartImportAsync();

        Assert.IsTrue(viewModel.IsSafeToReuseCurrentCard);
        Assert.IsTrue(viewModel.HasCompletedImport);
        Assert.IsFalse(string.IsNullOrWhiteSpace(viewModel.LastImportBasePath));

        viewModel.DestinationPath = environment.DestinationPaths[1];

        Assert.IsFalse(viewModel.IsSafeToReuseCurrentCard);
        Assert.IsFalse(viewModel.HasCompletedImport);
        Assert.AreEqual(string.Empty, viewModel.LastImportBasePath);
        Assert.AreEqual(string.Empty, viewModel.CompletionSummary);
    }

    [TestMethod]
    public async Task DuplicateOnlyImport_UsesDistinctCompletionMessage()
    {
        using var environment = WorkflowEnvironment.Create();
        using var viewModel = environment.CreateViewModel();
        await environment.PrepareAsync(viewModel);

        viewModel.EventName = "初回";
        await viewModel.StartImportAsync();
        Assert.IsTrue(viewModel.IsSafeToReuseCurrentCard);

        var rescan = await viewModel.ScanCardAsync(environment.CardPath);
        Assert.IsNotNull(rescan);
        Assert.IsTrue(rescan.IsReady);
        viewModel.EventName = "再確認";

        await viewModel.StartImportAsync();

        Assert.IsTrue(viewModel.IsSafeToReuseCurrentCard);
        StringAssert.Contains(viewModel.CompletionSummary, "新規コピーなし・既存コピーを検証済み");
    }

    [TestMethod]
    public void InjectedPreferencesStore_ReceivesDestinationChanges()
    {
        using var environment = WorkflowEnvironment.Create();
        var preferences = new TestPreferencesStore();
        var sessions = new StorageSessionTracker(environment.Provider);
        var roots = new CameraCardRootResolver(environment.Provider);
        using var viewModel = new MainWindowViewModel(
            environment.Provider,
            sessions,
            roots,
            preferences);

        viewModel.DestinationPath = environment.DestinationPaths[0];

        Assert.AreEqual(environment.DestinationPaths[0], preferences.Preferences.DestinationPath);
    }

    private sealed class WorkflowEnvironment : IDisposable
    {
        private WorkflowEnvironment(
            string root,
            string cardPath,
            IReadOnlyList<string> destinationPaths,
            ControlledVolumeProvider provider)
        {
            Root = root;
            CardPath = cardPath;
            DestinationPaths = destinationPaths;
            Provider = provider;
        }

        private string Root { get; }
        public string CardPath { get; }
        public IReadOnlyList<string> DestinationPaths { get; }
        public ControlledVolumeProvider Provider { get; }

        public static WorkflowEnvironment Create()
        {
            var root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerAppWorkflow-{Guid.NewGuid():N}");
            var card = Path.Combine(root, "card");
            var destinations = new[]
            {
                Path.Combine(root, "destination-a"),
                Path.Combine(root, "destination-b")
            };

            var media = Path.Combine(card, "DCIM", "100CAM");
            Directory.CreateDirectory(media);
            File.WriteAllBytes(Path.Combine(media, "photo.jpg"), [1, 2, 3, 4]);
            File.SetLastWriteTime(Path.Combine(media, "photo.jpg"), new DateTime(2026, 3, 4));
            foreach (var destination in destinations) Directory.CreateDirectory(destination);

            var provider = new ControlledVolumeProvider(
                card,
                [
                    new MountedVolumeInfo(card, "card-volume", true, false, "card-device"),
                    new MountedVolumeInfo(destinations[0], "destination-a-volume", false, false, "destination-a-device"),
                    new MountedVolumeInfo(destinations[1], "destination-b-volume", false, false, "destination-b-device")
                ]);
            return new WorkflowEnvironment(root, card, destinations, provider);
        }

        public MainWindowViewModel CreateViewModel()
        {
            var sessions = new StorageSessionTracker(Provider);
            var roots = new CameraCardRootResolver(Provider);
            return new MainWindowViewModel(
                Provider,
                sessions,
                roots,
                new TestPreferencesStore());
        }

        public async Task PrepareAsync(MainWindowViewModel viewModel)
        {
            var scan = await viewModel.ScanCardAsync(CardPath);
            Assert.IsNotNull(scan);
            Assert.IsTrue(scan.IsReady, scan.Message);
            viewModel.SetDestinationFromPicker(DestinationPaths[0]);
            viewModel.EventName = "撮影";
        }

        public void Dispose()
        {
            Provider.Dispose();
            try { if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true); } catch { }
        }
    }

    private sealed class ControlledVolumeProvider : IStorageVolumeProvider, IDisposable
    {
        private readonly string _cardRoot;
        private readonly IReadOnlyList<MountedVolumeInfo> _volumes;
        private int _blockNextEnumeration;
        private int _cardMounted = 1;

        public ControlledVolumeProvider(string cardRoot, IReadOnlyList<MountedVolumeInfo> volumes)
        {
            _cardRoot = PathSafety.Normalize(cardRoot);
            _volumes = volumes
                .Select(volume => volume with { RootPath = PathSafety.Normalize(volume.RootPath) })
                .ToArray();
        }

        public ManualResetEventSlim EnumerationBlocked { get; } = new(false);
        public ManualResetEventSlim ReleaseEnumeration { get; } = new(false);

        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public void BlockNextEnumeration()
        {
            EnumerationBlocked.Reset();
            ReleaseEnumeration.Reset();
            Volatile.Write(ref _blockNextEnumeration, 1);
        }

        public void UnmountCard() => Volatile.Write(ref _cardMounted, 0);

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes()
        {
            if (Interlocked.Exchange(ref _blockNextEnumeration, 0) == 1)
            {
                EnumerationBlocked.Set();
                if (!ReleaseEnumeration.Wait(TimeSpan.FromSeconds(10)))
                {
                    throw new TimeoutException("The controlled volume enumeration was not released.");
                }
            }

            var cardMounted = Volatile.Read(ref _cardMounted) == 1;
            return _volumes
                .Where(volume => cardMounted
                    || !string.Equals(volume.RootPath, _cardRoot, PathComparison))
                .ToArray();
        }

        public MountedVolumeInfo? ResolveVolumeForPath(string path)
        {
            var normalized = PathSafety.Normalize(path);
            return GetMountedVolumesWithoutGate()
                .Where(volume => PathSafety.IsSameOrDescendant(normalized, volume.RootPath, PathComparison))
                .OrderByDescending(volume => volume.RootPath.Length)
                .FirstOrDefault();
        }

        private IReadOnlyList<MountedVolumeInfo> GetMountedVolumesWithoutGate()
        {
            var cardMounted = Volatile.Read(ref _cardMounted) == 1;
            return _volumes
                .Where(volume => cardMounted
                    || !string.Equals(volume.RootPath, _cardRoot, PathComparison))
                .ToArray();
        }

        public void Dispose()
        {
            EnumerationBlocked.Dispose();
            ReleaseEnumeration.Dispose();
        }
    }
}
