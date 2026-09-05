using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class ImportBatchDrainSafetyTests
{
    private static readonly TimeSpan TestTimeout = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan PrematureReturnTimeout = TimeSpan.FromMilliseconds(200);

    [TestMethod]
    public async Task StorageCheckException_DrainsStartedCopyBeforeReturning()
    {
        EnsureBatchCanContainStartedAndPendingWork();
        using var env = DrainEnvironment.Create();
        env.AddCameraFile("DCIM/one.jpg", "first-data", new DateTime(2026, 5, 1));
        env.AddCameraFile("DCIM/two.jpg", "second-data", new DateTime(2026, 5, 1));

        var trigger = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        Assert.IsTrue(scan.IsReady, scan.Message);

        var importTask = StartImportWithCopyingGate(
            coordinator,
            scan.Session!,
            env,
            () =>
            {
                env.Durability.FirstFinalizeEntered.WaitAsync(TestTimeout).GetAwaiter().GetResult();
                trigger.TrySetResult(true);
                throw new InvalidOperationException("simulated storage-check failure");
            });

        ImportRunResult? result = null;
        try
        {
            await trigger.Task.WaitAsync(TestTimeout).ConfigureAwait(false);
            await AssertDoesNotCompleteWithin(importTask);
        }
        finally
        {
            env.Durability.ReleaseFirstFinalize();
            result = await importTask.WaitAsync(TestTimeout).ConfigureAwait(false);
        }

        Assert.IsNotNull(result);
        Assert.AreEqual(ImportSafetyStatus.Blocked, result!.Status);
        StringAssert.Contains(result.Message, "Import failed");
        AssertNoPartials(env.DestinationRoot);
    }

    [TestMethod]
    public async Task StorageCheckCancellation_DrainsStartedCopyBeforeReturning()
    {
        EnsureBatchCanContainStartedAndPendingWork();
        using var env = DrainEnvironment.Create();
        env.AddCameraFile("DCIM/one.jpg", "first-data", new DateTime(2026, 5, 2));
        env.AddCameraFile("DCIM/two.jpg", "second-data", new DateTime(2026, 5, 2));

        using var cancellation = new CancellationTokenSource();
        var trigger = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var coordinator = env.CreateCoordinator();
        var scan = coordinator.ScanCard(env.CardRoot);
        Assert.IsTrue(scan.IsReady, scan.Message);

        var importTask = StartImportWithCopyingGate(
            coordinator,
            scan.Session!,
            env,
            () =>
            {
                env.Durability.FirstFinalizeEntered.WaitAsync(TestTimeout).GetAwaiter().GetResult();
                cancellation.Cancel();
                trigger.TrySetResult(true);
                throw new OperationCanceledException(cancellation.Token);
            },
            cancellation.Token);

        ImportRunResult? result = null;
        try
        {
            await trigger.Task.WaitAsync(TestTimeout).ConfigureAwait(false);
            await AssertDoesNotCompleteWithin(importTask);
        }
        finally
        {
            env.Durability.ReleaseFirstFinalize();
            result = await importTask.WaitAsync(TestTimeout).ConfigureAwait(false);
        }

        Assert.IsNotNull(result);
        Assert.AreEqual(ImportSafetyStatus.Blocked, result!.Status);
        StringAssert.Contains(result.Message, "cancelled");
        AssertNoPartials(env.DestinationRoot);
    }

    private static Task<ImportRunResult> StartImportWithCopyingGate(
        ImportCoordinator coordinator,
        ImportScanSession session,
        DrainEnvironment env,
        Action storageCheckFailure,
        CancellationToken cancellationToken = default)
    {
        return coordinator.ImportAsync(
            session,
            env.DestinationRoot,
            "Drain",
            new ArmOnCopyingProgress(env.Provider, storageCheckFailure),
            cancellationToken);
    }

    private static void EnsureBatchCanContainStartedAndPendingWork()
    {
        if (Math.Clamp(Environment.ProcessorCount, 1, 4) < 2)
        {
            Assert.Inconclusive("The batch-drain regression requires at least two concurrent I/O slots.");
        }
    }

    private static void AssertNoPartials(string destinationRoot)
    {
        Assert.IsFalse(
            Directory.EnumerateFiles(destinationRoot, ".partial-*", SearchOption.AllDirectories).Any(),
            "ImportAsync returned before a started copy cleaned up its partial file.");
    }

    private static async Task AssertDoesNotCompleteWithin(Task task)
    {
        try
        {
            await task.WaitAsync(PrematureReturnTimeout).ConfigureAwait(false);
            Assert.Fail("ImportAsync returned before draining the started copy.");
        }
        catch (TimeoutException)
        {
            // Expected while the durability gate is intentionally closed.
        }
    }

    private sealed class ArmOnCopyingProgress(TriggerVolumeProvider provider, Action action) : IProgress<ImportProgress>
    {
        private bool _armed;

        public void Report(ImportProgress value)
        {
            if (_armed || value.Phase != ImportProgressPhase.Copying) return;
            _armed = true;
            provider.ArmSecondStorageCheck(action);
        }
    }

    private sealed class DrainEnvironment : IDisposable
    {
        private DrainEnvironment(
            string root,
            string cardRoot,
            string destinationRoot,
            TriggerVolumeProvider provider,
            GateDurabilityService durability)
        {
            Root = root;
            CardRoot = cardRoot;
            DestinationRoot = destinationRoot;
            Provider = provider;
            Durability = durability;
            Tracker = new StorageSessionTracker(provider);
        }

        public string Root { get; }
        public string CardRoot { get; }
        public string DestinationRoot { get; }
        public TriggerVolumeProvider Provider { get; }
        public GateDurabilityService Durability { get; }
        public StorageSessionTracker Tracker { get; }

        public static DrainEnvironment Create()
        {
            var root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerBatchDrain-{Guid.NewGuid():N}");
            var card = Path.Combine(root, "card");
            var destination = Path.Combine(root, "destination");
            Directory.CreateDirectory(Path.Combine(card, "DCIM"));
            Directory.CreateDirectory(destination);

            var provider = new TriggerVolumeProvider(
                new MountedVolumeInfo(card, "camera-volume", true, false, "physical-camera"),
                new MountedVolumeInfo(destination, "destination-volume", false, false, "physical-destination"));
            var durability = new GateDurabilityService();
            return new DrainEnvironment(root, card, destination, provider, durability);
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
            Provider,
            Durability);

        public void Dispose()
        {
            Durability.ReleaseFirstFinalize();
            try { Directory.Delete(Root, recursive: true); } catch { }
        }
    }

    private sealed class GateDurabilityService : IFileDurabilityService
    {
        private readonly TaskCompletionSource<bool> _firstFinalizeEntered =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<bool> _releaseFirstFinalize =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _finalizeCalls;

        public Task FirstFinalizeEntered => _firstFinalizeEntered.Task;

        public void ReleaseFirstFinalize() => _releaseFirstFinalize.TrySetResult(true);

        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc)
        {
            if (Interlocked.Increment(ref _finalizeCalls) == 1)
            {
                _firstFinalizeEntered.TrySetResult(true);
                _releaseFirstFinalize.Task.WaitAsync(TestTimeout).GetAwaiter().GetResult();
            }

            try
            {
                File.Move(temporaryPath, finalPath, overwrite: false);
                return new FinalizeFileResult(FinalizeFileStatus.Committed, true);
            }
            catch (Exception ex)
            {
                return new FinalizeFileResult(FinalizeFileStatus.Failed, false, ex.Message);
            }
        }

        public DurabilityResult EnsureDurable(string filePath) => new(true);
    }

    private sealed class TriggerVolumeProvider(MountedVolumeInfo first, MountedVolumeInfo second) : IStorageVolumeProvider
    {
        private Action? _pendingAction;
        private int _checksAfterArm;

        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public void ArmSecondStorageCheck(Action action)
        {
            _pendingAction = action;
            Volatile.Write(ref _checksAfterArm, 0);
        }

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes()
        {
            var check = Interlocked.Increment(ref _checksAfterArm);
            if (check == 2)
            {
                _pendingAction?.Invoke();
            }

            return [first, second];
        }

        public MountedVolumeInfo? ResolveVolumeForPath(string path)
        {
            var normalized = PathSafety.Normalize(path);
            return PathSafety.IsSameOrDescendant(normalized, first.RootPath, PathComparison)
                ? first
                : second;
        }
    }
}
