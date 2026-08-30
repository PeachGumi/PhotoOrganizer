using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App.Tests;

[TestClass]
[DoNotParallelize]
public sealed class StateTransitionPropertyTests
{
    private enum CardKind
    {
        Empty,
        Valid,
        BrokenIdentity,
        Gone
    }

    private sealed record CardCase(string Path, CardKind Kind);

    [TestMethod]
    public async Task RandomizedQueuedScanSequences_RemainFailClosedAndMatchReferenceModel()
    {
        const int scenarioCount = 48;

        for (var seed = 0; seed < scenarioCount; seed++)
        {
            using var temp = new TempDirectory(seed);
            var cards = CreateCards(temp.Path);
            var immutableSources = SnapshotPermanentSourceMedia(cards);

            using var provider = new GateVolumeProvider(cards.Select(ToVolume).ToArray());
            var sessions = new StorageSessionTracker(provider);
            var roots = new CameraCardRootResolver(provider);
            using var viewModel = new MainWindowViewModel(provider, sessions, roots);

            var random = new Random(unchecked(seed * 7919 + 104729));
            var active = cards[random.Next(3)]; // Empty, valid, or broken; never the disposable gone card.
            var firstScan = viewModel.ScanCardAsync(active.Path, autoDetected: true);

            Assert.IsTrue(
                provider.FirstEnumerationEntered.Wait(TimeSpan.FromSeconds(5)),
                $"seed={seed}: first scan never reached the controlled volume enumeration.");
            Assert.IsTrue(viewModel.IsBusy, $"seed={seed}: scan should be busy while the provider gate is held.");

            var modelQueue = new List<string>();
            var operationCount = random.Next(5, 13);
            var cancelled = false;

            for (var operation = 0; operation < operationCount; operation++)
            {
                if (!cancelled && random.Next(5) == 0)
                {
                    viewModel.CancelImport();
                    cancelled = true;
                    continue;
                }

                var queued = cards[random.Next(cards.Count)];
                var queuedResult = await viewModel.ScanCardAsync(queued.Path, autoDetected: true);
                Assert.IsNull(queuedResult, $"seed={seed}: a concurrent scan unexpectedly replaced the active scan.");
                ModelQueue(modelQueue, queued.Path, provider.PathComparison);
            }

            // Randomly make the disposable queued card disappear before queue draining.
            var gone = cards.Single(card => card.Kind == CardKind.Gone);
            var removeGone = random.Next(2) == 0;
            if (removeGone && Directory.Exists(gone.Path))
            {
                Directory.Delete(gone.Path, recursive: true);
            }

            provider.ReleaseFirstEnumeration.Set();
            var firstResult = await firstScan.WaitAsync(TimeSpan.FromSeconds(10));

            var expected = EvaluateModel(
                active,
                modelQueue,
                cards,
                cancelled,
                removeGone,
                provider.PathComparison);

            Assert.IsFalse(viewModel.IsBusy, $"seed={seed}: workflow remained busy after the controlled scan completed.");
            Assert.AreEqual(expected.SelectedPath, viewModel.SelectedSdPath, $"seed={seed}: selected-card state diverged from model.");
            Assert.AreEqual(expected.PendingCount, viewModel.PendingSdCount, $"seed={seed}: pending queue diverged from model.");
            Assert.AreNotEqual(
                "保存先コピー検証済み — SDカード再利用可能",
                viewModel.SafetyHeadline,
                $"seed={seed}: scan/queue activity alone must never produce reuse approval.");

            if (cancelled)
            {
                Assert.IsNull(firstResult, $"seed={seed}: cancelled scan should not return a successful/blocked scan result.");
                Assert.AreEqual("スキャンキャンセル", viewModel.ProgressLabel, $"seed={seed}: cancellation must remain visibly unverified.");
            }
            else if (active.Kind == CardKind.Valid)
            {
                Assert.IsNotNull(firstResult, $"seed={seed}: valid active card should produce a scan result.");
                Assert.IsTrue(firstResult.IsReady, $"seed={seed}: valid active card should be ready.");
            }
            else if (active.Kind == CardKind.BrokenIdentity)
            {
                Assert.IsNotNull(firstResult, $"seed={seed}: broken identity should return a blocked result.");
                Assert.AreEqual(ScanFailureReason.MissingPhysicalDeviceIdentity, firstResult.FailureReason, $"seed={seed}");
            }
            else
            {
                Assert.IsNotNull(firstResult, $"seed={seed}: empty card should return a benign no-media result.");
                Assert.IsTrue(firstResult.IsNoSupportedMedia, $"seed={seed}: empty card must be classified as no supported media.");
            }

            foreach (var snapshot in immutableSources)
            {
                Assert.IsTrue(File.Exists(snapshot.Key), $"seed={seed}: source media disappeared: {snapshot.Key}");
                CollectionAssert.AreEqual(
                    snapshot.Value,
                    File.ReadAllBytes(snapshot.Key),
                    $"seed={seed}: source media bytes changed: {snapshot.Key}");
            }
        }
    }

    private static IReadOnlyList<CardCase> CreateCards(string root)
    {
        var empty = Directory.CreateDirectory(Path.Combine(root, "00-empty")).FullName;
        Directory.CreateDirectory(Path.Combine(empty, "DCIM"));

        var validA = Directory.CreateDirectory(Path.Combine(root, "01-valid-a")).FullName;
        var validADcim = Directory.CreateDirectory(Path.Combine(validA, "DCIM", "100CAM")).FullName;
        File.WriteAllBytes(Path.Combine(validADcim, "a.jpg"), [1, 2, 3, 4, 5]);

        var broken = Directory.CreateDirectory(Path.Combine(root, "02-broken")).FullName;
        var brokenDcim = Directory.CreateDirectory(Path.Combine(broken, "DCIM")).FullName;
        File.WriteAllBytes(Path.Combine(brokenDcim, "broken.jpg"), [6, 7, 8, 9]);

        var validB = Directory.CreateDirectory(Path.Combine(root, "03-valid-b")).FullName;
        var validBDcim = Directory.CreateDirectory(Path.Combine(validB, "DCIM")).FullName;
        File.WriteAllBytes(Path.Combine(validBDcim, "b.jpg"), [10, 11, 12]);

        var gone = Directory.CreateDirectory(Path.Combine(root, "04-gone")).FullName;
        var goneDcim = Directory.CreateDirectory(Path.Combine(gone, "DCIM")).FullName;
        File.WriteAllBytes(Path.Combine(goneDcim, "gone.jpg"), [13, 14, 15]);

        return
        [
            new CardCase(empty, CardKind.Empty),
            new CardCase(validA, CardKind.Valid),
            new CardCase(broken, CardKind.BrokenIdentity),
            new CardCase(validB, CardKind.Valid),
            new CardCase(gone, CardKind.Gone)
        ];
    }

    private static Dictionary<string, byte[]> SnapshotPermanentSourceMedia(IEnumerable<CardCase> cards) =>
        cards
            .Where(card => card.Kind is not CardKind.Gone)
            .SelectMany(card => Directory.EnumerateFiles(card.Path, "*", SearchOption.AllDirectories))
            .ToDictionary(Path.GetFullPath, File.ReadAllBytes, PathComparer.Instance);

    private static MountedVolumeInfo ToVolume(CardCase card) => new(
        card.Path,
        $"volume-{Path.GetFileName(card.Path)}",
        IsRemovable: true,
        IsSystem: false,
        card.Kind == CardKind.BrokenIdentity ? null : $"device-{Path.GetFileName(card.Path)}");

    private static void ModelQueue(List<string> queue, string path, StringComparison comparison)
    {
        var normalized = PathSafety.Normalize(path);
        if (!queue.Any(existing => string.Equals(existing, normalized, comparison)))
        {
            queue.Add(normalized);
        }
    }

    private static (string SelectedPath, int PendingCount) EvaluateModel(
        CardCase active,
        List<string> queued,
        IReadOnlyList<CardCase> cards,
        bool cancelled,
        bool goneRemoved,
        StringComparison comparison)
    {
        var remaining = queued.ToList();

        if (cancelled)
        {
            return (string.Empty, remaining.Count);
        }

        if (active.Kind == CardKind.Valid)
        {
            remaining.RemoveAll(path => string.Equals(path, PathSafety.Normalize(active.Path), comparison));
            return (PathSafety.Normalize(active.Path), remaining.Count);
        }

        if (active.Kind == CardKind.BrokenIdentity)
        {
            return (string.Empty, remaining.Count);
        }

        while (remaining.Count > 0)
        {
            var next = remaining[0];
            remaining.RemoveAt(0);
            var card = cards.Single(candidate => string.Equals(PathSafety.Normalize(candidate.Path), next, comparison));

            if (card.Kind == CardKind.Gone && goneRemoved)
            {
                continue;
            }

            if (card.Kind == CardKind.Empty)
            {
                continue;
            }

            if (card.Kind == CardKind.BrokenIdentity)
            {
                return (string.Empty, remaining.Count);
            }

            return (PathSafety.Normalize(card.Path), remaining.Count);
        }

        return (string.Empty, 0);
    }

    private sealed class GateVolumeProvider(IReadOnlyList<MountedVolumeInfo> volumes)
        : IStorageVolumeProvider, IDisposable
    {
        private int _gateFirstEnumeration = 1;

        public ManualResetEventSlim FirstEnumerationEntered { get; } = new(false);
        public ManualResetEventSlim ReleaseFirstEnumeration { get; } = new(false);

        public StringComparison PathComparison => OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes()
        {
            if (Interlocked.Exchange(ref _gateFirstEnumeration, 0) == 1)
            {
                FirstEnumerationEntered.Set();
                if (!ReleaseFirstEnumeration.Wait(TimeSpan.FromSeconds(10)))
                {
                    throw new TimeoutException("Property test did not release the first volume enumeration.");
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
        public TempDirectory(int seed)
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PhotoOrganizerStateProperty-{seed}-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
