using System.Diagnostics;
using PhotoOrganizer.Core;

var fileCount = ReadInt(args, "--files", 120);
var sizeKiB = ReadInt(args, "--size-kib", 128);
var noiseCount = ReadInt(args, "--noise", 300);
var defaultParallelism = Math.Clamp(Environment.ProcessorCount, 1, 4);
var copyFileCount = ReadInt(args, "--copy-files", 4);
var copySizeMiB = ReadInt(args, "--copy-size-mib", 32);
var copyParallelism = ReadInt(args, "--copy-parallelism", defaultParallelism);
var hashParallelism = ReadInt(args, "--hash-parallelism", defaultParallelism);
var fileSize = checked(sizeKiB * 1024);

var root = Path.Combine(Path.GetTempPath(), "PhotoOrganizerIoBenchmark-" + Guid.NewGuid().ToString("N"));
var sourceRoot = Directory.CreateDirectory(Path.Combine(root, "source")).FullName;
var destinationRoot = Directory.CreateDirectory(Path.Combine(root, "destination")).FullName;

try
{
    var sources = new List<string>(fileCount);
    for (var i = 0; i < fileCount; i++)
    {
        var path = Path.Combine(sourceRoot, $"IMG_{i:D5}.jpg");
        WriteDeterministicFile(path, fileSize, seed: i + 1);
        sources.Add(path);
    }

    for (var i = 0; i < noiseCount; i++)
    {
        // Deliberately avoid the source size so this large metadata library has no
        // possible byte-identical candidate and therefore requires zero hashing.
        var path = Path.Combine(destinationRoot, $"noise_{i:D5}.bin");
        WriteDeterministicFile(path, fileSize + 1 + (i % 7), seed: 100_000 + i);
    }

    Console.WriteLine(
        $"Photo Organizer I/O benchmark: files={fileCount}, fileSize={fileSize:N0} bytes, " +
        $"libraryNoise={noiseCount}, hashParallelism={hashParallelism}");
    await RunNoCandidateScenario(sources, destinationRoot, hashParallelism);

    Directory.Delete(destinationRoot, recursive: true);
    Directory.CreateDirectory(destinationRoot);

    var matchingSources = new List<string>();
    for (var i = 0; i < sources.Count; i += 10)
    {
        var target = Path.Combine(destinationRoot, $"existing_{i:D5}.jpg");
        File.Copy(sources[i], target);
        matchingSources.Add(sources[i]);
    }

    await RunCandidateCacheScenario(sources, matchingSources, destinationRoot, fileSize, hashParallelism);

    await RunSafeCopyScenario(root, copyFileCount, copySizeMiB, copyParallelism);
    await RunEndToEndImportScenario(root, copyFileCount, copySizeMiB);

    var process = Process.GetCurrentProcess();
    Console.WriteLine($"peakWorkingSetBytes={process.PeakWorkingSet64:N0}");
    Console.WriteLine($"managedHeapBytes={GC.GetTotalMemory(forceFullCollection: true):N0}");
}
finally
{
    try { Directory.Delete(root, recursive: true); } catch { }
}

static async Task RunSafeCopyScenario(string root, int fileCount, int sizeMiB, int parallelism)
{
    var sourceRoot = Directory.CreateDirectory(Path.Combine(root, "copy-source")).FullName;
    var destinationRoot = Directory.CreateDirectory(Path.Combine(root, "copy-destination")).FullName;
    var fileSize = checked(sizeMiB * 1024 * 1024);
    var sources = new List<string>(fileCount);

    for (var i = 0; i < fileCount; i++)
    {
        var path = Path.Combine(sourceRoot, $"COPY_{i:D3}.mov");
        WriteDeterministicFile(path, fileSize, seed: 200_000 + i);
        sources.Add(path);
    }

    var stopwatch = Stopwatch.StartNew();
    await Parallel.ForEachAsync(
        sources,
        new ParallelOptions { MaxDegreeOfParallelism = parallelism },
        async (source, cancellationToken) =>
        {
            var result = await new SafeCopyService().CopyAsync(source, destinationRoot, cancellationToken);
            if (result.Status != CopyStatus.Copied)
            {
                throw new InvalidOperationException($"Safe copy benchmark failed: {result.Error}");
            }
        });
    stopwatch.Stop();

    var logicalBytes = (long)fileCount * fileSize;
    var logicalMiBPerSecond = logicalBytes / 1024d / 1024d / stopwatch.Elapsed.TotalSeconds;
    Console.WriteLine(
        $"safe-copy: files={fileCount}, fileSizeMiB={sizeMiB}, parallelism={parallelism}, logicalBytes={logicalBytes:N0}, " +
        $"elapsedMs={stopwatch.Elapsed.TotalMilliseconds:F1}, logicalMiBPerSecond={logicalMiBPerSecond:F1}");
}

static async Task RunEndToEndImportScenario(string root, int fileCount, int sizeMiB)
{
    var cardRoot = Directory.CreateDirectory(Path.Combine(root, "import-card")).FullName;
    var dcimRoot = Directory.CreateDirectory(Path.Combine(cardRoot, "DCIM")).FullName;
    var destinationRoot = Directory.CreateDirectory(Path.Combine(root, "import-destination")).FullName;
    var fileSize = checked(sizeMiB * 1024 * 1024);

    for (var i = 0; i < fileCount; i++)
    {
        var path = Path.Combine(dcimRoot, $"IMPORT_{i:D3}.mov");
        WriteDeterministicFile(path, fileSize, seed: 300_000 + i);
        File.SetLastWriteTime(path, new DateTime(2026, 8, 31, 12, 0, 0).AddSeconds(i));
    }

    var provider = new BenchmarkVolumeProvider(cardRoot, destinationRoot);
    var tracker = new StorageSessionTracker(provider);
    var classifier = new MediaClassifier();
    var coordinator = new ImportCoordinator(
        classifier,
        tracker,
        new CameraCardRootResolver(provider),
        provider);
    var scan = coordinator.ScanCard(cardRoot);
    if (!scan.IsReady)
    {
        throw new InvalidOperationException($"End-to-end benchmark scan failed: {scan.Message}");
    }

    var stopwatch = Stopwatch.StartNew();
    var result = await coordinator.ImportAsync(
        scan.Session!,
        destinationRoot,
        "Benchmark");
    stopwatch.Stop();

    if (!result.IsSafeToReuse)
    {
        throw new InvalidOperationException($"End-to-end benchmark import failed: {result.Message}");
    }

    var logicalBytes = (long)fileCount * fileSize;
    var logicalMiBPerSecond = logicalBytes / 1024d / 1024d / stopwatch.Elapsed.TotalSeconds;
    Console.WriteLine(
        $"end-to-end-import: files={fileCount}, fileSizeMiB={sizeMiB}, logicalBytes={logicalBytes:N0}, " +
        $"elapsedMs={stopwatch.Elapsed.TotalMilliseconds:F1}, logicalMiBPerSecond={logicalMiBPerSecond:F1}");
}

static async Task RunNoCandidateScenario(
    IReadOnlyList<string> sources,
    string destinationRoot,
    int hashParallelism)
{
    var hasher = new CountingHasher();
    var stopwatch = Stopwatch.StartNew();
    var result = await new DestinationLibrary(
            hasher: hasher,
            maxDegreeOfParallelism: hashParallelism)
        .FindVerifiedBackupsAsync(sources, destinationRoot);
    stopwatch.Stop();

    Console.WriteLine(
        $"no-candidate: matched={result.MatchedSources.Count}, hashCalls={hasher.HashCalls}, " +
        $"hashBytesRead={hasher.BytesRead:N0}, elapsedMs={stopwatch.Elapsed.TotalMilliseconds:F1}");

    if (hasher.BytesRead != 0)
    {
        throw new InvalidOperationException("No-candidate lookup unexpectedly hashed file bytes.");
    }
}

static async Task RunCandidateCacheScenario(
    IReadOnlyList<string> allSources,
    IReadOnlyList<string> matchingSources,
    string destinationRoot,
    int fileSize,
    int hashParallelism)
{
    var lookupHasher = new CountingHasher();
    var stopwatch = Stopwatch.StartNew();
    var lookup = await new DestinationLibrary(
            hasher: lookupHasher,
            maxDegreeOfParallelism: hashParallelism)
        .FindVerifiedBackupsAsync(allSources, destinationRoot);
    stopwatch.Stop();

    var destinationCandidateCount = Directory.EnumerateFiles(destinationRoot).Count();
    var maximumExpectedLookupBytes = (long)(allSources.Count + destinationCandidateCount) * fileSize;
    Console.WriteLine(
        $"candidate-cache: matched={lookup.MatchedSources.Count}, hashCalls={lookupHasher.HashCalls}, " +
        $"hashBytesRead={lookupHasher.BytesRead:N0}, maxExpectedWithPer-operationCache={maximumExpectedLookupBytes:N0}, " +
        $"elapsedMs={stopwatch.Elapsed.TotalMilliseconds:F1}");

    if (lookupHasher.BytesRead > maximumExpectedLookupBytes)
    {
        throw new InvalidOperationException("A destination candidate appears to have been hashed more than once in one lookup.");
    }

    var verificationHasher = new CountingHasher();
    stopwatch.Restart();
    var verification = await new FormatSafetyVerifier(
            new MediaClassifier(),
            hasher: verificationHasher,
            maxDegreeOfParallelism: hashParallelism)
        .VerifyAsync(matchingSources, destinationRoot);
    stopwatch.Stop();

    // Final reuse verification reads each successful source before matching and again
    // after destination durability proof. A successful destination candidate is also
    // read once as a prefilter and once after durable synchronization. This four-read
    // budget is intentional: the final source read prevents approving media that was
    // modified while destination durability was being established.
    var maximumExpectedVerificationBytes =
        (long)((matchingSources.Count * 2) + (destinationCandidateCount * 2)) * fileSize;
    Console.WriteLine(
        $"fresh-verification: safe={verification.IsSafe}, verified={verification.Verified}/{verification.Total}, " +
        $"hashCalls={verificationHasher.HashCalls}, hashBytesRead={verificationHasher.BytesRead:N0}, " +
        $"maxExpectedWithFreshSourceRehash={maximumExpectedVerificationBytes:N0}, " +
        $"elapsedMs={stopwatch.Elapsed.TotalMilliseconds:F1}");

    if (!verification.IsSafe)
    {
        throw new InvalidOperationException("Benchmark matching set failed real-byte verification.");
    }

    if (verificationHasher.BytesRead > maximumExpectedVerificationBytes)
    {
        throw new InvalidOperationException("Final verification exceeded the expected source/destination rehash budget.");
    }
}

static void WriteDeterministicFile(string path, int size, int seed)
{
    var random = new Random(seed);
    var buffer = new byte[Math.Min(size, 1024 * 1024)];
    using var stream = new FileStream(
        path,
        FileMode.CreateNew,
        FileAccess.Write,
        FileShare.None,
        bufferSize: buffer.Length,
        FileOptions.SequentialScan);

    var remaining = size;
    while (remaining > 0)
    {
        var count = Math.Min(remaining, buffer.Length);
        random.NextBytes(buffer.AsSpan(0, count));
        stream.Write(buffer, 0, count);
        remaining -= count;
    }
}

static int ReadInt(string[] args, string key, int fallback)
{
    foreach (var argument in args)
    {
        if (!argument.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase)) continue;
        if (int.TryParse(argument[(key.Length + 1)..], out var value) && value > 0) return value;
    }
    return fallback;
}

sealed class CountingHasher : IFileHasher
{
    private long _bytesRead;
    private int _hashCalls;

    public long BytesRead => Interlocked.Read(ref _bytesRead);
    public int HashCalls => Volatile.Read(ref _hashCalls);

    public async Task<string> Sha256Async(string path, CancellationToken cancellationToken = default)
    {
        Interlocked.Add(ref _bytesRead, new FileInfo(path).Length);
        Interlocked.Increment(ref _hashCalls);
        return await Hashing.Sha256Async(path, cancellationToken);
    }
}

sealed class BenchmarkVolumeProvider : IStorageVolumeProvider
{
    private readonly IReadOnlyList<MountedVolumeInfo> _volumes;

    public BenchmarkVolumeProvider(string cardRoot, string destinationRoot)
    {
        _volumes =
        [
            new MountedVolumeInfo(
                PathSafety.Normalize(cardRoot),
                "benchmark-card-volume",
                true,
                false,
                "benchmark-card-device"),
            new MountedVolumeInfo(
                PathSafety.Normalize(destinationRoot),
                "benchmark-destination-volume",
                false,
                false,
                "benchmark-destination-device")
        ];
    }

    public StringComparison PathComparison => OperatingSystem.IsWindows()
        ? StringComparison.OrdinalIgnoreCase
        : StringComparison.Ordinal;

    public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes() => _volumes;

    public MountedVolumeInfo? ResolveVolumeForPath(string path)
    {
        var normalized = PathSafety.Normalize(path);
        return _volumes
            .Where(volume => PathSafety.IsSameOrDescendant(normalized, volume.RootPath, PathComparison))
            .OrderByDescending(volume => volume.RootPath.Length)
            .FirstOrDefault();
    }
}
