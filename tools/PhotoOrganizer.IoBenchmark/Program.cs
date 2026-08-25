using System.Diagnostics;
using PhotoOrganizer.Core;

var fileCount = ReadInt(args, "--files", 120);
var sizeKiB = ReadInt(args, "--size-kib", 128);
var noiseCount = ReadInt(args, "--noise", 300);
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

    Console.WriteLine($"Photo Organizer I/O benchmark: files={fileCount}, fileSize={fileSize:N0} bytes, libraryNoise={noiseCount}");
    await RunNoCandidateScenario(sources, destinationRoot);

    Directory.Delete(destinationRoot, recursive: true);
    Directory.CreateDirectory(destinationRoot);

    var matchingSources = new List<string>();
    for (var i = 0; i < sources.Count; i += 10)
    {
        var target = Path.Combine(destinationRoot, $"existing_{i:D5}.jpg");
        File.Copy(sources[i], target);
        matchingSources.Add(sources[i]);
    }

    await RunCandidateCacheScenario(sources, matchingSources, destinationRoot, fileSize);

    var process = Process.GetCurrentProcess();
    Console.WriteLine($"peakWorkingSetBytes={process.PeakWorkingSet64:N0}");
    Console.WriteLine($"managedHeapBytes={GC.GetTotalMemory(forceFullCollection: true):N0}");
}
finally
{
    try { Directory.Delete(root, recursive: true); } catch { }
}

static async Task RunNoCandidateScenario(IReadOnlyList<string> sources, string destinationRoot)
{
    var hasher = new CountingHasher();
    var stopwatch = Stopwatch.StartNew();
    var result = await new DestinationLibrary(hasher: hasher)
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
    int fileSize)
{
    var lookupHasher = new CountingHasher();
    var stopwatch = Stopwatch.StartNew();
    var lookup = await new DestinationLibrary(hasher: lookupHasher)
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
    var verification = await new FormatSafetyVerifier(new MediaClassifier(), hasher: verificationHasher)
        .VerifyAsync(matchingSources, destinationRoot);
    stopwatch.Stop();

    var maximumExpectedVerificationBytes = (long)(matchingSources.Count + destinationCandidateCount) * fileSize;
    Console.WriteLine(
        $"fresh-verification: safe={verification.IsSafe}, verified={verification.Verified}/{verification.Total}, " +
        $"hashCalls={verificationHasher.HashCalls}, hashBytesRead={verificationHasher.BytesRead:N0}, " +
        $"maxExpectedWithPer-operationCache={maximumExpectedVerificationBytes:N0}, " +
        $"elapsedMs={stopwatch.Elapsed.TotalMilliseconds:F1}");

    if (!verification.IsSafe)
    {
        throw new InvalidOperationException("Benchmark matching set failed real-byte verification.");
    }

    if (verificationHasher.BytesRead > maximumExpectedVerificationBytes)
    {
        throw new InvalidOperationException("A destination candidate appears to have been rehashed within one fresh verification.");
    }
}

static void WriteDeterministicFile(string path, int size, int seed)
{
    var random = new Random(seed);
    var buffer = new byte[size];
    random.NextBytes(buffer);
    File.WriteAllBytes(path, buffer);
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
    public long BytesRead { get; private set; }
    public int HashCalls { get; private set; }

    public async Task<string> Sha256Async(string path, CancellationToken cancellationToken = default)
    {
        BytesRead += new FileInfo(path).Length;
        HashCalls++;
        return await Hashing.Sha256Async(path, cancellationToken);
    }
}
