using System.Diagnostics;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
[DoNotParallelize]
public sealed class ProcessCrashSafetyTests
{
    private static readonly string[] FinalizationCrashModes =
    [
        "before-finalize",
        "after-finalize-before-durable",
        "after-durable-before-return"
    ];

    [TestMethod]
    public async Task CrashMatrix_PreservesSourceNeverOverwritesAndRecoversSafely()
    {
        const int iterations = 18;

        for (var iteration = 0; iteration < iterations; iteration++)
        {
            using var temp = new TempDirectory();
            var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "card")).FullName;
            var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "library")).FullName;
            var source = Path.Combine(sourceDirectory, "photo.jpg");
            var checkpoint = Path.Combine(temp.Path, "checkpoint.txt");
            var payload = CreatePayload(iteration + 1000);
            File.WriteAllBytes(source, payload);

            var hasCollision = iteration % 2 == 1;
            var competitorPath = Path.Combine(destinationDirectory, "photo.jpg");
            var competitor = CreatePayload(iteration + 9000);
            if (hasCollision)
            {
                File.WriteAllBytes(competitorPath, competitor);
            }

            var mode = FinalizationCrashModes[iteration % FinalizationCrashModes.Length];
            await RunUntilCheckpointAndKillAsync(mode, source, destinationDirectory, checkpoint);

            CollectionAssert.AreEqual(payload, File.ReadAllBytes(source), $"Source changed after crash mode {mode}, iteration {iteration}.");
            if (hasCollision)
            {
                CollectionAssert.AreEqual(competitor, File.ReadAllBytes(competitorPath), $"Existing destination was overwritten in iteration {iteration}.");
            }

            var expectedFinal = Path.Combine(destinationDirectory, hasCollision ? "photo_2.jpg" : "photo.jpg");
            if (mode == "before-finalize")
            {
                Assert.IsFalse(File.Exists(expectedFinal), $"Final file appeared before finalization in iteration {iteration}.");
                Assert.AreEqual(1, Directory.EnumerateFiles(destinationDirectory, ".partial-*", SearchOption.TopDirectoryOnly).Count());
            }
            else
            {
                Assert.IsTrue(File.Exists(expectedFinal), $"Final file was missing after finalization in iteration {iteration}.");
                CollectionAssert.AreEqual(payload, File.ReadAllBytes(expectedFinal));
                Assert.IsFalse(Directory.EnumerateFiles(destinationDirectory, ".partial-*", SearchOption.TopDirectoryOnly).Any());
            }

            var recovery = await new SafeCopyService().CopyAsync(source, destinationDirectory);
            var expectedStatus = mode == "before-finalize" ? CopyStatus.Copied : CopyStatus.SkippedDuplicate;
            Assert.AreEqual(expectedStatus, recovery.Status, $"Recovery failed after {mode}, iteration {iteration}: {recovery.Error}");
            Assert.AreEqual(expectedFinal, recovery.DestinationPath);
            CollectionAssert.AreEqual(payload, File.ReadAllBytes(source));
            CollectionAssert.AreEqual(payload, File.ReadAllBytes(expectedFinal));

            if (hasCollision)
            {
                CollectionAssert.AreEqual(competitor, File.ReadAllBytes(competitorPath));
            }
        }
    }

    [TestMethod]
    public async Task CrashDuringExistingDuplicateDurability_PreservesBothCopiesAndRecovers()
    {
        for (var iteration = 0; iteration < 4; iteration++)
        {
            using var temp = new TempDirectory();
            var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "card")).FullName;
            var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "library")).FullName;
            var source = Path.Combine(sourceDirectory, "photo.jpg");
            var destination = Path.Combine(destinationDirectory, "photo.jpg");
            var checkpoint = Path.Combine(temp.Path, "checkpoint.txt");
            var payload = CreatePayload(iteration + 2000);
            File.WriteAllBytes(source, payload);
            File.WriteAllBytes(destination, payload);

            await RunUntilCheckpointAndKillAsync("duplicate-durability", source, destinationDirectory, checkpoint);

            CollectionAssert.AreEqual(payload, File.ReadAllBytes(source));
            CollectionAssert.AreEqual(payload, File.ReadAllBytes(destination));
            Assert.IsFalse(Directory.EnumerateFiles(destinationDirectory, ".partial-*", SearchOption.TopDirectoryOnly).Any());

            var recovery = await new SafeCopyService().CopyAsync(source, destinationDirectory);
            Assert.AreEqual(CopyStatus.SkippedDuplicate, recovery.Status, recovery.Error);
            Assert.AreEqual(destination, recovery.DestinationPath);
            CollectionAssert.AreEqual(payload, File.ReadAllBytes(source));
            CollectionAssert.AreEqual(payload, File.ReadAllBytes(destination));
        }
    }

    [TestMethod]
    public async Task RepeatedPreFinalizeCrashes_LeaveOnlyUnownedPartialsAndLaterImportStillSucceeds()
    {
        using var temp = new TempDirectory();
        var sourceDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "card")).FullName;
        var destinationDirectory = Directory.CreateDirectory(Path.Combine(temp.Path, "library")).FullName;
        var source = Path.Combine(sourceDirectory, "photo.jpg");
        var payload = CreatePayload(3000);
        File.WriteAllBytes(source, payload);

        const int crashCount = 6;
        for (var iteration = 0; iteration < crashCount; iteration++)
        {
            var checkpoint = Path.Combine(temp.Path, $"checkpoint-{iteration}.txt");
            await RunUntilCheckpointAndKillAsync("before-finalize", source, destinationDirectory, checkpoint);

            CollectionAssert.AreEqual(payload, File.ReadAllBytes(source));
            Assert.IsFalse(File.Exists(Path.Combine(destinationDirectory, "photo.jpg")));
            Assert.AreEqual(
                iteration + 1,
                Directory.EnumerateFiles(destinationDirectory, ".partial-*", SearchOption.TopDirectoryOnly).Count(),
                "A killed process may leave only its own never-finalized partial; existing partials must not be touched.");
        }

        var recovery = await new SafeCopyService().CopyAsync(source, destinationDirectory);
        Assert.AreEqual(CopyStatus.Copied, recovery.Status, recovery.Error);
        Assert.AreEqual(Path.Combine(destinationDirectory, "photo.jpg"), recovery.DestinationPath);
        CollectionAssert.AreEqual(payload, File.ReadAllBytes(source));
        CollectionAssert.AreEqual(payload, File.ReadAllBytes(recovery.DestinationPath!));
        Assert.AreEqual(crashCount, Directory.EnumerateFiles(destinationDirectory, ".partial-*", SearchOption.TopDirectoryOnly).Count());
    }

    private static async Task RunUntilCheckpointAndKillAsync(
        string mode,
        string source,
        string destinationDirectory,
        string checkpoint)
    {
        var workerAssembly = FindWorkerAssembly();
        var startInfo = new ProcessStartInfo("dotnet")
        {
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add(workerAssembly);
        startInfo.ArgumentList.Add(mode);
        startInfo.ArgumentList.Add(source);
        startInfo.ArgumentList.Add(destinationDirectory);
        startInfo.ArgumentList.Add(checkpoint);

        using var process = Process.Start(startInfo) ?? throw new AssertFailedException("Failed to start fault-injection worker.");
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(30);

        while (!File.Exists(checkpoint))
        {
            if (process.HasExited)
            {
                throw new AssertFailedException($"Fault-injection worker exited before checkpoint {mode} with code {process.ExitCode}.");
            }

            if (DateTime.UtcNow >= deadline)
            {
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch (InvalidOperationException)
                {
                }

                throw new AssertFailedException($"Timed out waiting for fault-injection checkpoint {mode}.");
            }

            await Task.Delay(20);
        }

        process.Kill(entireProcessTree: true);
        await process.WaitForExitAsync();
        Assert.IsTrue(process.HasExited);
    }

    private static string FindWorkerAssembly()
    {
        var root = FindRepositoryRoot();
        var configuration = typeof(ProcessCrashSafetyTests).Assembly
            .GetCustomAttribute<AssemblyConfigurationAttribute>()?.Configuration ?? "Debug";
        var binRoot = Path.Combine(root, "tests", "PhotoOrganizer.FaultInjection.Worker", "bin", configuration);
        if (!Directory.Exists(binRoot))
        {
            throw new AssertFailedException($"Fault-injection worker output directory does not exist: {binRoot}");
        }

        var runtimeConfig = Directory
            .EnumerateFiles(binRoot, "PhotoOrganizer.FaultInjection.Worker.runtimeconfig.json", SearchOption.AllDirectories)
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .FirstOrDefault();
        if (runtimeConfig is null)
        {
            throw new AssertFailedException($"Fault-injection worker runtimeconfig was not found under {binRoot}.");
        }

        var workerAssembly = Path.Combine(Path.GetDirectoryName(runtimeConfig)!, "PhotoOrganizer.FaultInjection.Worker.dll");
        if (!File.Exists(workerAssembly))
        {
            throw new AssertFailedException($"Fault-injection worker assembly was not found: {workerAssembly}");
        }

        return workerAssembly;
    }

    private static string FindRepositoryRoot()
    {
        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null)
        {
            if (File.Exists(Path.Combine(current.FullName, "PhotoOrganizer.slnx")))
            {
                return current.FullName;
            }

            current = current.Parent;
        }

        throw new AssertFailedException($"Repository root could not be found from {AppContext.BaseDirectory}.");
    }

    private static byte[] CreatePayload(int seed)
    {
        var bytes = new byte[64 * 1024];
        new Random(seed).NextBytes(bytes);
        return bytes;
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PhotoOrganizer-Crash-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try
            {
                Directory.Delete(Path, recursive: true);
            }
            catch
            {
            }
        }
    }
}
