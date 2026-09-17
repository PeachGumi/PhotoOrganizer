using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class CopyPathSafetyRegressionTests
{
    [TestMethod]
    public async Task SafeCopy_RejectsSourceFileSymlinkBeforeReadingIt()
    {
        using var temp = new TempDirectory();
        var realSource = Path.Combine(temp.Source, "real.jpg");
        var sourceAlias = Path.Combine(temp.Source, "photo.jpg");
        File.WriteAllText(realSource, "camera-bytes");
        CreateFileSymlinkOrSkip(sourceAlias, realSource);

        var result = await new SafeCopyService().CopyAsync(sourceAlias, temp.Destination);

        Assert.AreEqual(CopyStatus.Failed, result.Status);
        Assert.IsFalse(Directory.EnumerateFiles(temp.Destination).Any());
        Assert.AreEqual("camera-bytes", File.ReadAllText(realSource));
    }

    [TestMethod]
    public async Task SafeCopy_DoesNotReportCommittedWhenFinalPathIsReplacedBySymlink()
    {
        using var temp = new TempDirectory();
        EnsureFileSymlinkCapabilityOrSkip();
        var source = Path.Combine(temp.Source, "photo.jpg");
        var finalPath = Path.Combine(temp.Destination, "photo.jpg");
        File.WriteAllText(source, "camera-bytes");

        var result = await new SafeCopyService(new SymlinkFinalizationService(source))
            .CopyAsync(source, temp.Destination);

        Assert.AreEqual(CopyStatus.Failed, result.Status);
        Assert.IsTrue(File.Exists(finalPath));
        Assert.IsTrue(new FileInfo(finalPath).LinkTarget is not null);
        Assert.AreEqual("camera-bytes", File.ReadAllText(source));
    }

    [TestMethod]
    public async Task DestinationLookup_RejectsCandidateReplacedBySymlinkDuringHashing()
    {
        using var temp = new TempDirectory();
        EnsureFileSymlinkCapabilityOrSkip();
        var source = Path.Combine(temp.Source, "photo.jpg");
        var candidate = Path.Combine(temp.Destination, "existing.jpg");
        File.WriteAllText(source, "same-bytes");
        File.WriteAllText(candidate, "same-bytes");

        var result = await new DestinationLibrary(
                hasher: new ReplaceCandidateOnSourceHash(source, candidate))
            .FindVerifiedBackupsAsync([source], temp.Destination);

        Assert.AreEqual(0, result.MatchedSources.Count);
        Assert.IsTrue(result.Errors.Any(error => error.Contains("not direct", StringComparison.Ordinal)));
    }

    [TestMethod]
    public async Task DestinationLookup_RejectsSourceFileSymlink()
    {
        using var temp = new TempDirectory();
        var realSource = Path.Combine(temp.Source, "real.jpg");
        var sourceAlias = Path.Combine(temp.Source, "photo.jpg");
        var candidate = Path.Combine(temp.Destination, "existing.jpg");
        File.WriteAllText(realSource, "same-bytes");
        CreateFileSymlinkOrSkip(sourceAlias, realSource);
        File.WriteAllText(candidate, "same-bytes");

        var result = await new DestinationLibrary()
            .FindVerifiedBackupsAsync([sourceAlias], temp.Destination);

        Assert.AreEqual(0, result.MatchedSources.Count);
        Assert.IsTrue(result.Errors.Any(error => error.Contains("source path is not direct", StringComparison.Ordinal)));
    }

    [TestMethod]
    public async Task Hashing_HonorsPreCanceledTokenBeforeOpeningMissingPath()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            Hashing.Sha256Async(
                Path.Combine(Path.GetTempPath(), $"PhotoOrganizerMissing-{Guid.NewGuid():N}"),
                cancellation.Token));
    }

    [TestMethod]
    public async Task SafeCopy_HonorsPreCanceledTokenBeforeCheckingSource()
    {
        using var temp = new TempDirectory();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new SafeCopyService().CopyAsync(
                Path.Combine(temp.Source, "missing.jpg"),
                temp.Destination,
                cancellation.Token));
    }

    [TestMethod]
    public void Durability_RejectsSymlinkPath()
    {
        using var temp = new TempDirectory();
        var target = Path.Combine(temp.Destination, "target.jpg");
        var alias = Path.Combine(temp.Destination, "alias.jpg");
        File.WriteAllText(target, "bytes");
        CreateFileSymlinkOrSkip(alias, target);

        var result = new PlatformFileDurabilityService().EnsureDurable(alias);

        Assert.IsFalse(result.Success);
        StringAssert.Contains(result.Error ?? string.Empty, "direct");
        Assert.AreEqual("bytes", File.ReadAllText(target));
    }

    [TestMethod]
    public void Finalization_RejectsSymlinkFinalPathWithoutMovingTemporaryFile()
    {
        using var temp = new TempDirectory();
        var target = Path.Combine(temp.Destination, "target.jpg");
        var alias = Path.Combine(temp.Destination, "alias.jpg");
        var temporary = Path.Combine(temp.Destination, ".partial-test");
        File.WriteAllText(target, "existing");
        CreateFileSymlinkOrSkip(alias, target);
        File.WriteAllText(temporary, "new-bytes");

        var result = new PlatformFileDurabilityService()
            .FinalizeNewFile(temporary, alias, DateTime.UtcNow);

        Assert.AreEqual(FinalizeFileStatus.Failed, result.Status);
        Assert.IsFalse(result.FinalPathCreated);
        Assert.IsTrue(File.Exists(temporary));
        Assert.AreEqual("existing", File.ReadAllText(target));
    }

    private static void CreateFileSymlinkOrSkip(string alias, string target)
    {
        try
        {
            File.CreateSymbolicLink(alias, target);
        }
        catch (Exception ex) when (ex is PlatformNotSupportedException
                                   or UnauthorizedAccessException
                                   or IOException)
        {
            Assert.Inconclusive($"File symbolic links are unavailable in this test environment: {ex.Message}");
        }
    }

    private static void EnsureFileSymlinkCapabilityOrSkip()
    {
        var root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerSymlinkCapability-{Guid.NewGuid():N}");
        var target = Path.Combine(root, "target");
        var alias = Path.Combine(root, "alias");

        try
        {
            Directory.CreateDirectory(root);
            File.WriteAllText(target, "capability");
            File.CreateSymbolicLink(alias, target);
        }
        catch (Exception ex) when (ex is PlatformNotSupportedException
                                   or UnauthorizedAccessException
                                   or IOException)
        {
            Assert.Inconclusive($"File symbolic links are unavailable in this test environment: {ex.Message}");
        }
        finally
        {
            try { File.Delete(alias); } catch { }
            try { File.Delete(target); } catch { }
            try { Directory.Delete(root); } catch { }
        }
    }

    private sealed class SymlinkFinalizationService(string sourcePath) : IFileDurabilityService
    {
        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc)
        {
            File.Delete(temporaryPath);
            File.CreateSymbolicLink(finalPath, sourcePath);
            return new FinalizeFileResult(FinalizeFileStatus.Committed, true);
        }

        public DurabilityResult EnsureDurable(string filePath) => new(true);
    }

    private sealed class ReplaceCandidateOnSourceHash(
        string sourcePath,
        string candidatePath) : IFileHasher
    {
        private int _replaced;

        public async Task<string> Sha256Async(
            string path,
            CancellationToken cancellationToken = default)
        {
            if (Interlocked.Exchange(ref _replaced, 1) == 0
                && string.Equals(
                    Path.GetFullPath(path),
                    Path.GetFullPath(sourcePath),
                    OperatingSystem.IsWindows()
                        ? StringComparison.OrdinalIgnoreCase
                        : StringComparison.Ordinal))
            {
                File.Delete(candidatePath);
                File.CreateSymbolicLink(candidatePath, sourcePath);
            }

            return await Hashing.Sha256Async(path, cancellationToken).ConfigureAwait(false);
        }
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerCopyPathSafety-{Guid.NewGuid():N}");
            Source = Directory.CreateDirectory(Path.Combine(Root, "source")).FullName;
            Destination = Directory.CreateDirectory(Path.Combine(Root, "destination")).FullName;
        }

        public string Root { get; }
        public string Source { get; }
        public string Destination { get; }

        public void Dispose()
        {
            try { Directory.Delete(Root, recursive: true); } catch { }
        }
    }
}
