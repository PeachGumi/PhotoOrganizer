namespace PhotoOrganizer.Core;

public enum CopyStatus
{
    Copied,
    SkippedDuplicate,
    Failed
}

public sealed record CopyResult(CopyStatus Status, string? DestinationPath, string? Error = null);

public sealed class SafeCopyService
{
    private const int BufferSize = 4 * 1024 * 1024;
    private readonly IDurableFileCommitter _durableCommitter;

    public SafeCopyService(IDurableFileCommitter? durableCommitter = null)
    {
        _durableCommitter = durableCommitter ?? new PlatformDurableFileCommitter();
    }

    public async Task<CopyResult> CopyAsync(
        string sourcePath,
        string destinationDirectory,
        CancellationToken cancellationToken = default)
    {
        string? temporaryPathForCleanup = null;

        try
        {
            if (!File.Exists(sourcePath))
            {
                return new CopyResult(CopyStatus.Failed, null, "Source file does not exist.");
            }

            var sourceInfo = new FileInfo(sourcePath);
            if (sourceInfo.Length <= 0)
            {
                return new CopyResult(CopyStatus.Failed, null, "Zero-byte supported media cannot be imported safely.");
            }

            if (!PathSafety.TryValidateDirectFilesystemPath(destinationDirectory, out var pathError))
            {
                return new CopyResult(CopyStatus.Failed, null, $"Destination path is not a direct filesystem path: {pathError}");
            }

            Directory.CreateDirectory(destinationDirectory);

            // Re-check after creation so a path alias can never be accepted merely
            // because a leaf component did not exist during the first inspection.
            if (!PathSafety.TryValidateDirectFilesystemPath(destinationDirectory, out pathError))
            {
                return new CopyResult(CopyStatus.Failed, null, $"Destination path became unsafe: {pathError}");
            }

            var sourceSize = sourceInfo.Length;
            var sourceLastWriteUtc = sourceInfo.LastWriteTimeUtc;
            var sourceHashBefore = await Hashing.Sha256Async(sourcePath, cancellationToken).ConfigureAwait(false);

            var resolution = await ResolveDestinationAsync(
                sourcePath,
                destinationDirectory,
                sourceSize,
                sourceHashBefore,
                cancellationToken).ConfigureAwait(false);

            if (resolution.IsDuplicate)
            {
                return new CopyResult(CopyStatus.SkippedDuplicate, resolution.Path);
            }

            var transactionTemporaryPath = Path.Combine(destinationDirectory, $".partial-{Guid.NewGuid():N}");
            temporaryPathForCleanup = transactionTemporaryPath;

            await using (var source = new FileStream(
                             sourcePath,
                             FileMode.Open,
                             FileAccess.Read,
                             FileShare.Read,
                             BufferSize,
                             FileOptions.Asynchronous | FileOptions.SequentialScan))
            await using (var destination = new FileStream(
                             transactionTemporaryPath,
                             FileMode.CreateNew,
                             FileAccess.Write,
                             FileShare.None,
                             BufferSize,
                             FileOptions.Asynchronous | FileOptions.WriteThrough))
            {
                await source.CopyToAsync(destination, BufferSize, cancellationToken).ConfigureAwait(false);
                await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
                destination.Flush(flushToDisk: true);
            }

            var temporaryInfo = new FileInfo(transactionTemporaryPath);
            if (temporaryInfo.Length != sourceSize)
            {
                return new CopyResult(CopyStatus.Failed, null, "Temporary copy size verification failed.");
            }

            var temporaryHash = await Hashing.Sha256Async(transactionTemporaryPath, cancellationToken).ConfigureAwait(false);
            if (!string.Equals(sourceHashBefore, temporaryHash, StringComparison.Ordinal))
            {
                return new CopyResult(CopyStatus.Failed, null, "Temporary copy SHA-256 verification failed.");
            }

            sourceInfo.Refresh();
            if (!sourceInfo.Exists || sourceInfo.Length != sourceSize)
            {
                return new CopyResult(CopyStatus.Failed, null, "Source media changed while copying.");
            }

            var sourceHashAfter = await Hashing.Sha256Async(sourcePath, cancellationToken).ConfigureAwait(false);
            if (!string.Equals(sourceHashBefore, sourceHashAfter, StringComparison.Ordinal))
            {
                return new CopyResult(CopyStatus.Failed, null, "Source media changed while copying.");
            }

            if (!PathSafety.TryValidateDirectFilesystemPath(destinationDirectory, out pathError))
            {
                return new CopyResult(CopyStatus.Failed, null, $"Destination path changed before finalization: {pathError}");
            }

            // Apply metadata before the durable move so no application metadata write
            // occurs after the platform-specific commit barrier.
            File.SetLastWriteTimeUtc(transactionTemporaryPath, sourceLastWriteUtc);

            var finalPath = resolution.Path;
            while (true)
            {
                var commitStatus = _durableCommitter.CommitNoReplace(transactionTemporaryPath, finalPath);
                if (commitStatus == DurableCommitStatus.Committed)
                {
                    temporaryPathForCleanup = null;
                    break;
                }

                // Another writer won the destination name between resolution and
                // finalization. Never replace it; re-resolve against its real bytes.
                resolution = await ResolveDestinationAsync(
                    sourcePath,
                    destinationDirectory,
                    sourceSize,
                    sourceHashBefore,
                    cancellationToken).ConfigureAwait(false);

                if (resolution.IsDuplicate)
                {
                    return new CopyResult(CopyStatus.SkippedDuplicate, resolution.Path);
                }

                finalPath = resolution.Path;
            }

            var finalInfo = new FileInfo(finalPath);
            if (finalInfo.Length != sourceSize)
            {
                return new CopyResult(CopyStatus.Failed, finalPath, "Final copy size verification failed.");
            }

            var finalHash = await Hashing.Sha256Async(finalPath, cancellationToken).ConfigureAwait(false);
            if (!string.Equals(sourceHashBefore, finalHash, StringComparison.Ordinal))
            {
                return new CopyResult(CopyStatus.Failed, finalPath, "Final copy SHA-256 verification failed.");
            }

            return new CopyResult(CopyStatus.Copied, finalPath);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return new CopyResult(CopyStatus.Failed, null, ex.Message);
        }
        finally
        {
            // The only file this service deletes is its own never-finalized temporary file.
            if (temporaryPathForCleanup is not null && File.Exists(temporaryPathForCleanup))
            {
                try
                {
                    File.Delete(temporaryPathForCleanup);
                }
                catch
                {
                    // A leaked .partial file is preferable to deleting or overwriting user data.
                }
            }
        }
    }

    private static async Task<(string Path, bool IsDuplicate)> ResolveDestinationAsync(
        string sourcePath,
        string destinationDirectory,
        long sourceSize,
        string sourceHash,
        CancellationToken cancellationToken)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var extension = Path.GetExtension(sourcePath);

        for (var suffix = 1; ; suffix++)
        {
            var candidateName = suffix == 1 ? $"{fileName}{extension}" : $"{fileName}_{suffix}{extension}";
            var candidate = Path.Combine(destinationDirectory, candidateName);
            if (!File.Exists(candidate))
            {
                return (candidate, false);
            }

            var candidateInfo = new FileInfo(candidate);
            if (candidateInfo.Length != sourceSize)
            {
                continue;
            }

            var candidateHash = await Hashing.Sha256Async(candidate, cancellationToken).ConfigureAwait(false);
            if (string.Equals(candidateHash, sourceHash, StringComparison.Ordinal))
            {
                return (candidate, true);
            }
        }
    }
}
