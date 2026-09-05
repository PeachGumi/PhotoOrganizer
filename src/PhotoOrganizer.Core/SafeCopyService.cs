using System.Buffers;
using System.Security.Cryptography;

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
    private const int BufferSize = 8 * 1024 * 1024;
    private readonly IFileDurabilityService _durability;

    public SafeCopyService(IFileDurabilityService? durability = null)
    {
        _durability = durability ?? new PlatformFileDurabilityService();
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
            var transactionTemporaryPath = Path.Combine(destinationDirectory, $".partial-{Guid.NewGuid():N}");
            string sourceHashDuringCopy;

            await using (var source = new FileStream(
                             sourcePath,
                             FileMode.Open,
                             FileAccess.Read,
                             FileShare.Read,
                             bufferSize: 1,
                             FileOptions.Asynchronous | FileOptions.SequentialScan))
            await using (var destination = new FileStream(
                             transactionTemporaryPath,
                             FileMode.CreateNew,
                             FileAccess.Write,
                             FileShare.None,
                             bufferSize: 1,
                             FileOptions.Asynchronous))
            {
                // Cleanup ownership starts only after CreateNew succeeds. If another
                // process already owns the randomly selected name, its file must never
                // be deleted by this transaction's finally block.
                temporaryPathForCleanup = transactionTemporaryPath;
                sourceHashDuringCopy = await CopyAndHashAsync(
                    source,
                    destination,
                    cancellationToken).ConfigureAwait(false);
                destination.Flush(flushToDisk: true);
            }

            var temporaryInfo = new FileInfo(transactionTemporaryPath);
            if (temporaryInfo.Length != sourceSize)
            {
                return new CopyResult(CopyStatus.Failed, null, "Temporary copy size verification failed.");
            }

            var temporaryHash = await Hashing
                .Sha256Async(transactionTemporaryPath, cancellationToken)
                .ConfigureAwait(false);
            if (!string.Equals(sourceHashDuringCopy, temporaryHash, StringComparison.Ordinal))
            {
                return new CopyResult(CopyStatus.Failed, null, "Temporary copy SHA-256 verification failed.");
            }

            sourceInfo.Refresh();
            if (!sourceInfo.Exists || sourceInfo.Length != sourceSize)
            {
                return new CopyResult(CopyStatus.Failed, null, "Source media changed while copying.");
            }

            var sourceHashAfter = await Hashing
                .Sha256Async(sourcePath, cancellationToken)
                .ConfigureAwait(false);
            if (!string.Equals(sourceHashDuringCopy, sourceHashAfter, StringComparison.Ordinal))
            {
                return new CopyResult(CopyStatus.Failed, null, "Source media changed while copying.");
            }

            if (!PathSafety.TryValidateDirectFilesystemPath(destinationDirectory, out pathError))
            {
                return new CopyResult(CopyStatus.Failed, null, $"Destination path changed before finalization: {pathError}");
            }

            // Resolve collisions only after the copy hash and the fresh source
            // hash have been established. A race can still claim this name before
            // finalization; the durable no-clobber loop below handles that case.
            var resolution = await ResolveDestinationAsync(
                sourcePath,
                destinationDirectory,
                sourceSize,
                sourceHashDuringCopy,
                cancellationToken).ConfigureAwait(false);

            if (resolution.IsDuplicate)
            {
                if (!TryDeleteTemporaryBeforeDuplicate(transactionTemporaryPath, out var cleanupError))
                {
                    return new CopyResult(CopyStatus.Failed, null, cleanupError);
                }

                temporaryPathForCleanup = null;
                return await VerifyExistingDuplicateAsync(resolution.Path, sourceSize, sourceHashDuringCopy, cancellationToken)
                    .ConfigureAwait(false);
            }

            var finalPath = resolution.Path;
            while (true)
            {
                var finalization = _durability.FinalizeNewFile(transactionTemporaryPath, finalPath, sourceLastWriteUtc);
                if (finalization.FinalPathCreated)
                {
                    // Once the temporary file has become a finalized user-library file,
                    // cleanup must never delete it even when durability later fails.
                    temporaryPathForCleanup = null;
                }

                if (finalization.Status == FinalizeFileStatus.Committed)
                {
                    break;
                }

                if (finalization.Status == FinalizeFileStatus.Failed)
                {
                    return new CopyResult(
                        CopyStatus.Failed,
                        finalization.FinalPathCreated ? finalPath : null,
                        finalization.Error ?? "Durable destination finalization failed.");
                }

                resolution = await ResolveDestinationAsync(
                    sourcePath,
                    destinationDirectory,
                    sourceSize,
                    sourceHashDuringCopy,
                    cancellationToken).ConfigureAwait(false);

                if (resolution.IsDuplicate)
                {
                    if (!TryDeleteTemporaryBeforeDuplicate(transactionTemporaryPath, out var cleanupError))
                    {
                        return new CopyResult(CopyStatus.Failed, null, cleanupError);
                    }

                    temporaryPathForCleanup = null;
                    return await VerifyExistingDuplicateAsync(resolution.Path, sourceSize, sourceHashDuringCopy, cancellationToken)
                        .ConfigureAwait(false);
                }

                finalPath = resolution.Path;
            }

            var finalInfo = new FileInfo(finalPath);
            if (finalInfo.Length != sourceSize)
            {
                return new CopyResult(CopyStatus.Failed, finalPath, "Final copy size verification failed.");
            }

            var finalHash = await Hashing
                .Sha256Async(finalPath, cancellationToken)
                .ConfigureAwait(false);
            if (!string.Equals(sourceHashDuringCopy, finalHash, StringComparison.Ordinal))
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

    private static async Task<string> CopyAndHashAsync(
        FileStream source,
        FileStream destination,
        CancellationToken cancellationToken)
    {
        byte[]? currentBuffer = null;
        byte[]? nextBuffer = null;

        try
        {
            currentBuffer = ArrayPool<byte>.Shared.Rent(BufferSize);
            nextBuffer = ArrayPool<byte>.Shared.Rent(BufferSize);

            using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            var current = currentBuffer;
            var next = nextBuffer;
            var currentCount = await source
                .ReadAsync(current.AsMemory(0, BufferSize), cancellationToken)
                .ConfigureAwait(false);

            while (currentCount > 0)
            {
                // Keep the two pooled buffers independent so the next source read,
                // current destination write, and current CPU hash can overlap without
                // mutating bytes in flight. The hash is used only after both I/O tasks
                // succeed.
                var writeTask = destination
                    .WriteAsync(current.AsMemory(0, currentCount), cancellationToken)
                    .AsTask();
                var readTask = source
                    .ReadAsync(next.AsMemory(0, BufferSize), cancellationToken)
                    .AsTask();

                hash.AppendData(current.AsSpan(0, currentCount));
                await Task.WhenAll(writeTask, readTask).ConfigureAwait(false);

                currentCount = readTask.Result;
                (current, next) = (next, current);
            }

            return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
        }
        finally
        {
            if (currentBuffer is not null) ArrayPool<byte>.Shared.Return(currentBuffer);
            if (nextBuffer is not null) ArrayPool<byte>.Shared.Return(nextBuffer);
        }
    }

    private async Task<CopyResult> VerifyExistingDuplicateAsync(
        string path, long sourceSize, string sourceHash, CancellationToken cancellationToken)
    {
        var durability = _durability.EnsureDurable(path);
        if (!durability.Success)
        {
            return new CopyResult(CopyStatus.Failed, path,
                durability.Error ?? "Existing duplicate could not be committed durably.");
        }

        if (!PathSafety.TryValidateDirectFilesystemPath(path, out var pathError))
        {
            return new CopyResult(CopyStatus.Failed, path, $"Existing duplicate path changed: {pathError}");
        }

        var info = new FileInfo(path);
        if (!info.Exists || info.Length != sourceSize
            || !string.Equals(sourceHash,
                await Hashing.Sha256Async(path, cancellationToken).ConfigureAwait(false),
                StringComparison.Ordinal))
        {
            return new CopyResult(CopyStatus.Failed, path, "Existing duplicate changed during durable verification.");
        }

        return new CopyResult(CopyStatus.SkippedDuplicate, path);
    }

    private static bool TryDeleteTemporaryBeforeDuplicate(string temporaryPath, out string? error)
    {
        error = null;

        try
        {
            // Duplicate durability may block or fail after this point. Release our
            // never-finalized temporary before entering that call so a process crash
            // cannot leave a copy transaction behind for a duplicate-only result.
            File.Delete(temporaryPath);
            return true;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            error = $"Temporary copy cleanup failed before duplicate verification: {ex.Message}";
            return false;
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

            // A directory occupies the same namespace entry as a file. Treat it as
            // an immutable collision and move on instead of retrying the same path.
            if (Directory.Exists(candidate))
            {
                continue;
            }

            if (!File.Exists(candidate))
            {
                return (candidate, false);
            }

            // A file-level symlink/reparse point is an immutable collision, never a
            // byte-identical independent backup of the source.
            if (!PathSafety.TryValidateDirectFilesystemPath(candidate, out _))
            {
                continue;
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
