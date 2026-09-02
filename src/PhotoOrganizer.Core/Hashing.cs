using System.Buffers;
using System.Security.Cryptography;

namespace PhotoOrganizer.Core;

public static class Hashing
{
    private const int BufferSize = 8 * 1024 * 1024;

    public static async Task<string> Sha256Async(string path, CancellationToken cancellationToken = default)
    {
        await using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 1,
            options: FileOptions.Asynchronous | FileOptions.SequentialScan);

        byte[]? currentBuffer = null;
        byte[]? nextBuffer = null;

        try
        {
            currentBuffer = ArrayPool<byte>.Shared.Rent(BufferSize);
            nextBuffer = ArrayPool<byte>.Shared.Rent(BufferSize);
            var current = currentBuffer;
            var next = nextBuffer;
            using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);

            var currentCount = await stream
                .ReadAsync(current.AsMemory(0, BufferSize), cancellationToken)
                .ConfigureAwait(false);

            while (currentCount > 0)
            {
                // Read the next chunk while the CPU hashes the current immutable
                // buffer. This keeps one sequential read in flight without allowing
                // unbounded read-ahead or changing the resulting SHA-256.
                var readTask = stream
                    .ReadAsync(next.AsMemory(0, BufferSize), cancellationToken)
                    .AsTask();
                hash.AppendData(current.AsSpan(0, currentCount));
                currentCount = await readTask.ConfigureAwait(false);
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
}
