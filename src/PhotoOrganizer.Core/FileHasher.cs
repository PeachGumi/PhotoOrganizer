namespace PhotoOrganizer.Core;

public interface IFileHasher
{
    Task<string> Sha256Async(string path, CancellationToken cancellationToken = default);
}

public sealed class Sha256FileHasher : IFileHasher
{
    public Task<string> Sha256Async(string path, CancellationToken cancellationToken = default) =>
        Hashing.Sha256Async(path, cancellationToken);
}
