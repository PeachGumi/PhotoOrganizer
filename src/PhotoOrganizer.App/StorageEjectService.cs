using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed record StorageEjectResult(bool Success, string Message)
{
    public static StorageEjectResult Succeeded(string message) => new(true, message);
    public static StorageEjectResult Failed(string message) => new(false, message);
}

public interface IStorageEjectService
{
    bool IsSupported { get; }
    StorageEjectResult Eject(MountedVolumeInfo volume);
}

internal sealed class UnsupportedStorageEjectService : IStorageEjectService
{
    public bool IsSupported => false;

    public StorageEjectResult Eject(MountedVolumeInfo volume) =>
        StorageEjectResult.Failed("この環境ではSDカードの取り出しを利用できません。");
}
