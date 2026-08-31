using System.Text.RegularExpressions;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

internal static partial class StorageEjectServiceFactory
{
    public static IStorageEjectService Create() => OperatingSystem.IsMacOS()
        ? new MacStorageEjectService()
        : new UnsupportedStorageEjectService();
}

internal sealed partial class MacStorageEjectService : IStorageEjectService
{
    public bool IsSupported => true;

    public StorageEjectResult Eject(MountedVolumeInfo volume)
    {
        var device = GetWholeDiskIdentifier(volume.PhysicalDeviceFingerprint);
        if (device is null)
        {
            return StorageEjectResult.Failed("SDカードの物理デバイスを特定できませんでした。");
        }

        var result = BoundedProcessRunner.Run(
            "/usr/sbin/diskutil",
            ["eject", device],
            TimeSpan.FromSeconds(20));
        if (result is null || result.TimedOut)
        {
            return StorageEjectResult.Failed("取り出し処理が応答しませんでした。Finderで使用中のファイルを閉じて、もう一度試してください。");
        }
        if (result.ExitCode == 0)
        {
            return StorageEjectResult.Succeeded("SDカードを安全に取り出しました。");
        }

        var detail = string.IsNullOrWhiteSpace(result.StandardError)
            ? result.StandardOutput.Trim()
            : result.StandardError.Trim();
        return StorageEjectResult.Failed(string.IsNullOrWhiteSpace(detail)
            ? "SDカードを取り出せませんでした。使用中のアプリを閉じて、もう一度試してください。"
            : $"SDカードを取り出せませんでした: {detail}");
    }

    internal static string? GetWholeDiskIdentifier(string? fingerprint)
    {
        const string prefix = "mac-whole-disk:";
        if (string.IsNullOrWhiteSpace(fingerprint)
            || !fingerprint.StartsWith(prefix, StringComparison.Ordinal)) return null;

        var identifier = fingerprint[prefix.Length..];
        return WholeDiskPattern().IsMatch(identifier) ? identifier : null;
    }

    [GeneratedRegex("^disk[0-9]+$", RegexOptions.CultureInvariant)]
    private static partial Regex WholeDiskPattern();
}
