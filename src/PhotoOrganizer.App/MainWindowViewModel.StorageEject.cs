using Avalonia.Media;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed partial class MainWindowViewModel
{
    public bool EjectSupported => _storageEjectService.IsSupported;
    public bool CanEjectSelectedSd => EjectSupported
        && !IsBusy
        && IsSafeToReuseCurrentCard
        && _scanSession is not null;

    public async Task EjectSelectedSdAsync()
    {
        if (_disposed || !CanEjectSelectedSd || _scanSession is null) return;

        var session = _scanSession;
        var volume = ResolveVerifiedEjectVolume(session);
        if (volume is null)
        {
            IsSafeToReuseCurrentCard = false;
            SetBlocked("SDカードの物理デバイス情報が変化したため、安全に取り出せません。カードを再スキャンしてください。");
            SetProgressState("SDカードを取り出せませんでした");
            return;
        }

        var result = await Task.Run(() => _storageEjectService.Eject(volume)).ConfigureAwait(true);
        if (_disposed) return;

        if (!result.Success)
        {
            SafetyHeadline = "SDカードを取り出せませんでした";
            SafetyDetail = result.Message;
            SafetyBrush = Brushes.DarkOrange;
            ProgressLabel = "SDカードを取り出せませんでした";
            AppendLog($"SDカード取り出し失敗: {result.Message}");
            return;
        }

        AppendLog($"SDカード取り出し完了: {volume.RootPath}");
        ClearScanSession();
        SetProgressState("SDカードを安全に取り出しました");
    }

    private MountedVolumeInfo? ResolveVerifiedEjectVolume(ImportScanSession session)
    {
        var volume = _volumeProvider.ResolveVolumeForPath(session.CardRoot);
        if (volume is null || volume.IsSystem || !volume.IsRemovable) return null;
        if (!string.Equals(volume.Fingerprint, session.SourceIdentity.Fingerprint, StringComparison.Ordinal)) return null;
        if (!string.Equals(
                volume.PhysicalDeviceFingerprint,
                session.SourceIdentity.PhysicalDeviceFingerprint,
                StringComparison.Ordinal)) return null;
        return volume;
    }
}
