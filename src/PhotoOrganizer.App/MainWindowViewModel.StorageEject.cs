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
        ImportScanSession session;
        MountedVolumeInfo volume;
        lock (_importStartGate)
        {
            if (_disposed || !CanEjectSelectedSd || _scanSession is null) return;

            session = _scanSession;
            var resolvedVolume = ResolveVerifiedEjectVolume(session);
            if (resolvedVolume is null)
            {
                IsSafeToReuseCurrentCard = false;
                SetBlocked("SDカードの物理デバイス情報が変化したため、安全に取り出せません。カードを再スキャンしてください。");
                SetProgressState("SDカードを取り出せませんでした");
                return;
            }

            volume = resolvedVolume;

            _isEjecting = true;
            _ejectingSession = session;
            _ejectRemovalObserved = false;
        }

        var shouldDrainPending = false;
        RaiseCommandState();
        SetProgressState("SDカードを取り出し中…", indeterminate: true);
        AppendLog($"SDカード取り出し開始: {volume.RootPath}");

        try
        {
            var result = await Task.Run(() => _storageEjectService.Eject(volume)).ConfigureAwait(true);
            if (_disposed) return;

            // A removal callback may have cleared this session while the platform
            // eject call was in flight. Never clear a newer selection or publish a
            // success state for that stale operation.
            if (!ReferenceEquals(session, _scanSession))
            {
                AppendLog("SDカード取り出し結果を破棄しました。処理対象がすでに変更されています。");
                return;
            }

            if (!result.Success)
            {
                if (_ejectRemovalObserved)
                {
                    ClearScanSession();
                    SetBlocked("SDカードの取り出し結果を確認できませんでした。カードを再接続して再スキャンしてください。");
                    SetProgressState("SDカードを取り出せませんでした");
                }
                else
                {
                    SafetyHeadline = "SDカードを取り出せませんでした";
                    SafetyDetail = result.Message;
                    SafetyBrush = Brushes.DarkOrange;
                    ProgressLabel = "SDカードを取り出せませんでした";
                }
                AppendLog($"SDカード取り出し失敗: {result.Message}");
                return;
            }

            AppendLog($"SDカード取り出し完了: {volume.RootPath}");
            ClearScanSession();
            SetProgressState("SDカードを安全に取り出しました");
            shouldDrainPending = true;
        }
        catch (Exception exception)
        {
            if (!_disposed)
            {
                IsSafeToReuseCurrentCard = false;
                SetBlocked("SDカードの取り出し中に予期しないエラーが発生しました。カードを再スキャンしてください。");
                SetProgressState("SDカードを取り出せませんでした");
                AppendLog($"SDカード取り出し失敗（予期しないエラー）: {exception.Message}");
            }

            throw;
        }
        finally
        {
            _isEjecting = false;
            _ejectingSession = null;
            _ejectRemovalObserved = false;
            if (!_disposed)
            {
                RaiseCommandState();
                if (shouldDrainPending)
                {
                    try
                    {
                        await ScanNextPendingIfPossibleAsync().ConfigureAwait(true);
                    }
                    catch (Exception exception)
                    {
                        ReportOperationFailure("待機中SDカードのスキャン", exception);
                    }
                }
            }
        }
    }

    private MountedVolumeInfo? ResolveVerifiedEjectVolume(ImportScanSession session)
    {
        StorageSessionIdentity? currentIdentity;
        try
        {
            currentIdentity = _storageSessions.Capture(session.CardRoot);
        }
        catch
        {
            return null;
        }

        if (currentIdentity is null
            || currentIdentity.SessionId != session.SourceIdentity.SessionId
            || !string.Equals(currentIdentity.Fingerprint, session.SourceIdentity.Fingerprint, StringComparison.Ordinal)
            || !string.Equals(
                currentIdentity.PhysicalDeviceFingerprint,
                session.SourceIdentity.PhysicalDeviceFingerprint,
                StringComparison.Ordinal))
        {
            return null;
        }

        var volume = _volumeProvider.ResolveVolumeForPath(session.CardRoot);
        if (volume is null || volume.IsSystem || !volume.IsRemovable) return null;
        if (!string.Equals(volume.RootPath, currentIdentity.RootPath, _storageSessions.PathComparison)) return null;
        if (!string.Equals(volume.Fingerprint, session.SourceIdentity.Fingerprint, StringComparison.Ordinal)) return null;
        if (!string.Equals(
                volume.PhysicalDeviceFingerprint,
                session.SourceIdentity.PhysicalDeviceFingerprint,
                StringComparison.Ordinal)) return null;
        return volume;
    }
}
