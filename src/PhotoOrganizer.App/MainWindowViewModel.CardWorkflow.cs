using System.Diagnostics;
using Avalonia.Threading;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed partial class MainWindowViewModel
{
    private bool _drainingPendingCards;

    public async Task InitializeAsync()
    {
        if (_disposed) return;

        try
        {
            var candidates = await Task.Run(() => _cardRoots.GetCandidateRoots()).ConfigureAwait(true);

            foreach (var candidate in candidates)
            {
                if (_disposed || IsBusy || _scanSession is not null) return;

                var result = await ScanCardAsync(candidate, autoDetected: true).ConfigureAwait(true);
                if (_disposed || result is null || result.IsReady) return;
                if (!result.IsNoSupportedMedia) return;
            }
        }
        catch (Exception exception)
        {
            ReportOperationFailure("起動時のSDカード検出", exception);
        }
    }

    public async Task<ScanSessionResult?> ScanCardAsync(string path, bool autoDetected = false)
    {
        if (_disposed || string.IsNullOrWhiteSpace(path)) return null;
        if (IsBusy)
        {
            QueueCard(path);
            return null;
        }

        var scanCancellation = new CancellationTokenSource();
        var continuePendingAfterScan = false;
        _scanCancellation = scanCancellation;
        _isScanning = true;
        RaiseCommandState();
        ClearCompletion();
        ClearScanSession();
        SelectedSdContextPath = path;
        SetNotVerified("SDカードをスキャン中です。再利用しないでください。");
        SetProgressState("SDカードをスキャン中…", indeterminate: true);
        AppendLog($"スキャン開始: {path}");

        try
        {
            RebuildCoordinator();
            var result = await Task.Run(
                () => _coordinator.ScanCard(path, scanCancellation.Token),
                scanCancellation.Token).ConfigureAwait(true);
            if (_disposed) return null;

            if (!result.IsReady)
            {
                if (autoDetected && result.IsNoSupportedMedia)
                {
                    continuePendingAfterScan = true;
                    ClearScanSession();
                    AppendLog($"自動選択スキップ（対象メディアなし）: {path}");
                    SetProgressState("待機中");
                    return result;
                }

                var userMessage = GetScanFailureMessage(result);
                SetBlocked(userMessage);
                SetProgressState("スキャン失敗");
                AppendLog($"スキャン失敗: {userMessage} / 技術情報: {result.Message}");
                foreach (var error in result.Errors.Take(5)) AppendLog($"  {error}");
                return result;
            }

            _scanSession = result.Session;
            OnPropertyChanged(nameof(ShowMediaSummary));
            SelectedSdPath = result.Session!.CardRoot;
            SelectedSdContextPath = result.Session.CardRoot;
            RemoveQueuedCard(result.Session.CardRoot);
            UpdateCounts(result.Session.Files);
            RefreshLightweightInputValidation();
            await ValidateDestinationAsync().ConfigureAwait(true);
            SetNotVerified("完全スキャン済み。取り込み後の再スキャン、保存先実ファイルのサイズ・SHA-256照合、永続媒体への同期が完了するまでSDカードを再利用しないでください。");
            SetProgressState("取り込み準備完了");
            AppendLog($"スキャン完了: {result.Session.Files.Count} 件 / {CountLabel}");
            RequestShowWindow?.Invoke();
            return result;
        }
        catch (OperationCanceledException)
        {
            if (_disposed) return null;

            SetNotVerified("SDカードのスキャンをキャンセルしました。再スキャンと最終検証が完了するまで再利用しないでください。");
            SetProgressState("スキャンキャンセル");
            AppendLog("SDカードのスキャンをキャンセルしました。安全確認は未検証のままです。");
            return null;
        }
        catch (Exception exception)
        {
            if (_disposed)
            {
                Trace.WriteLine($"Photo Organizer card scan failed after disposal: {exception}");
                return null;
            }

            SetBlocked("SDカードのスキャン中に予期しないエラーが発生しました。カードを挿し直して再選択してください。");
            SetProgressState("スキャン失敗");
            AppendLog($"スキャン失敗（予期しないエラー）: {exception.Message}");
            return null;
        }
        finally
        {
            if (ReferenceEquals(_scanCancellation, scanCancellation))
            {
                _scanCancellation = null;
                scanCancellation.Dispose();
            }

            _isScanning = false;
            if (!_disposed)
            {
                RaiseCommandState();
                if (continuePendingAfterScan && !_drainingPendingCards)
                {
                    await ScanNextPendingIfPossibleAsync().ConfigureAwait(true);
                }
            }
        }
    }

    private static string GetScanFailureMessage(ScanSessionResult result) => result.FailureReason switch
    {
        ScanFailureReason.InvalidCardRoot =>
            "カメラで使用したSDカード全体を確認できませんでした。SDカード、またはカード内のフォルダを選択し直してください。",
        ScanFailureReason.MissingMountSessionIdentity =>
            "SDカードの接続状態を安全に確認できませんでした。カードを挿し直して再選択してください。",
        ScanFailureReason.MissingPhysicalDeviceIdentity =>
            "SDカードの物理デバイス情報を確認できないため、安全に処理できません。カードを挿し直して再選択してください。",
        ScanFailureReason.StorageChanged =>
            "スキャン中にSDカードの接続状態が変化しました。カードを挿し直して再スキャンしてください。",
        ScanFailureReason.IncompleteScan =>
            "SDカード全体を読み取れませんでした。カードやカードリーダーの接続を確認して再スキャンしてください。",
        ScanFailureReason.NoSupportedMedia =>
            "取り込み対象の写真・動画が見つかりませんでした。対象形式はJPG/JPEG・設定済みRAW・MOV/MP4です。",
        _ =>
            "SDカードを安全に確認できませんでした。カードを挿し直して再選択してください。"
    };

    private void OnVolumeMounted(string root) =>
        PostUiTask("カメラカード検出", () => HandleVolumeMountedAsync(root));

    private async Task HandleVolumeMountedAsync(string root)
    {
        if (_disposed) return;

        string? candidate = null;
        for (var attempt = 0; attempt < 5 && candidate is null; attempt++)
        {
            if (_disposed) return;
            candidate = await Task.Run(() => _cardRoots.Resolve(root)).ConfigureAwait(true);
            if (candidate is null) await Task.Delay(300).ConfigureAwait(true);
        }
        if (_disposed || candidate is null) return;

        AppendLog($"カメラカード検出: {candidate}");
        if (_scanSession is not null || IsBusy)
        {
            QueueCard(candidate);
            RequestShowWindow?.Invoke();
            return;
        }

        await ScanCardAsync(candidate, autoDetected: true).ConfigureAwait(true);
    }

    private void OnVolumeRemoved(string root) =>
        PostUiTask("ストレージ取り外し処理", () => HandleVolumeRemovedAsync(root));

    private async Task HandleVolumeRemovedAsync(string root)
    {
        if (_disposed) return;

        RemoveQueuedCard(root, volumeRemoval: true);
        var selectedRoot = _scanSession?.CardRoot ?? SelectedSdContextPath;
        var selectedRemoved = !string.IsNullOrWhiteSpace(selectedRoot)
            && PathSafety.IsSameOrDescendant(selectedRoot, root, _storageSessions.PathComparison);
        var destinationRemoved = !string.IsNullOrWhiteSpace(DestinationPath)
            && PathSafety.IsSameOrDescendant(DestinationPath, root, _storageSessions.PathComparison);

        if (selectedRemoved)
        {
            ClearScanSession();
            SetBlocked("処理対象のSDカードが取り外されました。必要な確認が完了していないため、処理結果を再確認してください。");
            SetProgressState("SDカードが取り外されました");
            AppendLog("SDカード取り外し: スキャン結果と安全確認状態をリセットしました。");
        }

        if (destinationRemoved)
        {
            IsSafeToReuseCurrentCard = false;
            DestinationNeedsReselection = true;
            SetBlocked("保存先ボリュームが取り外されました。保存先を再確認して取り込み・検証をやり直してください。");
            if (!IsBusy) SetProgressState("保存先を再確認してください");
            AppendLog("保存先ボリューム取り外し: 安全確認状態をリセットしました。");
        }

        if ((selectedRemoved || destinationRemoved) && _isProcessing)
        {
            AppendLog("警告: 処理中に使用ストレージが取り外されました。最終判定はfail-closedになります。");
        }

        if (selectedRemoved && !IsBusy)
        {
            await ScanNextPendingIfPossibleAsync().ConfigureAwait(true);
        }
    }

    private void QueueCard(string path)
    {
        var normalized = PathSafety.Normalize(path);
        if (_scanSession is not null
            && string.Equals(_scanSession.CardRoot, normalized, _storageSessions.PathComparison)) return;
        if (_pendingCards.Any(item => string.Equals(item, normalized, _storageSessions.PathComparison))) return;
        _pendingCards.Add(normalized);
        RaisePendingState();
        AppendLog($"SDカード待機: {normalized}");
    }

    private void RemoveQueuedCard(string path, bool volumeRemoval = false)
    {
        var normalized = PathSafety.Normalize(path);
        _pendingCards.RemoveAll(candidate => volumeRemoval
            ? PathSafety.IsSameOrDescendant(candidate, normalized, _storageSessions.PathComparison)
            : string.Equals(candidate, normalized, _storageSessions.PathComparison));
        RaisePendingState();
    }

    private async Task ScanNextPendingIfPossibleAsync()
    {
        if (_drainingPendingCards || _disposed || IsBusy || _scanSession is not null) return;

        _drainingPendingCards = true;
        try
        {
            while (_pendingCards.Count > 0)
            {
                if (_disposed || IsBusy || _scanSession is not null) return;

                var next = _pendingCards[0];
                _pendingCards.RemoveAt(0);
                RaisePendingState();
                var exists = await Task.Run(() => Directory.Exists(next)).ConfigureAwait(true);
                if (_disposed) return;
                if (!exists) continue;

                var result = await ScanCardAsync(next, autoDetected: true).ConfigureAwait(true);
                if (_scanSession is not null) return;
                if (result is null || !result.IsNoSupportedMedia) return;
            }
        }
        finally
        {
            _drainingPendingCards = false;
        }
    }

    private void PostUiTask(string operation, Func<Task> taskFactory)
    {
        if (_disposed) return;

        try
        {
            Dispatcher.UIThread.Post(() => _ = ObserveUiTaskAsync(operation, taskFactory));
        }
        catch (Exception exception)
        {
            Trace.WriteLine($"Photo Organizer could not dispatch {operation}: {exception}");
        }
    }

    private async Task ObserveUiTaskAsync(string operation, Func<Task> taskFactory)
    {
        try
        {
            await taskFactory().ConfigureAwait(true);
        }
        catch (OperationCanceledException) when (_disposed)
        {
        }
        catch (Exception exception)
        {
            ReportOperationFailure(operation, exception);
        }
    }
}
