using System.Diagnostics;
using Avalonia.Media;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed partial class MainWindowViewModel
{
    public async Task StartImportAsync()
    {
        if (_disposed || !CanImport || _scanSession is null) return;

        var session = _scanSession;
        var destination = DestinationPath;
        var eventName = EventName.Trim();
        var importCancellation = new CancellationTokenSource();
        _isProcessing = true;
        _importCancellation = importCancellation;
        RaiseCommandState();
        SetNotVerified("コピー処理中です。最終検証が完了するまでSDカードを再利用しないでください。");
        ProgressLabel = "コピー準備中...";
        AppendLog("取り込み開始。コピー完了だけではSDカード再利用可能とは判定しません。");

        try
        {
            var progress = new Progress<ImportProgress>(update =>
            {
                if (_disposed) return;

                try
                {
                    ProgressLabel = update.Message;
                    if (update.Phase == ImportProgressPhase.Verifying)
                    {
                        SafetyHeadline = "保存先コピーを検証中 — SDカードを再利用しないでください";
                        SafetyBrush = Brushes.DarkOrange;
                    }
                    if (update.Phase == ImportProgressPhase.Rescanning)
                    {
                        AppendLog("コピー処理完了。SDカード全体の再スキャン、実bytes検証、保存先の永続化確認を開始します。");
                    }
                }
                catch (Exception exception)
                {
                    Trace.WriteLine($"Photo Organizer progress update failed: {exception}");
                }
            });

            // ImportAsync performs synchronous scanner and storage work as well as
            // asynchronous copy/verification. Run the complete operation away from
            // the dispatcher so a large card cannot freeze the workflow window.
            var result = await Task.Run(
                () => _coordinator.ImportAsync(
                    session,
                    destination,
                    eventName,
                    progress,
                    importCancellation.Token),
                importCancellation.Token).ConfigureAwait(true);
            if (_disposed) return;

            AppendLog($"処理結果: コピー {result.Summary.Copied} / 既存一致 {result.Summary.SkippedAlreadyBackedUp} / 失敗 {result.Summary.Failed}");
            if (!string.IsNullOrWhiteSpace(result.Summary.BasePath)) AppendLog($"保存先: {result.Summary.BasePath}");
            foreach (var warning in result.Summary.Warnings.Take(5)) AppendLog($"警告: {warning}");
            foreach (var error in result.Summary.Errors.Take(5)) AppendLog($"エラー: {error}");

            if (result.IsSafeToReuse)
            {
                var verified = result.Verification?.Verified ?? 0;
                SafetyHeadline = "保存先コピー検証済み — SDカード再利用可能";
                SafetyDetail = $"対象メディア {verified} 件について、取り込み後のSD再スキャン、保存先実ファイルのサイズ・SHA-256一致、永続媒体への同期（durable commit）を確認済みです。これは指定保存先1か所へのコピー検証であり、二重バックアップ済みという意味ではありません。";
                SafetyBrush = Brushes.ForestGreen;
                ProgressLabel = "保存先コピー・永続化検証済み";
                AppendLog($"最終確認完了: {verified} 件を実ファイルとSHA-256照合し、保存先への永続化を確認しました。SDカードを再利用できます。");
            }
            else
            {
                SetBlocked(result.Message);
                ProgressLabel = result.Summary.Failed > 0 ? "一部失敗" : "要確認";
                AppendLog($"最終確認失敗: {result.Message}");
            }
        }
        catch (OperationCanceledException)
        {
            if (_disposed) return;

            SetBlocked("取り込みをキャンセルしました。最終検証が完了していないため、SDカードを再利用しないでください。");
            ProgressLabel = "取り込みキャンセル";
            AppendLog("取り込みをキャンセルしました。SDカードの安全確認は未完了です。");
        }
        catch (Exception exception)
        {
            if (_disposed)
            {
                Trace.WriteLine($"Photo Organizer import failed after disposal: {exception}");
            }
            else
            {
                ReportOperationFailure("取り込み", exception);
            }
        }
        finally
        {
            if (ReferenceEquals(_importCancellation, importCancellation))
            {
                _importCancellation = null;
                importCancellation.Dispose();
            }

            _isProcessing = false;
            if (!_disposed)
            {
                RaiseCommandState();

                if (_scanSession is null)
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

    public void CancelImport()
    {
        if (_disposed) return;

        if (_isScanning)
        {
            AppendLog("スキャンキャンセル要求: 安全確認は未検証のままです。");
            try
            {
                _scanCancellation?.Cancel();
            }
            catch (Exception exception)
            {
                ReportOperationFailure("スキャンキャンセル", exception);
            }
            return;
        }

        if (!_isProcessing) return;
        AppendLog("キャンセル要求: 最終検証が完了しないためSDカードは再利用不可のままです。");
        try
        {
            _importCancellation?.Cancel();
        }
        catch (Exception exception)
        {
            ReportOperationFailure("取り込みキャンセル", exception);
        }
    }
}
