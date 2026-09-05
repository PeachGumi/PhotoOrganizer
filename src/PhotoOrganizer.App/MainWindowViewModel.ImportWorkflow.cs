using System.Diagnostics;
using Avalonia.Media;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed partial class MainWindowViewModel
{
    public async Task StartImportAsync()
    {
        ImportScanSession session;
        string destination;
        string eventName;
        CancellationTokenSource importCancellation;

        lock (_importStartGate)
        {
            if (_disposed
                || IsBusy
                || IsSafeToReuseCurrentCard
                || DestinationNeedsReselection
                || _scanSession is null
                || string.IsNullOrWhiteSpace(DestinationPath)
                || string.IsNullOrWhiteSpace(EventName)) return;

            session = _scanSession;
            destination = DestinationPath;
            eventName = EventName.Trim();
            importCancellation = new CancellationTokenSource();
            _isProcessing = true;
            _importCancellation = importCancellation;
        }

        try
        {
            ClearCompletion();
            IsSafeToReuseCurrentCard = false;
            RaiseCommandState();
            SetNotVerified("コピー処理中です。最終検証が完了するまでSDカードを再利用しないでください。");
            SetProgressState("コピー準備中…", indeterminate: true);
            AppendLog("取り込み開始。コピー完了だけではSDカード再利用可能とは判定しません。");

            await ValidateDestinationAsync().ConfigureAwait(true);
            importCancellation.Token.ThrowIfCancellationRequested();
            if (!IsCurrentImportInput(session, destination, eventName))
            {
                HandleImportInputChanged();
                return;
            }

            if (HasInputValidationError)
            {
                SetProgressState("入力内容を確認してください");
                RaiseCommandState();
                return;
            }

            var progress = new Progress<ImportProgress>(update =>
            {
                if (_disposed) return;

                try
                {
                    ApplyImportProgress(update);
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

            var result = await Task.Run(
                () => _coordinator.ImportAsync(
                    session,
                    destination,
                    eventName,
                    progress,
                    importCancellation.Token),
                importCancellation.Token).ConfigureAwait(true);
            if (_disposed) return;

            importCancellation.Token.ThrowIfCancellationRequested();
            if (!IsCurrentImportInput(session, destination, eventName))
            {
                HandleImportInputChanged();
                return;
            }

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
                IsSafeToReuseCurrentCard = true;
                SetCompletion(result.Summary, verified);
                SetProgressState("取り込み・検証完了", verified, Math.Max(1, verified));
                AppendLog($"最終確認完了: {verified} 件を実ファイルとSHA-256照合し、保存先への永続化を確認しました。SDカードを再利用できます。");
            }
            else
            {
                IsSafeToReuseCurrentCard = false;
                SetBlocked(GetImportFailureMessage(result));
                SetProgressState(result.Summary.Failed > 0 ? "一部の取り込みに失敗" : "安全確認を完了できませんでした");
                AppendLog($"最終確認失敗: {result.Message}");
            }
        }
        catch (OperationCanceledException)
        {
            if (_disposed) return;

            IsSafeToReuseCurrentCard = false;
            SetBlocked("取り込みをキャンセルしました。最終検証が完了していないため、SDカードを再利用しないでください。");
            SetProgressState("取り込みキャンセル");
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
                IsSafeToReuseCurrentCard = false;
                ReportOperationFailure("取り込み", exception);
            }
        }
        finally
        {
            lock (_importStartGate)
            {
                if (ReferenceEquals(_importCancellation, importCancellation))
                {
                    _importCancellation = null;
                    _isProcessing = false;
                }
            }

            importCancellation.Dispose();
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

    private bool IsCurrentImportInput(
        ImportScanSession session,
        string destination,
        string eventName)
    {
        if (_disposed
            || !ReferenceEquals(session, _scanSession)
            || !string.Equals(destination, DestinationPath, StringComparison.Ordinal)
            || !string.Equals(eventName, EventName.Trim(), StringComparison.Ordinal)
            || !string.Equals(session.CardRoot, SelectedSdContextPath, _storageSessions.PathComparison)
            || !Directory.Exists(session.CardRoot))
        {
            return false;
        }

        try
        {
            return _storageSessions.Matches(session.SourceIdentity, session.CardRoot);
        }
        catch
        {
            return false;
        }
    }

    private void HandleImportInputChanged()
    {
        if (_disposed) return;

        IsSafeToReuseCurrentCard = false;
        ClearCompletion();
        SetBlocked("取り込み開始後にSDカード、保存先、イベント名が変更されたか、SDカードが取り外されました。入力を確認して取り込みをやり直してください。");
        SetProgressState("取り込み入力が変更されました");
        AppendLog("取り込み入力またはSDカードの接続状態が変化したため、安全確認を未完了に戻しました。");
    }

    private static string GetImportFailureMessage(ImportRunResult result)
    {
        if (result.Message.Contains("cancel", StringComparison.OrdinalIgnoreCase))
        {
            return "取り込みがキャンセルされました。最終検証が完了していないため、SDカードを再利用しないでください。";
        }

        if (result.Summary.Failed > 0)
        {
            return "一部の写真・動画を安全にコピーできませんでした。SDカードは再利用せず、保存先の空き容量と接続状態を確認してからやり直してください。";
        }

        if (result.Verification is not null)
        {
            return "保存先コピーの最終検証を完了できませんでした。SDカードは再利用せず、カードと保存先の接続状態を確認してから再度取り込んでください。";
        }

        return "取り込み前または取り込み中の安全確認を完了できませんでした。SDカードと保存先の接続状態を確認し、SDカードを再スキャンしてからやり直してください。";
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
