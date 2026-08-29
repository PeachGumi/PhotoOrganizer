using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Avalonia.Media;
using Avalonia.Threading;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed class MainWindowViewModel : INotifyPropertyChanged, IDisposable
{
    private readonly IStorageVolumeProvider _volumeProvider;
    private readonly StorageSessionTracker _storageSessions;
    private readonly CameraCardRootResolver _cardRoots;
    private readonly AppPreferencesStore _preferencesStore = new();
    private readonly IStartupRegistrationService _startupRegistration = new StartupRegistrationService();
    private readonly List<string> _logLines = [];
    private readonly List<string> _pendingCards = [];

    private MediaClassifier _classifier;
    private ImportCoordinator _coordinator;
    private ImportScanSession? _scanSession;
    private CancellationTokenSource? _scanCancellation;
    private CancellationTokenSource? _importCancellation;
    private bool _isScanning;
    private bool _isProcessing;
    private bool _changingAutoStart;
    private bool _disposed;

    private string _destinationPath;
    private string _eventName = string.Empty;
    private string _selectedSdPath = string.Empty;
    private string _rawExtensionsText;
    private string _countLabel = "RAW:0 / JPG:0 / MP4:0";
    private string _progressLabel = "待機中";
    private string _logText = string.Empty;
    private string _safetyHeadline = "未検証 — 最終確認が完了するまでSDカードを再利用しないでください";
    private string _safetyDetail = "判定対象: JPG/JPEG・設定済みRAW・MOV/MP4。その他の形式は取り込み・判定対象外です。";
    private IBrush _safetyBrush = Brushes.DarkOrange;
    private bool _autoStart;

    public MainWindowViewModel(
        IStorageVolumeProvider volumeProvider,
        StorageSessionTracker storageSessions,
        CameraCardRootResolver cardRoots)
    {
        _volumeProvider = volumeProvider;
        _storageSessions = storageSessions;
        _cardRoots = cardRoots;

        var preferences = _preferencesStore.Load();
        _destinationPath = preferences.DestinationPath;
        _rawExtensionsText = string.Join(", ", preferences.RawExtensions);
        _autoStart = _startupRegistration.IsEnabled();
        _classifier = new MediaClassifier(ParseRawExtensions(_rawExtensionsText));
        _coordinator = CreateCoordinator();

        _storageSessions.VolumeMounted += OnVolumeMounted;
        _storageSessions.VolumeRemoved += OnVolumeRemoved;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action? RequestShowWindow;

    public string DestinationPath
    {
        get => _destinationPath;
        set
        {
            if (!SetField(ref _destinationPath, value)) return;
            InvalidateSafety("保存先が変更されました。次の取り込み後に再検証が必要です。");
            SavePreferences();
            RaiseCommandState();
        }
    }

    public string EventName
    {
        get => _eventName;
        set
        {
            if (!SetField(ref _eventName, value)) return;
            RaiseCommandState();
        }
    }

    public string SelectedSdPath
    {
        get => _selectedSdPath;
        private set => SetField(ref _selectedSdPath, value);
    }

    public string RawExtensionsText
    {
        get => _rawExtensionsText;
        set
        {
            if (!SetField(ref _rawExtensionsText, value)) return;
            if (!IsBusy)
            {
                RebuildCoordinator();
                ClearScanSession();
                InvalidateSafety("RAW拡張子設定が変更されました。SDカードを再スキャンしてください。");
            }
            SavePreferences();
            RaiseCommandState();
        }
    }

    public string CountLabel
    {
        get => _countLabel;
        private set => SetField(ref _countLabel, value);
    }

    public string ProgressLabel
    {
        get => _progressLabel;
        private set => SetField(ref _progressLabel, value);
    }

    public string LogText
    {
        get => _logText;
        private set => SetField(ref _logText, value);
    }

    public string SafetyHeadline
    {
        get => _safetyHeadline;
        private set => SetField(ref _safetyHeadline, value);
    }

    public string SafetyDetail
    {
        get => _safetyDetail;
        private set => SetField(ref _safetyDetail, value);
    }

    public IBrush SafetyBrush
    {
        get => _safetyBrush;
        private set => SetField(ref _safetyBrush, value);
    }

    public bool AutoStart
    {
        get => _autoStart;
        set
        {
            if (_changingAutoStart || value == _autoStart) return;

            if (_startupRegistration.SetEnabled(value, out var error))
            {
                SetField(ref _autoStart, value);
                AppendLog(value ? "ログイン時自動起動を有効にしました。" : "ログイン時自動起動を無効にしました。");
                SavePreferences();
                return;
            }

            AppendLog($"ログイン時自動起動の設定に失敗しました: {error}");
            _changingAutoStart = true;
            SetField(ref _autoStart, _startupRegistration.IsEnabled());
            _changingAutoStart = false;
        }
    }

    public bool AutoStartSupported => _startupRegistration.IsSupported;
    public bool IsBusy => _isScanning || _isProcessing;
    public bool IsProcessing => _isProcessing;
    public bool CanImport => !IsBusy
        && _scanSession is not null
        && !string.IsNullOrWhiteSpace(DestinationPath)
        && !string.IsNullOrWhiteSpace(EventName);
    public bool CanCancel => _isScanning || _isProcessing;
    public int PendingSdCount => _pendingCards.Count;
    public string PendingSdText => PendingSdCount == 0
        ? string.Empty
        : $"待機中のSDカード: {PendingSdCount} 枚 — 現在のカードを切り替えず保持しています";

    public async Task InitializeAsync()
    {
        if (_disposed) return;

        try
        {
            // Volume enumeration may invoke diskutil/WMI and must never run on the
            // Avalonia dispatcher. ScanCardAsync performs the same isolation for the
            // actual card scan below.
            var candidates = await Task.Run(() => _cardRoots.GetCandidateRoots()).ConfigureAwait(true);
            if (_disposed || candidates.Count == 0) return;

            await ScanCardAsync(candidates[0], autoDetected: true).ConfigureAwait(true);
            if (_disposed) return;

            foreach (var candidate in candidates.Skip(1)) QueueCard(candidate);
        }
        catch (Exception exception)
        {
            ReportOperationFailure("起動時のSDカード検出", exception);
        }
    }

    public async Task ScanCardAsync(string path, bool autoDetected = false)
    {
        if (_disposed || string.IsNullOrWhiteSpace(path)) return;
        if (IsBusy)
        {
            QueueCard(path);
            return;
        }

        var scanCancellation = new CancellationTokenSource();
        _scanCancellation = scanCancellation;
        _isScanning = true;
        RaiseCommandState();
        ClearScanSession();
        InvalidateSafety("SDカードをスキャン中です。再利用しないでください。");
        ProgressLabel = "SDカードをスキャン中...";
        AppendLog($"スキャン開始: {path}");

        try
        {
            RebuildCoordinator();
            var result = await Task.Run(
                () => _coordinator.ScanCard(path, scanCancellation.Token),
                scanCancellation.Token).ConfigureAwait(true);
            if (_disposed) return;

            if (!result.IsReady)
            {
                if (autoDetected && result.Message.Contains("No supported media", StringComparison.OrdinalIgnoreCase))
                {
                    AppendLog($"自動選択スキップ（対象メディアなし）: {path}");
                    ProgressLabel = "待機中";
                    return;
                }

                SetBlocked(result.Message);
                ProgressLabel = "スキャン失敗";
                AppendLog($"スキャン失敗: {result.Message}");
                foreach (var error in result.Errors.Take(5)) AppendLog($"  {error}");
                return;
            }

            _scanSession = result.Session;
            SelectedSdPath = result.Session!.CardRoot;
            RemoveQueuedCard(result.Session.CardRoot);
            UpdateCounts(result.Session.Files);
            SetNotVerified("完全スキャン済み。取り込み後の再スキャン、保存先実ファイルのサイズ・SHA-256照合、永続媒体への同期が完了するまでSDカードを再利用しないでください。");
            ProgressLabel = "取り込み準備完了";
            AppendLog($"スキャン完了: {result.Session.Files.Count} 件 / {CountLabel}");
            RequestShowWindow?.Invoke();
        }
        catch (OperationCanceledException)
        {
            if (_disposed) return;

            ClearScanSession();
            SetNotVerified("SDカードのスキャンをキャンセルしました。再スキャンと最終検証が完了するまで再利用しないでください。");
            ProgressLabel = "スキャンキャンセル";
            AppendLog("SDカードのスキャンをキャンセルしました。安全確認は未検証のままです。");
        }
        catch (Exception exception)
        {
            if (_disposed)
            {
                Trace.WriteLine($"Photo Organizer card scan failed after disposal: {exception}");
                return;
            }

            ClearScanSession();
            SetBlocked("SDカードのスキャン中に予期しないエラーが発生しました。SDカードを再利用しないでください。");
            ProgressLabel = "スキャン失敗";
            AppendLog($"スキャン失敗（予期しないエラー）: {exception.Message}");
        }
        finally
        {
            if (ReferenceEquals(_scanCancellation, scanCancellation))
            {
                _scanCancellation = null;
                scanCancellation.Dispose();
            }

            _isScanning = false;
            if (!_disposed) RaiseCommandState();
        }
    }

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

    public void SetDestinationFromPicker(string path)
    {
        if (!_disposed && !IsBusy) DestinationPath = path;
    }

    public bool CanCloseWindow()
    {
        if (_disposed || !IsBusy) return true;
        AppendLog("処理中の通常終了を抑止しました。必要ならキャンセルして処理終了後に閉じてください。");
        return false;
    }

    private ImportCoordinator CreateCoordinator() => new(
        _classifier,
        _storageSessions,
        _cardRoots,
        _volumeProvider);

    private void RebuildCoordinator()
    {
        _classifier = new MediaClassifier(ParseRawExtensions(RawExtensionsText));
        _coordinator = CreateCoordinator();
    }

    private static string[] ParseRawExtensions(string text) => text
        .Split([',', ';', ' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .Select(extension => extension.StartsWith('.') ? extension.ToLowerInvariant() : "." + extension.ToLowerInvariant())
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToArray();

    private void UpdateCounts(IEnumerable<string> files)
    {
        var raw = 0;
        var jpg = 0;
        var video = 0;
        foreach (var file in files)
        {
            switch (_classifier.Classify(file))
            {
                case MediaKind.Raw: raw++; break;
                case MediaKind.Jpeg: jpg++; break;
                case MediaKind.Video: video++; break;
            }
        }
        CountLabel = $"RAW:{raw} / JPG:{jpg} / MP4:{video}";
    }

    private void SetBlocked(string message)
    {
        SafetyHeadline = "SDカードを再利用しないでください";
        SafetyDetail = message;
        SafetyBrush = Brushes.IndianRed;
    }

    private void SetNotVerified(string detail)
    {
        SafetyHeadline = "未検証 — 最終確認が完了するまでSDカードを再利用しないでください";
        SafetyDetail = detail;
        SafetyBrush = Brushes.DarkOrange;
    }

    private void InvalidateSafety(string detail) => SetNotVerified(detail);

    private void ClearScanSession()
    {
        _scanSession = null;
        SelectedSdPath = string.Empty;
        CountLabel = "RAW:0 / JPG:0 / MP4:0";
        RaiseCommandState();
    }

    private void OnVolumeMounted(string root)
    {
        PostUiTask("カメラカード検出", () => HandleVolumeMountedAsync(root));
    }

    private async Task HandleVolumeMountedAsync(string root)
    {
        if (_disposed) return;

        string? candidate = null;
        for (var attempt = 0; attempt < 5 && candidate is null; attempt++)
        {
            if (_disposed) return;

            // Resolving a nested path enumerates mounted volumes and can invoke
            // diskutil/WMI. Keep that work off the dispatcher even though this
            // handler itself is resumed on the UI thread.
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

    private void OnVolumeRemoved(string root)
    {
        PostUiTask("ストレージ取り外し処理", () => HandleVolumeRemovedAsync(root));
    }

    private async Task HandleVolumeRemovedAsync(string root)
    {
        if (_disposed) return;

        RemoveQueuedCard(root, volumeRemoval: true);
        var selectedRemoved = _scanSession is not null
            && PathSafety.IsSameOrDescendant(_scanSession.CardRoot, root, _storageSessions.PathComparison);
        var destinationRemoved = !string.IsNullOrWhiteSpace(DestinationPath)
            && PathSafety.IsSameOrDescendant(DestinationPath, root, _storageSessions.PathComparison);

        if (selectedRemoved)
        {
            ClearScanSession();
            SetBlocked("選択中のSDカードが取り外されました。再スキャンと最終検証なしに再利用可能とは判定しません。");
            AppendLog("SDカード取り外し: スキャン結果と安全確認状態をリセットしました。");
        }

        if (destinationRemoved)
        {
            SetBlocked("保存先ボリュームが取り外されました。保存先を再確認して取り込み・検証をやり直してください。");
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
        if (_disposed || IsBusy || _scanSession is not null) return;
        while (_pendingCards.Count > 0)
        {
            if (_disposed || IsBusy || _scanSession is not null) return;

            var next = _pendingCards[0];
            _pendingCards.RemoveAt(0);
            RaisePendingState();
            var exists = await Task.Run(() => Directory.Exists(next)).ConfigureAwait(true);
            if (_disposed) return;
            if (!exists) continue;
            await ScanCardAsync(next, autoDetected: true).ConfigureAwait(true);
            if (_scanSession is not null) return;
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
            // Disposal cancellation is expected and must not be reported as a user
            // failure after the application lifetime has ended.
        }
        catch (Exception exception)
        {
            ReportOperationFailure(operation, exception);
        }
    }

    public void ReportUiFailure(string operation, Exception exception)
    {
        ReportOperationFailure(operation, exception);
    }

    private void ReportOperationFailure(string operation, Exception exception)
    {
        if (_disposed)
        {
            Trace.WriteLine($"Photo Organizer {operation} failed after disposal: {exception}");
            return;
        }

        try
        {
            SetBlocked($"{operation}中に予期しないエラーが発生しました。SDカードを再利用しないでください。");
            ProgressLabel = $"{operation}失敗";
            AppendLog($"{operation}失敗（予期しないエラー）: {exception.Message}");
        }
        catch (Exception reportException)
        {
            // A failing PropertyChanged subscriber must not turn a recoverable
            // operation error into an unhandled dispatcher exception.
            Trace.WriteLine($"Photo Organizer could not report {operation} failure: {reportException}");
        }
    }

    private void AppendLog(string message)
    {
        var timestamp = DateTime.Now.ToString("HH:mm:ss");
        _logLines.Add($"{timestamp} {message}");
        if (_logLines.Count > 1000) _logLines.RemoveRange(0, _logLines.Count - 1000);
        LogText = string.Join(Environment.NewLine, _logLines);
    }

    private void SavePreferences()
    {
        var preferences = new AppPreferences(DestinationPath, ParseRawExtensions(RawExtensionsText), AutoStart);
        if (!_preferencesStore.Save(preferences, out var error) && !string.IsNullOrWhiteSpace(error))
        {
            AppendLog($"設定保存失敗: {error}");
        }
    }

    private void RaiseCommandState()
    {
        OnPropertyChanged(nameof(IsBusy));
        OnPropertyChanged(nameof(IsProcessing));
        OnPropertyChanged(nameof(CanImport));
        OnPropertyChanged(nameof(CanCancel));
    }

    private void RaisePendingState()
    {
        OnPropertyChanged(nameof(PendingSdCount));
        OnPropertyChanged(nameof(PendingSdText));
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    public void Dispose()
    {
        if (_disposed) return;

        _disposed = true;
        _storageSessions.VolumeMounted -= OnVolumeMounted;
        _storageSessions.VolumeRemoved -= OnVolumeRemoved;

        try { _scanCancellation?.Cancel(); }
        catch (Exception exception)
        {
            Trace.WriteLine($"Photo Organizer scan cancellation during disposal failed: {exception}");
        }

        try { _importCancellation?.Cancel(); }
        catch (Exception exception)
        {
            Trace.WriteLine($"Photo Organizer import cancellation during disposal failed: {exception}");
        }
    }
}
