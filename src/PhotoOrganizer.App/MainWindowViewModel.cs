using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Avalonia.Media;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed partial class MainWindowViewModel : INotifyPropertyChanged, IDisposable
{
    private readonly IStorageVolumeProvider _volumeProvider;
    private readonly StorageSessionTracker _storageSessions;
    private readonly CameraCardRootResolver _cardRoots;
    private readonly IStorageEjectService _storageEjectService;
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
    private bool _isSafeToReuseCurrentCard;
    private bool _destinationNeedsReselection;
    private bool _disposed;

    private string _destinationPath;
    private string _eventName = string.Empty;
    private string _selectedSdPath = string.Empty;
    private string _selectedSdContextPath = string.Empty;
    private string _rawExtensionsText;
    private string _countLabel = "RAW:0 / JPG:0 / 動画:0";
    private string _progressLabel = "待機中";
    private int _progressValue;
    private int _progressMaximum = 1;
    private bool _progressIsIndeterminate;
    private string _lightweightInputValidationMessage = string.Empty;
    private string _destinationStorageValidationMessage = string.Empty;
    private string _lastImportBasePath = string.Empty;
    private string _completionSummary = string.Empty;
    private string _logText = string.Empty;
    private string _safetyHeadline = "未検証 — 最終確認が完了するまでSDカードを再利用しないでください";
    private string _safetyDetail = "判定対象: JPG/JPEG・設定済みRAW・MOV/MP4。その他の形式は取り込み・判定対象外です。";
    private IBrush _safetyBrush = Brushes.DarkOrange;
    private bool _showSafetyPanel;
    private bool _autoStart;

    public MainWindowViewModel(
        IStorageVolumeProvider volumeProvider,
        StorageSessionTracker storageSessions,
        CameraCardRootResolver cardRoots)
        : this(volumeProvider, storageSessions, cardRoots, StorageEjectServiceFactory.Create())
    {
    }

    public MainWindowViewModel(
        IStorageVolumeProvider volumeProvider,
        StorageSessionTracker storageSessions,
        CameraCardRootResolver cardRoots,
        IStorageEjectService storageEjectService)
    {
        _volumeProvider = volumeProvider;
        _storageSessions = storageSessions;
        _cardRoots = cardRoots;
        _storageEjectService = storageEjectService;

        var preferences = _preferencesStore.Load();
        _destinationPath = preferences.DestinationPath;
        _rawExtensionsText = string.Join(", ", preferences.RawExtensions);
        _autoStart = _startupRegistration.IsEnabled();
        _classifier = new MediaClassifier(MediaClassifier.ParseRawExtensions(_rawExtensionsText));
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
            DestinationNeedsReselection = false;
            _destinationStorageValidationMessage = string.Empty;
            RefreshLightweightInputValidation();
            if (!IsSafeToReuseCurrentCard)
            {
                SetNotVerified("保存先が変更されました。次の取り込み後に再検証が必要です。");
            }
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
            RefreshLightweightInputValidation();
            RaiseCommandState();
        }
    }

    public string SelectedSdPath
    {
        get => _selectedSdPath;
        private set => SetField(ref _selectedSdPath, value);
    }

    private string SelectedSdContextPath
    {
        get => _selectedSdContextPath;
        set
        {
            if (string.Equals(_selectedSdContextPath, value, StringComparison.Ordinal)) return;
            _selectedSdContextPath = value;
            OnPropertyChanged(nameof(HasSelectedSd));
            OnPropertyChanged(nameof(SelectedSdDisplay));
            RaiseWorkflowState();
        }
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
                SetNotVerified("RAW拡張子設定が変更されました。SDカードを再スキャンしてください。");
            }
            SavePreferences();
            RaiseCommandState();
        }
    }

    public string CountLabel
    {
        get => _countLabel;
        private set
        {
            if (!SetField(ref _countLabel, value)) return;
            RaiseWorkflowState();
        }
    }

    public string ProgressLabel
    {
        get => _progressLabel;
        private set => SetField(ref _progressLabel, value);
    }

    public int ProgressValue
    {
        get => _progressValue;
        private set => SetField(ref _progressValue, value);
    }

    public int ProgressMaximum
    {
        get => _progressMaximum;
        private set => SetField(ref _progressMaximum, Math.Max(1, value));
    }

    public bool ProgressIsIndeterminate
    {
        get => _progressIsIndeterminate;
        private set => SetField(ref _progressIsIndeterminate, value);
    }

    public string InputValidationMessage => !string.IsNullOrWhiteSpace(_lightweightInputValidationMessage)
        ? _lightweightInputValidationMessage
        : _destinationStorageValidationMessage;

    public bool HasInputValidationError => !string.IsNullOrWhiteSpace(InputValidationMessage);

    public string LastImportBasePath
    {
        get => _lastImportBasePath;
        private set
        {
            if (!SetField(ref _lastImportBasePath, value)) return;
            OnPropertyChanged(nameof(HasCompletedImport));
        }
    }

    public string CompletionSummary
    {
        get => _completionSummary;
        private set
        {
            if (!SetField(ref _completionSummary, value)) return;
            OnPropertyChanged(nameof(HasCompletedImport));
        }
    }

    public bool HasCompletedImport => !string.IsNullOrWhiteSpace(LastImportBasePath)
        && !string.IsNullOrWhiteSpace(CompletionSummary);

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

    public bool ShowSafetyPanel
    {
        get => _showSafetyPanel;
        private set => SetField(ref _showSafetyPanel, value);
    }

    public bool IsSafeToReuseCurrentCard
    {
        get => _isSafeToReuseCurrentCard;
        private set
        {
            if (!SetField(ref _isSafeToReuseCurrentCard, value)) return;
            OnPropertyChanged(nameof(CanImport));
            OnPropertyChanged(nameof(CanEjectSelectedSd));
            RaiseWorkflowState();
        }
    }

    public bool DestinationNeedsReselection
    {
        get => _destinationNeedsReselection;
        private set
        {
            if (!SetField(ref _destinationNeedsReselection, value)) return;
            OnPropertyChanged(nameof(CanImport));
            RaiseWorkflowState();
        }
    }

    public bool AutoStart
    {
        get => _autoStart;
        set
        {
            if (value == _autoStart) return;

            if (_startupRegistration.SetEnabled(value, out var error))
            {
                SetField(ref _autoStart, value);
                AppendLog(value ? "ログイン時自動起動を有効にしました。" : "ログイン時自動起動を無効にしました。");
                return;
            }

            AppendLog($"ログイン時自動起動の設定に失敗しました: {error}");
            SetField(ref _autoStart, _startupRegistration.IsEnabled());
        }
    }

    public bool AutoStartSupported => _startupRegistration.IsSupported;
    public bool IsBusy => _isScanning || _isProcessing;
    public bool IsProcessing => _isProcessing;
    public bool HasSelectedSd => !string.IsNullOrWhiteSpace(SelectedSdContextPath);
    public bool ShowMediaSummary => _scanSession is not null;
    public string SelectedSdDisplay => HasSelectedSd ? SelectedSdContextPath : "未選択";
    public bool CanImport => !IsBusy
        && !IsSafeToReuseCurrentCard
        && !DestinationNeedsReselection
        && !HasInputValidationError
        && _scanSession is not null
        && !string.IsNullOrWhiteSpace(DestinationPath)
        && !string.IsNullOrWhiteSpace(EventName);
    public bool CanCancel => IsBusy;
    public int PendingSdCount => _pendingCards.Count;
    public bool HasPendingCards => PendingSdCount > 0;
    public string PendingSdText => PendingSdCount == 0
        ? string.Empty
        : $"待機中のSDカード: {PendingSdCount} 枚";

    public string WorkflowHeadline
    {
        get
        {
            if (_isScanning) return "SDカードを確認しています";
            if (_isProcessing) return "取り込みと安全確認を実行しています";
            if (IsSafeToReuseCurrentCard && _scanSession is not null) return "取り込みと検証が完了しました";
            if (_scanSession is null && HasSelectedSd) return "SDカードを再選択してください";
            if (_scanSession is null) return "SDカードを選択してください";
            if (DestinationNeedsReselection) return "保存先を再選択してください";
            if (string.IsNullOrWhiteSpace(DestinationPath)) return "保存先を選択してください";
            if (string.IsNullOrWhiteSpace(EventName)) return "イベント名を入力してください";
            if (HasInputValidationError) return "入力内容を確認してください";
            return "取り込みの準備ができました";
        }
    }

    public string WorkflowDetail
    {
        get
        {
            if (_isScanning)
            {
                return "カード全体をスキャンして、対象ファイルとストレージの状態を確認しています。完了までカードを取り外さないでください。";
            }

            if (_isProcessing)
            {
                return "コピー後にSDカードを再スキャンし、保存先の実ファイルとSHA-256を照合します。完了表示までカードを取り外さないでください。";
            }

            if (IsSafeToReuseCurrentCard && _scanSession is not null)
            {
                return "SDカードを再利用できます。保存先を開いて確認することもできます。カードを取り外すか別のカードを選択すると次の取り込みへ進めます。";
            }

            if (_scanSession is null && HasSelectedSd)
            {
                return "選択したカードの確認が完了していません。下の状態を確認し、カードを挿し直すか「SDカードを選択…」から選択し直してください。";
            }

            if (_scanSession is null)
            {
                return "SDカードを挿すと自動検出します。検出されない場合は「SDカードを選択…」からカード内のフォルダを選択してください。";
            }

            if (DestinationNeedsReselection)
            {
                return "設定していた保存先が取り外されました。「選択…」から現在利用できる保存先を選び直してください。";
            }

            if (string.IsNullOrWhiteSpace(DestinationPath))
            {
                return "写真ライブラリの親フォルダを指定してください。保存先は次回起動時も保持されます。";
            }

            if (string.IsNullOrWhiteSpace(EventName))
            {
                return "この撮影を識別できる短い名前を入力してください。例: 旅行、運動会、撮影会。";
            }

            if (HasInputValidationError)
            {
                return InputValidationMessage;
            }

            return $"{CountLabel} を確認しました。保存先プレビューを確認して、取り込みを開始してください。";
        }
    }

    public void SetDestinationFromPicker(string path)
    {
        if (_disposed || IsBusy) return;

        var wasBlocked = DestinationNeedsReselection;
        DestinationNeedsReselection = false;
        if (string.Equals(DestinationPath, path, _storageSessions.PathComparison))
        {
            if (wasBlocked && HasSelectedSd && !IsSafeToReuseCurrentCard)
            {
                SetNotVerified("保存先を再確認しました。取り込み後の最終検証が完了するまでSDカードを再利用しないでください。");
            }
            _destinationStorageValidationMessage = string.Empty;
            RefreshLightweightInputValidation();
            RaiseCommandState();
            return;
        }

        DestinationPath = path;
    }

    public async Task ValidateDestinationAsync()
    {
        if (_disposed) return;

        RefreshLightweightInputValidation();
        if (!string.IsNullOrWhiteSpace(_lightweightInputValidationMessage)) return;

        var session = _scanSession;
        var destination = DestinationPath;
        if (session is null || string.IsNullOrWhiteSpace(destination))
        {
            SetDestinationStorageValidationMessage(string.Empty);
            return;
        }

        string validationMessage;
        try
        {
            validationMessage = await Task.Run(() => GetDestinationStorageValidationMessage(session, destination)).ConfigureAwait(true);
        }
        catch (Exception exception)
        {
            AppendLog($"保存先の事前確認に失敗: {exception.Message}");
            validationMessage = "保存先の状態を確認できませんでした。「選択…」から保存先を選び直してください。";
        }

        if (_disposed
            || !ReferenceEquals(session, _scanSession)
            || !string.Equals(destination, DestinationPath, StringComparison.Ordinal)) return;

        SetDestinationStorageValidationMessage(validationMessage);
    }

    public bool CanCloseWindow()
    {
        if (_disposed || !IsBusy) return true;
        AppendLog("処理中の通常終了を抑止しました。必要ならキャンセルして処理終了後に閉じてください。");
        return false;
    }

    public void ReportAuxiliaryFailure(string operation, Exception exception)
    {
        if (_disposed) return;
        ProgressLabel = $"{operation}に失敗しました";
        AppendLog($"{operation}失敗: {exception.Message}");
    }

    private ImportCoordinator CreateCoordinator() => new(
        _classifier,
        _storageSessions,
        _cardRoots,
        _volumeProvider);

    private void RebuildCoordinator()
    {
        _classifier = new MediaClassifier(MediaClassifier.ParseRawExtensions(RawExtensionsText));
        _coordinator = CreateCoordinator();
    }

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
        CountLabel = $"RAW:{raw} / JPG:{jpg} / 動画:{video}";
    }

    private void SetBlocked(string message)
    {
        SafetyHeadline = "SDカードを再利用しないでください";
        SafetyDetail = message;
        SafetyBrush = Brushes.IndianRed;
        if (HasSelectedSd || IsBusy || ShowSafetyPanel) ShowSafetyPanel = true;
    }

    private void SetNotVerified(string detail)
    {
        SafetyHeadline = "未検証 — 最終確認が完了するまでSDカードを再利用しないでください";
        SafetyDetail = detail;
        SafetyBrush = Brushes.DarkOrange;
        if (HasSelectedSd || IsBusy || ShowSafetyPanel) ShowSafetyPanel = true;
    }

    private void ClearScanSession()
    {
        _scanSession = null;
        OnPropertyChanged(nameof(ShowMediaSummary));
        SelectedSdPath = string.Empty;
        SelectedSdContextPath = string.Empty;
        CountLabel = "RAW:0 / JPG:0 / 動画:0";
        ShowSafetyPanel = false;
        IsSafeToReuseCurrentCard = false;
        _destinationStorageValidationMessage = string.Empty;
        RefreshLightweightInputValidation();
        RaiseCommandState();
    }

    private void ClearCompletion()
    {
        LastImportBasePath = string.Empty;
        CompletionSummary = string.Empty;
    }

    private void SetCompletion(string basePath, int verified)
    {
        LastImportBasePath = basePath;
        CompletionSummary = $"{verified} 件の写真・動画について、保存先の実ファイルとSHA-256一致・永続化を確認しました。";
    }

    private void SetProgressState(string label, int current = 0, int total = 1, bool indeterminate = false)
    {
        ProgressLabel = label;
        ProgressMaximum = total;
        ProgressValue = Math.Clamp(current, 0, ProgressMaximum);
        ProgressIsIndeterminate = indeterminate;
    }

    private void ApplyImportProgress(ImportProgress update)
    {
        switch (update.Phase)
        {
            case ImportProgressPhase.Copying:
            {
                var total = Math.Max(1, update.Total);
                var isComplete = update.Current >= update.Total;
                var displayCurrent = isComplete ? update.Total : Math.Min(update.Current + 1, update.Total);
                SetProgressState(
                    isComplete ? "コピー完了 — SDカードの安全確認へ進みます" : $"コピー中 {displayCurrent} / {update.Total}",
                    update.Current,
                    total,
                    indeterminate: false);
                break;
            }
            case ImportProgressPhase.Rescanning:
                SetProgressState("SDカードを再スキャン中…", 0, Math.Max(1, update.Total), indeterminate: true);
                break;
            case ImportProgressPhase.Verifying:
                SetProgressState("保存先のサイズ・SHA-256・永続化を検証中…", update.Current, Math.Max(1, update.Total), indeterminate: true);
                break;
            case ImportProgressPhase.Scanning:
                SetProgressState("SDカードを確認中…", update.Current, Math.Max(1, update.Total), indeterminate: true);
                break;
        }
    }

    private void RefreshLightweightInputValidation()
    {
        var message = string.Empty;
        var session = _scanSession;

        if (session is not null && !string.IsNullOrWhiteSpace(DestinationPath))
        {
            try
            {
                var destination = PathSafety.Normalize(DestinationPath);
                if (PathSafety.IsSameOrDescendant(destination, session.CardRoot, _storageSessions.PathComparison)
                    || PathSafety.IsSameOrDescendant(session.CardRoot, destination, _storageSessions.PathComparison))
                {
                    message = "保存先にはSDカード自身や、その親・子フォルダを指定できません。別のストレージ上のフォルダを選択してください。";
                }
            }
            catch
            {
                message = "保存先のパスを確認してください。「選択…」からフォルダを選ぶと確実です。";
            }
        }

        if (string.IsNullOrWhiteSpace(message) && !string.IsNullOrWhiteSpace(EventName))
        {
            var sanitized = ImportCoordinator.SanitizeEventName(EventName.Trim());
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                message = "イベント名にフォルダ名として使える文字を入力してください。";
            }
        }

        if (string.Equals(_lightweightInputValidationMessage, message, StringComparison.Ordinal)) return;
        _lightweightInputValidationMessage = message;
        RaiseInputValidationState();
    }

    private string GetDestinationStorageValidationMessage(ImportScanSession session, string destinationPath)
    {
        string destination;
        try
        {
            destination = PathSafety.Normalize(destinationPath);
        }
        catch
        {
            return "保存先のパスを確認してください。「選択…」からフォルダを選ぶと確実です。";
        }

        if (!Directory.Exists(destination))
        {
            return "保存先フォルダが見つかりません。接続状態を確認して、現在利用できるフォルダを選択してください。";
        }

        var identity = _storageSessions.Capture(destination);
        if (identity is null)
        {
            return "保存先のストレージ情報を確認できません。「選択…」から別の保存先を選択してください。";
        }

        if (string.IsNullOrWhiteSpace(identity.PhysicalDeviceFingerprint))
        {
            return "保存先の物理ストレージを安全に確認できません。物理デバイスを識別できる別の保存先を選択してください。";
        }

        if (identity.SessionId == session.SourceIdentity.SessionId
            || string.Equals(identity.Fingerprint, session.SourceIdentity.Fingerprint, StringComparison.Ordinal))
        {
            return "保存先はSDカードとは別のボリュームを選択してください。";
        }

        if (string.Equals(identity.PhysicalDeviceFingerprint, session.SourceIdentity.PhysicalDeviceFingerprint, StringComparison.Ordinal))
        {
            return "保存先はSDカードとは別の物理ストレージを選択してください。同じ物理ドライブ内の別ボリュームは使用できません。";
        }

        return string.Empty;
    }

    private void SetDestinationStorageValidationMessage(string message)
    {
        if (string.Equals(_destinationStorageValidationMessage, message, StringComparison.Ordinal)) return;
        _destinationStorageValidationMessage = message;
        RaiseInputValidationState();
    }

    private void RaiseInputValidationState()
    {
        OnPropertyChanged(nameof(InputValidationMessage));
        OnPropertyChanged(nameof(HasInputValidationError));
        OnPropertyChanged(nameof(CanImport));
        RaiseWorkflowState();
    }

    public void ReportUiFailure(string operation, Exception exception) =>
        ReportOperationFailure(operation, exception);

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
            SetProgressState($"{operation}失敗");
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
        var preferences = new AppPreferences(
            DestinationPath,
            MediaClassifier.ParseRawExtensions(RawExtensionsText));
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
        OnPropertyChanged(nameof(CanEjectSelectedSd));
        RaiseWorkflowState();
    }

    private void RaiseWorkflowState()
    {
        OnPropertyChanged(nameof(WorkflowHeadline));
        OnPropertyChanged(nameof(WorkflowDetail));
    }

    private void RaisePendingState()
    {
        OnPropertyChanged(nameof(PendingSdCount));
        OnPropertyChanged(nameof(HasPendingCards));
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
