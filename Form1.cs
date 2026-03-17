using MetadataExtractor;
using MetadataExtractor.Formats.Exif;
using System.Management;
using System.Text.Json;

namespace PhotoOrganizer;

public partial class Form1 : Form
{
    private readonly NotifyIcon _trayIcon;
    private readonly ManagementEventWatcher? _insertWatcher;
    private readonly System.Windows.Forms.Timer _drivePollTimer;
    private TextBox _destinationText = null!;
    private TextBox _eventNameText = null!;
    private TextBox _selectedSdText = null!;
    private TextBox _logText = null!;
    private Panel _mainPanel = null!;
    private Label _countLabel = null!;
    private Label _progressLabel = null!;
    private Button _selectSdButton = null!;
    private Button _startButton = null!;

    private List<string> _scannedFiles = [];
    private bool _isProcessing;
    private bool _allowClose;
    private bool _isPollingDrives;
    private HashSet<string> _knownCandidateDrives = new(StringComparer.OrdinalIgnoreCase);
    private int _logTop;
    private readonly string _statePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "PhotoOrganizer",
        "state.json");
    private readonly string _configPath = Path.Combine(
        AppContext.BaseDirectory,
        "config.json");
    private HashSet<string> _rawExt = new(StringComparer.OrdinalIgnoreCase);
    private static readonly string[] DefaultRawExtensions =
    [
        ".arw", ".cr2", ".cr3", ".nef", ".dng", ".raf", ".rw2", ".orf", ".pef"
    ];

    public Form1()
    {
        InitializeComponent();
        Text = "Photo Organizer";
        Width = 860;
        Height = 620;
        BuildUi();
        LoadConfig();
        LoadState();

        _trayIcon = new NotifyIcon
        {
            Icon = SystemIcons.Application,
            Text = "Photo Organizer",
            Visible = true,
            ContextMenuStrip = BuildTrayMenu()
        };
        _trayIcon.DoubleClick += (_, _) => ShowFromTray();
        _drivePollTimer = new System.Windows.Forms.Timer { Interval = 3000 };
        _drivePollTimer.Tick += async (_, _) => await PollForInsertedDriveAsync();
        _drivePollTimer.Start();

        try
        {
            _insertWatcher = new ManagementEventWatcher(
                new WqlEventQuery("SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2"));
            _insertWatcher.EventArrived += OnVolumeInserted;
            _insertWatcher.Start();
        }
        catch (Exception ex)
        {
            AppendLog($"SD監視開始失敗: {ex.Message}");
        }

        Load += (_, _) => HideToTray();
        Shown += async (_, _) => await InitializeSdSelectionAsync();
        FormClosing += OnFormClosing;
    }

    private ContextMenuStrip BuildTrayMenu()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("表示", null, (_, _) => ShowFromTray());
        menu.Items.Add("終了", null, (_, _) =>
        {
            _allowClose = true;
            Close();
        });
        return menu;
    }

    private void BuildUi()
    {
        _mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12) };
        Controls.Add(_mainPanel);

        var y = 12;
        _mainPanel.Controls.Add(MakeLabel("保存先パス", 12, y)); y += 22;
        _destinationText = MakeTextBox(12, y, 800);
        _destinationText.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
        _destinationText.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _destinationText.TextChanged += (_, _) => SaveState();
        _mainPanel.Controls.Add(_destinationText); y += 34;

        _mainPanel.Controls.Add(MakeLabel("イベント名", 12, y)); y += 22;
        _eventNameText = MakeTextBox(12, y, 800);
        _eventNameText.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _mainPanel.Controls.Add(_eventNameText); y += 34;

        _mainPanel.Controls.Add(MakeLabel("選択中のSDカード", 12, y)); y += 22;
        _selectedSdText = MakeTextBox(12, y, 800);
        _selectedSdText.ReadOnly = true;
        _selectedSdText.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _mainPanel.Controls.Add(_selectedSdText); y += 36;

        _selectSdButton = new Button { Left = 12, Top = y, Width = 130, Height = 32, Text = "SDカード選択" };
        _selectSdButton.Click += async (_, _) => await SelectSdAndScanAsync();
        _mainPanel.Controls.Add(_selectSdButton);

        _startButton = new Button { Left = 156, Top = y, Width = 120, Height = 32, Text = "処理開始" };
        _startButton.Click += async (_, _) => await StartProcessAsync();
        _mainPanel.Controls.Add(_startButton);
        y += 42;

        _countLabel = MakeLabel("RAW:0 / JPG:0 / MP4:0", 12, y);
        _mainPanel.Controls.Add(_countLabel); y += 30;
        _progressLabel = MakeLabel("待機中", 12, y);
        _mainPanel.Controls.Add(_progressLabel); y += 30;
        _logTop = y;

        _logText = new TextBox
        {
            Left = 12,
            Top = _logTop,
            Width = 800,
            Height = 200,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            ReadOnly = true
        };
        _logText.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        _mainPanel.Controls.Add(_logText);
        _mainPanel.Resize += (_, _) => AdjustLogArea();
        AdjustLogArea();
    }

    private async Task<bool> SelectSdAndScanAsync(string? forcedRoot = null, bool autoDetected = false)
    {
        if (_isProcessing) return false;

        var root = forcedRoot;
        if (string.IsNullOrWhiteSpace(root))
        {
            using var dlg = new FolderBrowserDialog { Description = "SDカードを選択" };
            if (dlg.ShowDialog(this) != DialogResult.OK) return false;
            root = dlg.SelectedPath;
        }

        AppendLog($"スキャン開始: {root}");
        List<string> files;
        try
        {
            files = await Task.Run(() => EnumerateMediaFiles(root!));
        }
        catch (Exception ex)
        {
            AppendLog($"スキャン失敗: {ex.Message}");
            return false;
        }

        if (autoDetected && files.Count == 0)
        {
            AppendLog($"自動選択スキップ（メディアなし）: {root}");
            return false;
        }

        _selectedSdText.Text = root ?? "";
        _scannedFiles = files;
        SaveState();
        var (raw, jpg, mp4) = CountByType(files);
        _countLabel.Text = $"RAW:{raw} / JPG:{jpg} / MP4:{mp4}";
        AppendLog($"{files.Count} 件検出 / RAW:{raw} / JPG:{jpg} / MP4:{mp4}");
        return true;
    }

    private async Task StartProcessAsync()
    {
        if (_isProcessing) return;
        if (string.IsNullOrWhiteSpace(_destinationText.Text) || string.IsNullOrWhiteSpace(_eventNameText.Text) || _scannedFiles.Count == 0)
        {
            AppendLog("保存先・イベント名・SDカード選択を確認してください。");
            return;
        }

        _isProcessing = true;
        _startButton.Enabled = false;
        _selectSdButton.Enabled = false;
        try
        {
            var filesToProcess = _scannedFiles.ToList();
            var destination = _destinationText.Text.Trim();
            var eventName = _eventNameText.Text.Trim();
            var dateKey = await Task.Run(() => ResolveDateKey(filesToProcess));
            var folderName = $"{dateKey}_{SanitizeName(eventName)}";
            var yearPath = Path.Combine(destination, dateKey[..4]);
            var basePath = Path.Combine(yearPath, folderName);
            System.IO.Directory.CreateDirectory(Path.Combine(basePath, "RAW"));
            System.IO.Directory.CreateDirectory(Path.Combine(basePath, "JPG"));
            System.IO.Directory.CreateDirectory(Path.Combine(basePath, "MP4"));

            AppendLog("処理開始");
            SetProgress($"処理中... 0/{filesToProcess.Count}");
            var result = await Task.Run(() =>
            {
                var raw = 0; var jpg = 0; var mp4 = 0; var skipUnsupported = 0; var skipDuplicate = 0; var failed = 0; var processed = 0;
                var failedFiles = new List<string>();
                foreach (var src in filesToProcess)
                {
                    var kind = GetMediaKind(src);
                    if (kind == "RAW") raw++;
                    else if (kind == "JPG") jpg++;
                    else if (kind == "MP4") mp4++;

                    var status = ProcessOneFile(src, basePath, out _);
                    if (status == FileProcessStatus.SkippedUnsupported) skipUnsupported++;
                    else if (status == FileProcessStatus.SkippedDuplicate) skipDuplicate++;
                    else if (status == FileProcessStatus.Failed)
                    {
                        failed++;
                        failedFiles.Add(src);
                    }

                    processed++;
                    if (processed % 25 == 0 || processed == filesToProcess.Count)
                    {
                        SetProgress($"処理中... {processed}/{filesToProcess.Count}");
                    }
                }

                if (failedFiles.Count > 0)
                {
                    Thread.Sleep(800);
                    var retryFailed = new List<string>();
                    foreach (var src in failedFiles)
                    {
                        var retryStatus = ProcessOneFile(src, basePath, out _);
                        if (retryStatus == FileProcessStatus.Failed)
                        {
                            retryFailed.Add(src);
                        }
                    }
                    failed = retryFailed.Count;
                }

                return (raw, jpg, mp4, skipUnsupported, skipDuplicate, failed);
            });

            AppendLog($"完了: RAW:{result.raw} / JPG:{result.jpg} / MP4:{result.mp4} / 重複スキップ:{result.skipDuplicate} / 非対応スキップ:{result.skipUnsupported} / 失敗:{result.failed}");
            AppendLog($"保存先: {basePath}");
            SaveState();
        }
        catch (Exception ex)
        {
            AppendLog($"保存に失敗しました: {ex.Message}");
        }
        finally
        {
            _isProcessing = false;
            _startButton.Enabled = true;
            _selectSdButton.Enabled = true;
            SetProgress("待機中");
        }
    }

    private void OnVolumeInserted(object sender, EventArrivedEventArgs e)
    {
        var drive = e.NewEvent.Properties["DriveName"]?.Value?.ToString();
        if (string.IsNullOrWhiteSpace(drive))
        {
            BeginInvoke(new Action(() => _ = PollForInsertedDriveAsync()));
            return;
        }
        if (!drive.EndsWith("\\")) drive += "\\";
        BeginInvoke(new Action(() => _ = HandleVolumeInsertedAsync(drive)));
    }

    private async Task HandleVolumeInsertedAsync(string drive)
    {
        try
        {
            var ready = false;
            for (var i = 0; i < 5; i++)
            {
                if (IsCandidateSdDrive(drive))
                {
                    ready = true;
                    break;
                }
                await Task.Delay(1000);
            }
            if (!ready) return;

            AppendLog($"SDカード検出: {drive}");
            var selected = await SelectSdAndScanAsync(drive, autoDetected: true);
            if (selected) ShowFromTray();
        }
        catch (Exception ex)
        {
            AppendLog($"SD自動処理エラー: {ex.Message}");
        }
    }

    private List<string> EnumerateMediaFiles(string root)
    {
        var list = new List<string>();
        var dirs = new Stack<string>();
        dirs.Push(root);
        while (dirs.Count > 0)
        {
            var current = dirs.Pop();
            string[] subDirs;
            string[] files;
            try
            {
                subDirs = System.IO.Directory.GetDirectories(current);
                files = System.IO.Directory.GetFiles(current);
            }
            catch
            {
                continue;
            }

            foreach (var sub in subDirs)
            {
                dirs.Push(sub);
            }

            foreach (var file in files)
            {
                var name = Path.GetFileName(file);
                if (name.StartsWith(".")) continue;
                var ext = Path.GetExtension(file);
                if (GetMediaKind(file) is not null)
                {
                    list.Add(file);
                }
            }
        }
        return list;
    }

    private (int raw, int jpg, int mp4) CountByType(IEnumerable<string> files)
    {
        var raw = 0; var jpg = 0; var mp4 = 0;
        foreach (var f in files)
        {
            var kind = GetMediaKind(f);
            if (kind == "RAW") raw++;
            else if (kind == "JPG") jpg++;
            else if (kind == "MP4") mp4++;
        }
        return (raw, jpg, mp4);
    }

    private static string ResolveDateKey(List<string> files)
    {
        var target = files.OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase).FirstOrDefault();
        if (target is null) return DateTime.Now.ToString("yyyy-MM-dd");
        try
        {
            var directories = MetadataExtractor.ImageMetadataReader.ReadMetadata(target);
            var exif = directories.OfType<ExifSubIfdDirectory>().FirstOrDefault();
            DateTime dt;
            if (exif is not null && exif.TryGetDateTime(ExifDirectoryBase.TagDateTimeOriginal, out dt))
            {
                return dt.ToString("yyyy-MM-dd");
            }
            if (exif is not null && exif.TryGetDateTime(ExifDirectoryBase.TagDateTimeDigitized, out dt))
            {
                return dt.ToString("yyyy-MM-dd");
            }
            dt = File.GetCreationTime(target);
            return dt.ToString("yyyy-MM-dd");
        }
        catch
        {
            return File.GetCreationTime(target).ToString("yyyy-MM-dd");
        }
    }

    private static string SanitizeName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        return name;
    }

    private string? GetMediaKind(string path)
    {
        var ext = Path.GetExtension(path);
        if (_rawExt.Contains(ext)) return "RAW";
        if (ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase) || ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase)) return "JPG";
        if (ext.Equals(".mp4", StringComparison.OrdinalIgnoreCase) || ext.Equals(".mov", StringComparison.OrdinalIgnoreCase)) return "MP4";
        return null;
    }

    private FileProcessStatus ProcessOneFile(string src, string basePath, out string? error)
    {
        error = null;
        var kind = GetMediaKind(src);
        if (kind is null) return FileProcessStatus.SkippedUnsupported;

        var dst = Path.Combine(basePath, kind, Path.GetFileName(src));
        try
        {
            var srcInfo = new FileInfo(src);
            if (File.Exists(dst))
            {
                var dstInfo = new FileInfo(dst);
                if (IsSameByTimeAndSize(srcInfo, dstInfo))
                {
                    return FileProcessStatus.SkippedDuplicate;
                }
            }

            File.Copy(src, dst, true);
            File.SetLastWriteTimeUtc(dst, srcInfo.LastWriteTimeUtc);
            var copiedInfo = new FileInfo(dst);
            if (!IsSameByTimeAndSize(srcInfo, copiedInfo))
            {
                error = "整合性チェック失敗";
                return FileProcessStatus.Failed;
            }

            return FileProcessStatus.Copied;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return FileProcessStatus.Failed;
        }
    }

    private static bool IsSameByTimeAndSize(FileInfo src, FileInfo dst)
    {
        if (src.Length != dst.Length) return false;
        var diff = Math.Abs((src.LastWriteTimeUtc - dst.LastWriteTimeUtc).TotalSeconds);
        return diff <= 2;
    }

    private void SetProgress(string text)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(SetProgress), text);
            return;
        }

        _progressLabel.Text = text;
    }

    private void AppendLog(string line)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(AppendLog), line);
            return;
        }

        _logText.AppendText($"{DateTime.Now:HH:mm:ss} {line}{Environment.NewLine}");
        _logText.SelectionStart = _logText.TextLength;
        _logText.ScrollToCaret();
    }

    private bool IsCandidateSdDrive(string driveRoot)
    {
        try
        {
            var normalized = driveRoot.EndsWith("\\") ? driveRoot : $"{driveRoot}\\";
            var info = new DriveInfo(normalized);
            if (!info.IsReady) return false;
            if (info.DriveType is DriveType.CDRom or DriveType.Network) return false;
            var systemRoot = Path.GetPathRoot(Environment.SystemDirectory);
            if (!string.IsNullOrWhiteSpace(systemRoot) && normalized.Equals(systemRoot, StringComparison.OrdinalIgnoreCase)) return false;
            if (info.DriveType == DriveType.Removable) return true;
            return System.IO.Directory.Exists(Path.Combine(normalized, "DCIM")) ||
                   System.IO.Directory.Exists(Path.Combine(normalized, "PRIVATE"));
        }
        catch
        {
            return false;
        }
    }

    private async Task InitializeSdSelectionAsync()
    {
        if (!string.IsNullOrWhiteSpace(_selectedSdText.Text) && IsCandidateSdDrive(_selectedSdText.Text))
        {
            var reused = await SelectSdAndScanAsync(_selectedSdText.Text, autoDetected: true);
            if (reused) return;
        }

        var candidates = DriveInfo.GetDrives()
            .Where(d => d.DriveType == DriveType.Removable && d.IsReady)
            .OrderBy(d => d.Name, StringComparer.OrdinalIgnoreCase)
            .Select(d => d.RootDirectory.FullName)
            .ToList();
        foreach (var drive in candidates)
        {
            if (await SelectSdAndScanAsync(drive, autoDetected: true))
            {
                _knownCandidateDrives = [.. candidates];
                return;
            }
        }
        _knownCandidateDrives = [.. candidates];
    }

    private async Task PollForInsertedDriveAsync()
    {
        if (_isPollingDrives) return;
        _isPollingDrives = true;
        try
        {
            var current = DriveInfo.GetDrives()
                .Where(d => IsCandidateSdDrive(d.RootDirectory.FullName))
                .Select(d => d.RootDirectory.FullName)
                .OrderBy(d => d, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var inserted = current.Where(d => !_knownCandidateDrives.Contains(d)).ToList();
            _knownCandidateDrives = [.. current];
            foreach (var drive in inserted)
            {
                await HandleVolumeInsertedAsync(drive);
            }
        }
        finally
        {
            _isPollingDrives = false;
        }
    }

    private void LoadConfig()
    {
        _rawExt = new HashSet<string>(DefaultRawExtensions, StringComparer.OrdinalIgnoreCase);
        try
        {
            if (!File.Exists(_configPath))
            {
                return;
            }

            var json = File.ReadAllText(_configPath);
            var config = JsonSerializer.Deserialize<AppConfig>(json);
            if (config?.RawExtensions is null || config.RawExtensions.Count == 0)
            {
                return;
            }

            _rawExt = new HashSet<string>(
                config.RawExtensions
                    .Select(NormalizeExtension)
                    .Where(e => !string.IsNullOrWhiteSpace(e)),
                StringComparer.OrdinalIgnoreCase);
            if (_rawExt.Count == 0)
            {
                _rawExt = new HashSet<string>(DefaultRawExtensions, StringComparer.OrdinalIgnoreCase);
            }
        }
        catch (Exception ex)
        {
            AppendLog($"設定ファイル読み込み失敗: {ex.Message}");
            _rawExt = new HashSet<string>(DefaultRawExtensions, StringComparer.OrdinalIgnoreCase);
        }
    }

    private static string NormalizeExtension(string ext)
    {
        if (string.IsNullOrWhiteSpace(ext)) return "";
        var value = ext.Trim();
        if (!value.StartsWith(".")) value = $".{value}";
        return value.ToLowerInvariant();
    }

    private void LoadState()
    {
        try
        {
            if (!File.Exists(_statePath)) return;
            var json = File.ReadAllText(_statePath);
            var state = JsonSerializer.Deserialize<AppState>(json);
            if (state is null) return;
            if (!string.IsNullOrWhiteSpace(state.DestinationPath)) _destinationText.Text = state.DestinationPath;
            if (!string.IsNullOrWhiteSpace(state.SelectedSdPath)) _selectedSdText.Text = state.SelectedSdPath;
        }
        catch (Exception ex)
        {
            AppendLog($"前回設定の読み込み失敗: {ex.Message}");
        }
    }

    private void SaveState()
    {
        try
        {
            var dir = Path.GetDirectoryName(_statePath);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            var state = new AppState
            {
                DestinationPath = _destinationText.Text.Trim(),
                SelectedSdPath = _selectedSdText.Text.Trim()
            };
            File.WriteAllText(_statePath, JsonSerializer.Serialize(state));
        }
        catch (Exception ex)
        {
            AppendLog($"前回設定の保存失敗: {ex.Message}");
        }
    }

    private void HideToTray()
    {
        Hide();
        ShowInTaskbar = false;
    }

    private void ShowFromTray()
    {
        Show();
        WindowState = FormWindowState.Normal;
        ShowInTaskbar = true;
        Activate();
    }

    private void AdjustLogArea()
    {
        if (_mainPanel is null || _logText is null) return;
        var available = _mainPanel.ClientSize.Height - _logTop - 12;
        _logText.Height = Math.Max(80, available);
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        if (!_allowClose)
        {
            e.Cancel = true;
            HideToTray();
            return;
        }

        _insertWatcher?.Stop();
        _insertWatcher?.Dispose();
        _drivePollTimer.Stop();
        _drivePollTimer.Dispose();
        SaveState();
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
    }

    private static Label MakeLabel(string text, int left, int top)
        => new() { Left = left, Top = top, Width = 800, Height = 20, Text = text };

    private static TextBox MakeTextBox(int left, int top, int width)
        => new() { Left = left, Top = top, Width = width };

    private sealed class AppState
    {
        public string DestinationPath { get; set; } = "";
        public string SelectedSdPath { get; set; } = "";
    }

    private sealed class AppConfig
    {
        public List<string> RawExtensions { get; set; } = [];
    }

    private enum FileProcessStatus
    {
        Copied,
        SkippedUnsupported,
        SkippedDuplicate,
        Failed
    }
}
