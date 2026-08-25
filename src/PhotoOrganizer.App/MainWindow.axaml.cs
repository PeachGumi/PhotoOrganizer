using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed partial class MainWindow : Window
{
    private readonly MainWindowViewModel? _viewModel;
    private readonly BackgroundPreferencesStore _backgroundPreferences = new();
    private readonly IStartupRegistrationService _startupRegistration = new StartupRegistrationService();
    private bool _allowExplicitClose;
    private bool _changingBackgroundPreference;

    public MainWindow()
    {
        InitializeComponent();
    }

    public MainWindow(MainWindowViewModel viewModel) : this()
    {
        _viewModel = viewModel;
        DataContext = viewModel;
        viewModel.PropertyChanged += OnViewModelPropertyChanged;

        Closing += OnClosing;
        Closed += OnClosed;

        LoadBackgroundPreference();
        UpdateDestinationPreview();
    }

    private void OnClosing(object? sender, WindowClosingEventArgs e)
    {
        if (_allowExplicitClose) return;

        // Closing the workflow window is not application quit. Keep monitoring in
        // the tray/menu bar. While processing, keep the window visible as well.
        e.Cancel = true;
        if (_viewModel is not null && !_viewModel.CanCloseWindow()) return;
        Hide();
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        if (_viewModel is not null) _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
    }

    internal void AllowExplicitClose() => _allowExplicitClose = true;

    private async void SelectDestination_Click(object? sender, RoutedEventArgs e)
    {
        if (_viewModel is null || !StorageProvider.CanPickFolder) return;

        var folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "写真の保存先を選択",
            AllowMultiple = false
        });

        var path = folders.FirstOrDefault()?.TryGetLocalPath();
        if (!string.IsNullOrWhiteSpace(path)) _viewModel.SetDestinationFromPicker(path);
    }

    private async void SelectSd_Click(object? sender, RoutedEventArgs e)
    {
        if (_viewModel is null || !StorageProvider.CanPickFolder) return;

        var folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "SDカードを選択（DCIM/PRIVATE内のフォルダを選択してもカード全体を検証します）",
            AllowMultiple = false
        });

        var path = folders.FirstOrDefault()?.TryGetLocalPath();
        if (!string.IsNullOrWhiteSpace(path)) await _viewModel.ScanCardAsync(path);
    }

    private async void Import_Click(object? sender, RoutedEventArgs e)
    {
        if (_viewModel is not null) await _viewModel.StartImportAsync();
    }

    private void Cancel_Click(object? sender, RoutedEventArgs e)
    {
        _viewModel?.CancelImport();
    }

    private void StartInBackground_Changed(object? sender, RoutedEventArgs e)
    {
        if (_changingBackgroundPreference || _viewModel is null) return;

        var desired = StartInBackgroundCheckBox.IsChecked == true;
        var previous = _backgroundPreferences.Load().StartInBackground;
        if (desired == previous)
        {
            UpdateBackgroundSettingStatus(desired, null);
            return;
        }

        if (!_backgroundPreferences.Save(desired, out var error))
        {
            RevertBackgroundCheckBox(previous);
            UpdateBackgroundSettingStatus(previous, $"設定保存失敗: {error}");
            return;
        }

        // If login startup is already enabled, rewrite the registration so its
        // command-line arguments match the newly persisted background preference.
        if (_viewModel.AutoStart && !_startupRegistration.SetEnabled(true, out var startupError))
        {
            _backgroundPreferences.Save(previous, out _);
            RevertBackgroundCheckBox(previous);
            UpdateBackgroundSettingStatus(previous, $"自動起動設定の更新失敗: {startupError}");
            return;
        }

        UpdateBackgroundSettingStatus(desired, null);
    }

    private void LoadBackgroundPreference()
    {
        var enabled = _backgroundPreferences.Load().StartInBackground;
        RevertBackgroundCheckBox(enabled);
        UpdateBackgroundSettingStatus(enabled, null);
    }

    private void RevertBackgroundCheckBox(bool value)
    {
        _changingBackgroundPreference = true;
        StartInBackgroundCheckBox.IsChecked = value;
        _changingBackgroundPreference = false;
    }

    private void UpdateBackgroundSettingStatus(bool enabled, string? error)
    {
        BackgroundSettingStatus.Text = error ?? (enabled
            ? "次回起動は常駐状態から開始します。SD検出時は画面を表示します。"
            : "次回起動はメイン画面を表示します。");
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(MainWindowViewModel.DestinationPath) or nameof(MainWindowViewModel.EventName))
        {
            UpdateDestinationPreview();
        }
    }

    private void UpdateDestinationPreview()
    {
        if (_viewModel is null) return;

        var destination = string.IsNullOrWhiteSpace(_viewModel.DestinationPath)
            ? "[保存先]"
            : _viewModel.DestinationPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var eventName = ImportCoordinator.SanitizeEventName(_viewModel.EventName.Trim());
        if (string.IsNullOrWhiteSpace(eventName)) eventName = "[イベント名]";

        var separator = Path.DirectorySeparatorChar;
        DestinationPreviewText.Text =
            $"{destination}{separator}YYYY{separator}YYYY-MM-DD_{eventName}{separator}[RAW|JPG|MP4]";
    }
}
