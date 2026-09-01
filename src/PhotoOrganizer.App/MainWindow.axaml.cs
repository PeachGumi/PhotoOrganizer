using System.ComponentModel;
using System.Diagnostics;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed partial class MainWindow : Window
{
    private readonly MainWindowViewModel? _viewModel;
    private bool _allowExplicitClose;

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
        UpdateDestinationPreview();
    }

    private void OnClosing(object? sender, WindowClosingEventArgs e)
    {
        if (_allowExplicitClose) return;

        e.Cancel = true;
        if (_viewModel is not null && !_viewModel.CanCloseWindow()) return;
        Hide();
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        if (_viewModel is not null) _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
    }

    internal void AllowExplicitClose() => _allowExplicitClose = true;

    private void Settings_Click(object? sender, RoutedEventArgs e)
    {
        if (Application.Current is App app) app.ShowSettingsWindow();
    }

    private void Diagnostics_Click(object? sender, RoutedEventArgs e)
    {
        if (Application.Current is App app) app.ShowDiagnosticsWindow();
    }

    private async void SelectDestination_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            if (_viewModel is null || !StorageProvider.CanPickFolder) return;

            var folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "写真の保存先を選択",
                AllowMultiple = false
            });

            var path = folders.FirstOrDefault()?.TryGetLocalPath();
            if (!string.IsNullOrWhiteSpace(path))
            {
                _viewModel.SetDestinationFromPicker(path);
                await _viewModel.ValidateDestinationAsync();
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception exception)
        {
            ReportUiFailure("保存先選択", exception);
        }
    }

    private async void Destination_LostFocus(object? sender, RoutedEventArgs e)
    {
        try
        {
            if (_viewModel is not null) await _viewModel.ValidateDestinationAsync();
        }
        catch (Exception exception)
        {
            ReportUiFailure("保存先確認", exception);
        }
    }

    private async void SelectSd_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            if (_viewModel is null || !StorageProvider.CanPickFolder) return;

            var folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "SDカードまたはカード内のフォルダを選択",
                AllowMultiple = false
            });

            var path = folders.FirstOrDefault()?.TryGetLocalPath();
            if (!string.IsNullOrWhiteSpace(path)) await _viewModel.ScanCardAsync(path);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception exception)
        {
            ReportUiFailure("SDカード選択", exception);
        }
    }

    private async void Import_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            if (_viewModel is not null) await _viewModel.StartImportAsync();
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception exception)
        {
            ReportUiFailure("取り込み", exception);
        }
    }

    private void Cancel_Click(object? sender, RoutedEventArgs e)
    {
        _viewModel?.CancelImport();
    }

    private void OpenDestination_Click(object? sender, RoutedEventArgs e)
    {
        if (_viewModel is null || string.IsNullOrWhiteSpace(_viewModel.LastImportBasePath)) return;

        try
        {
            if (!Directory.Exists(_viewModel.LastImportBasePath))
            {
                throw new DirectoryNotFoundException("取り込み先フォルダが見つかりません。");
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = _viewModel.LastImportBasePath,
                UseShellExecute = true
            });
        }
        catch (Exception exception)
        {
            _viewModel.ReportAuxiliaryFailure("保存先を開く", exception);
        }
    }

    private async void EjectSd_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            if (_viewModel is not null) await _viewModel.EjectSelectedSdAsync();
        }
        catch (Exception exception)
        {
            ReportUiFailure("SDカード取り出し", exception);
        }
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(MainWindowViewModel.DestinationPath) or nameof(MainWindowViewModel.EventName))
        {
            try
            {
                UpdateDestinationPreview();
            }
            catch (Exception exception)
            {
                ReportUiFailure("保存先プレビュー更新", exception);
            }
        }
    }

    private void ReportUiFailure(string operation, Exception exception)
    {
        if (_viewModel is not null)
        {
            _viewModel.ReportUiFailure(operation, exception);
        }
        else
        {
            Trace.WriteLine($"Photo Organizer {operation} failed before view model initialization: {exception}");
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
