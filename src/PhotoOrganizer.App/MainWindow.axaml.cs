using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;

namespace PhotoOrganizer.App;

public sealed partial class MainWindow : Window
{
    private readonly MainWindowViewModel? _viewModel;

    public MainWindow()
    {
        InitializeComponent();
    }

    public MainWindow(MainWindowViewModel viewModel) : this()
    {
        _viewModel = viewModel;
        DataContext = viewModel;
        viewModel.RequestShowWindow += ShowAndActivate;

        Opened += OnOpened;
        Closing += OnClosing;
        Closed += OnClosed;
    }

    private async void OnOpened(object? sender, EventArgs e)
    {
        if (_viewModel is null) return;
        await _viewModel.InitializeAsync();
    }

    private void OnClosing(object? sender, WindowClosingEventArgs e)
    {
        if (_viewModel is null) return;
        if (!_viewModel.CanCloseWindow()) e.Cancel = true;
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        if (_viewModel is null) return;
        _viewModel.RequestShowWindow -= ShowAndActivate;
        _viewModel.Dispose();
    }

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

    private void ShowAndActivate()
    {
        if (!IsVisible) Show();
        WindowState = WindowState.Normal;
        Activate();
    }
}
