using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Threading;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed partial class App : Application
{
    private StorageMonitor? _storageMonitor;
    private IClassicDesktopStyleApplicationLifetime? _desktop;
    private MainWindow? _mainWindow;
    private MainWindowViewModel? _viewModel;
    private bool _explicitQuitRequested;

    public PlatformStorageVolumeProvider StorageProvider { get; } = new();
    public StorageSessionTracker StorageSessions { get; }
    public CameraCardRootResolver CameraCardRoots { get; }

    public App()
    {
        StorageSessions = new StorageSessionTracker(StorageProvider);
        CameraCardRoots = new CameraCardRootResolver(StorageProvider);
    }

    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            _desktop = desktop;
            desktop.ShutdownMode = ShutdownMode.OnExplicitShutdown;

            _storageMonitor = new StorageMonitor(StorageSessions);
            _storageMonitor.Start();

            _viewModel = new MainWindowViewModel(
                StorageProvider,
                StorageSessions,
                CameraCardRoots);
            _viewModel.RequestShowWindow += ShowMainWindow;

            _mainWindow = new MainWindow(_viewModel);

            var backgroundPreferences = new BackgroundPreferencesStore().Load();
            var requestedBackground = desktop.Args.Any(
                argument => string.Equals(argument, "--background", StringComparison.OrdinalIgnoreCase));
            var startHidden = requestedBackground || backgroundPreferences.StartInBackground;

            if (!startHidden)
            {
                desktop.MainWindow = _mainWindow;
            }

            desktop.ShutdownRequested += OnShutdownRequested;
            desktop.Exit += OnExit;

            // Initial card discovery must also run when the application starts hidden.
            // A detected camera card will request that the workflow window be shown.
            Dispatcher.UIThread.Post(() => _ = _viewModel.InitializeAsync());
        }

        base.OnFrameworkInitializationCompleted();
    }

    private void TrayIcon_Clicked(object? sender, EventArgs e) => ShowMainWindow();

    private void ShowWindowMenu_Click(object? sender, EventArgs e) => ShowMainWindow();

    private void QuitMenu_Click(object? sender, EventArgs e) => RequestQuit();

    private void ShowMainWindow()
    {
        if (_desktop is null || _mainWindow is null) return;

        if (_desktop.MainWindow is null)
        {
            _desktop.MainWindow = _mainWindow;
        }

        if (!_mainWindow.IsVisible) _mainWindow.Show();
        _mainWindow.WindowState = WindowState.Normal;
        _mainWindow.Activate();
    }

    private void RequestQuit()
    {
        if (_desktop is null) return;

        if (_viewModel?.IsProcessing == true)
        {
            _viewModel.CanCloseWindow();
            ShowMainWindow();
            return;
        }

        _explicitQuitRequested = true;
        _mainWindow?.AllowExplicitClose();
        _desktop.Shutdown();
    }

    private void OnShutdownRequested(object? sender, ShutdownRequestedEventArgs e)
    {
        if (_explicitQuitRequested || _viewModel?.IsProcessing != true) return;

        // A normal quit/logout request must never interrupt an active import and then
        // leave the user with a stale reuse approval. Forced process termination can
        // still happen at the OS level, but safety approval is memory-only and is not
        // restored after restart.
        e.Cancel = true;
        _viewModel.CanCloseWindow();
        ShowMainWindow();
    }

    private void OnExit(object? sender, ControlledApplicationLifetimeExitEventArgs e)
    {
        if (_viewModel is not null)
        {
            _viewModel.RequestShowWindow -= ShowMainWindow;
            _viewModel.Dispose();
            _viewModel = null;
        }

        _storageMonitor?.Dispose();
        _storageMonitor = null;
    }
}
