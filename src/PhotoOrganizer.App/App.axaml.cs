using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed partial class App : Application
{
    private StorageMonitor? _storageMonitor;

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
            _storageMonitor = new StorageMonitor(StorageSessions);
            _storageMonitor.Start();

            var viewModel = new MainWindowViewModel(
                StorageProvider,
                StorageSessions,
                CameraCardRoots);

            desktop.MainWindow = new MainWindow(viewModel);
            desktop.Exit += (_, _) => _storageMonitor?.Dispose();
        }

        base.OnFrameworkInitializationCompleted();
    }
}
