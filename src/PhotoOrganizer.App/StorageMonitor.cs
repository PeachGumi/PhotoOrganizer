using System.Management;
using System.Runtime.Versioning;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed class StorageMonitor : IDisposable
{
    private readonly StorageSessionTracker _tracker;
    private readonly object _refreshGate = new();
    private Timer? _pollTimer;
    private FileSystemWatcher? _macVolumesWatcher;
    private ManagementEventWatcher? _windowsWatcher;
    private bool _started;

    public StorageMonitor(StorageSessionTracker tracker)
    {
        _tracker = tracker;
    }

    public void Start()
    {
        if (_started) return;
        _started = true;

        SafeRefresh();
        _pollTimer = new Timer(_ => SafeRefresh(), null, TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(1));

        if (OperatingSystem.IsWindows())
        {
            StartWindowsWatcher();
        }
        else if (OperatingSystem.IsMacOS())
        {
            StartMacWatcher();
        }
    }

    [SupportedOSPlatform("windows")]
    private void StartWindowsWatcher()
    {
        try
        {
            _windowsWatcher = new ManagementEventWatcher(
                new WqlEventQuery("SELECT * FROM Win32_VolumeChangeEvent"));
            _windowsWatcher.EventArrived += OnWindowsVolumeChanged;
            _windowsWatcher.Start();
        }
        catch
        {
            _windowsWatcher?.Dispose();
            _windowsWatcher = null;
            // The periodic poll remains active as the fallback detector.
        }
    }

    [SupportedOSPlatform("windows")]
    private void OnWindowsVolumeChanged(object sender, EventArrivedEventArgs e)
    {
        try
        {
            var eventType = Convert.ToUInt16(e.NewEvent.Properties["EventType"]?.Value ?? 0);
            var driveName = e.NewEvent.Properties["DriveName"]?.Value?.ToString();
            if (eventType == 3 && !string.IsNullOrWhiteSpace(driveName))
            {
                _tracker.MarkRemoved(PathSafety.Normalize(driveName));
            }
        }
        catch
        {
            // Refresh below is still fail-closed.
        }

        SafeRefresh();
    }

    private void StartMacWatcher()
    {
        try
        {
            if (!Directory.Exists("/Volumes")) return;
            _macVolumesWatcher = new FileSystemWatcher("/Volumes")
            {
                IncludeSubdirectories = false,
                NotifyFilter = NotifyFilters.DirectoryName,
                EnableRaisingEvents = true
            };
            _macVolumesWatcher.Created += OnMacVolumeChanged;
            _macVolumesWatcher.Changed += OnMacVolumeChanged;
            _macVolumesWatcher.Deleted += OnMacVolumeRemoved;
            _macVolumesWatcher.Renamed += OnMacVolumeRenamed;
            _macVolumesWatcher.Error += (_, _) => SafeRefresh();
        }
        catch
        {
            _macVolumesWatcher?.Dispose();
            _macVolumesWatcher = null;
        }
    }

    private void OnMacVolumeChanged(object sender, FileSystemEventArgs e) => SafeRefresh();

    private void OnMacVolumeRemoved(object sender, FileSystemEventArgs e)
    {
        _tracker.MarkRemoved(e.FullPath);
        SafeRefresh();
    }

    private void OnMacVolumeRenamed(object sender, RenamedEventArgs e)
    {
        _tracker.MarkRemoved(e.OldFullPath);
        SafeRefresh();
    }

    private void SafeRefresh()
    {
        if (!Monitor.TryEnter(_refreshGate)) return;
        try
        {
            _tracker.Refresh();
        }
        catch
        {
            // A failed enumeration cannot create a new safe identity. The next event/poll retries.
        }
        finally
        {
            Monitor.Exit(_refreshGate);
        }
    }

    public void Dispose()
    {
        _started = false;
        _pollTimer?.Dispose();
        _pollTimer = null;

        if (_windowsWatcher is not null)
        {
            try { _windowsWatcher.Stop(); } catch { }
            _windowsWatcher.Dispose();
            _windowsWatcher = null;
        }

        _macVolumesWatcher?.Dispose();
        _macVolumesWatcher = null;
    }
}
