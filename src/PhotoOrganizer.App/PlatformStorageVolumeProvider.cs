using System.Diagnostics;
using System.Globalization;
using System.Management;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed class PlatformStorageVolumeProvider : IStorageVolumeProvider
{
    public StringComparison PathComparison => OperatingSystem.IsWindows()
        ? StringComparison.OrdinalIgnoreCase
        : StringComparison.Ordinal;

    public IReadOnlyList<MountedVolumeInfo> GetMountedVolumes()
    {
        var systemRoot = OperatingSystem.IsWindows()
            ? Path.GetPathRoot(Environment.SystemDirectory)
            : "/";

        DriveInfo[] drives;
        try
        {
            drives = DriveInfo.GetDrives();
        }
        catch
        {
            // Enumeration failure cannot be allowed to manufacture or preserve an
            // unverified identity. Callers treat an empty set as fail-closed.
            return [];
        }

        var volumes = new List<MountedVolumeInfo>();
        foreach (var drive in drives)
        {
            try
            {
                if (!drive.IsReady) continue;
                var root = PathSafety.Normalize(drive.RootDirectory.FullName);
                if (OperatingSystem.IsMacOS()
                    && root != "/"
                    && !root.StartsWith("/Volumes/", StringComparison.Ordinal))
                {
                    continue;
                }

                string? fingerprint;
                string? physicalDeviceFingerprint;

                if (OperatingSystem.IsWindows())
                {
                    fingerprint = TryGetWindowsVolumeGuid(root);
                    physicalDeviceFingerprint = TryGetWindowsPhysicalDiskFingerprint(root);
                }
                else if (OperatingSystem.IsMacOS())
                {
                    // diskutil exposes both persistent volume identity and the
                    // containing whole-disk identity. Resolve them together so a safety
                    // refresh never launches a second subprocess for the same volume.
                    var identity = TryGetMacStorageIdentity(root);
                    fingerprint = identity?.VolumeFingerprint;
                    physicalDeviceFingerprint = identity?.PhysicalDeviceFingerprint;
                }
                else
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(fingerprint)) continue;

                var isSystem = !string.IsNullOrWhiteSpace(systemRoot)
                    && string.Equals(PathSafety.Normalize(systemRoot), root, PathComparison);
                var isRemovable = drive.DriveType == DriveType.Removable
                    || (OperatingSystem.IsMacOS() && root.StartsWith("/Volumes/", StringComparison.Ordinal));

                volumes.Add(new MountedVolumeInfo(
                    root,
                    fingerprint,
                    isRemovable,
                    isSystem,
                    physicalDeviceFingerprint));
            }
            catch
            {
                // Identity enumeration is fail-closed. Unreadable volumes are not candidates.
            }
        }

        return volumes;
    }

    public MountedVolumeInfo? ResolveVolumeForPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;

        string current;
        try
        {
            current = PathSafety.Normalize(path);
        }
        catch
        {
            return null;
        }

        while (!File.Exists(current) && !Directory.Exists(current))
        {
            var parent = Directory.GetParent(current)?.FullName;
            if (string.IsNullOrWhiteSpace(parent) || string.Equals(parent, current, PathComparison)) break;
            current = parent;
        }

        if (!File.Exists(current) && !Directory.Exists(current)) return null;

        return GetMountedVolumes()
            .Where(v => PathSafety.IsSameOrDescendant(current, v.RootPath, PathComparison))
            .OrderByDescending(v => PathSafety.Normalize(v.RootPath).Length)
            .FirstOrDefault();
    }

    private static string? TryGetWindowsVolumeGuid(string root)
    {
        var mountPoint = root.EndsWith('\\') ? root : root + "\\";
        var buffer = new StringBuilder(1024);
        return GetVolumeNameForVolumeMountPoint(mountPoint, buffer, (uint)buffer.Capacity)
            ? "windows-volume:" + buffer.ToString()
            : null;
    }

    [SupportedOSPlatform("windows")]
    private static string? TryGetWindowsPhysicalDiskFingerprint(string root)
    {
        try
        {
            var logicalDeviceId = Path.GetPathRoot(root)?.TrimEnd('\\');
            if (string.IsNullOrWhiteSpace(logicalDeviceId) || logicalDeviceId.Length != 2 || logicalDeviceId[1] != ':')
            {
                return null;
            }

            var query = $"ASSOCIATORS OF {{Win32_LogicalDisk.DeviceID='{logicalDeviceId}'}} WHERE AssocClass=Win32_LogicalDiskToPartition";
            using var searcher = new ManagementObjectSearcher("root\\CIMV2", query);
            using var results = searcher.Get();
            var diskIndices = new SortedSet<uint>();

            foreach (ManagementObject partition in results)
            {
                using (partition)
                {
                    var diskIndex = partition["DiskIndex"];
                    if (diskIndex is null) continue;
                    diskIndices.Add(Convert.ToUInt32(diskIndex, CultureInfo.InvariantCulture));
                }
            }

            // A volume backed by multiple physical disks cannot be represented by
            // one unambiguous physical-device identity. Fail closed rather than risk
            // accepting a destination that overlaps the camera-card disk.
            if (diskIndices.Count != 1) return null;
            return $"windows-physical-disk:{diskIndices.Min}";
        }
        catch
        {
            // Physical identity is a safety signal. Callers fail closed when unavailable.
        }

        return null;
    }

    [SupportedOSPlatform("macos")]
    private static MacStorageIdentity? TryGetMacStorageIdentity(string root)
    {
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "/usr/sbin/diskutil",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };
            process.StartInfo.ArgumentList.Add("info");
            process.StartInfo.ArgumentList.Add("-plist");
            process.StartInfo.ArgumentList.Add(root);

            if (!process.Start()) return null;

            // Begin draining both redirected pipes before waiting. Reading either
            // pipe synchronously first can deadlock if the child blocks or fills the
            // other pipe, which would make the nominal timeout ineffective.
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(5000))
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                try { process.WaitForExit(1000); } catch { }
                return null;
            }

            var output = stdout.GetAwaiter().GetResult();
            _ = stderr.GetAwaiter().GetResult();
            if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output)) return null;

            return MacDiskutilInfoParser.Parse(output);
        }
        catch
        {
            // Storage identity is a safety signal. Callers fail closed when unavailable.
            return null;
        }
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetVolumeNameForVolumeMountPoint(
        string lpszVolumeMountPoint,
        StringBuilder lpszVolumeName,
        uint cchBufferLength);
}
