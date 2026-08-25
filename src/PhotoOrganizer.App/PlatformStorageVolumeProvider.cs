using System.Diagnostics;
using System.Runtime.InteropServices;
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

        var volumes = new List<MountedVolumeInfo>();
        foreach (var drive in DriveInfo.GetDrives())
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

                var fingerprint = TryGetFingerprint(root);
                if (string.IsNullOrWhiteSpace(fingerprint)) continue;

                var isSystem = !string.IsNullOrWhiteSpace(systemRoot)
                    && string.Equals(PathSafety.Normalize(systemRoot), root, PathComparison);
                var isRemovable = drive.DriveType == DriveType.Removable
                    || (OperatingSystem.IsMacOS() && root.StartsWith("/Volumes/", StringComparison.Ordinal));

                volumes.Add(new MountedVolumeInfo(root, fingerprint, isRemovable, isSystem));
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

    private static string? TryGetFingerprint(string root)
    {
        if (OperatingSystem.IsWindows()) return TryGetWindowsVolumeGuid(root);
        if (OperatingSystem.IsMacOS()) return TryGetMacDeviceFingerprint(root);
        return null;
    }

    private static string? TryGetWindowsVolumeGuid(string root)
    {
        var mountPoint = root.EndsWith('\\') ? root : root + "\\";
        var buffer = new StringBuilder(1024);
        return GetVolumeNameForVolumeMountPoint(mountPoint, buffer, (uint)buffer.Capacity)
            ? "windows-volume:" + buffer.ToString()
            : null;
    }

    private static string? TryGetMacDeviceFingerprint(string root)
    {
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "/usr/bin/stat",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };
            process.StartInfo.ArgumentList.Add("-f");
            process.StartInfo.ArgumentList.Add("%d");
            process.StartInfo.ArgumentList.Add(root);

            if (!process.Start()) return null;
            var output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit(3000);
            if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output)) return null;
            return "mac-device:" + output;
        }
        catch
        {
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
