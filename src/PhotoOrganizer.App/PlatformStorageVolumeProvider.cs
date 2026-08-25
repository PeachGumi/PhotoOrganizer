using System.Diagnostics;
using System.Globalization;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
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

                var physicalDeviceFingerprint = TryGetPhysicalDeviceFingerprint(root);
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

    private static string? TryGetFingerprint(string root)
    {
        if (OperatingSystem.IsWindows()) return TryGetWindowsVolumeGuid(root);
        if (OperatingSystem.IsMacOS()) return TryGetMacDeviceFingerprint(root);
        return null;
    }

    private static string? TryGetPhysicalDeviceFingerprint(string root)
    {
        if (OperatingSystem.IsWindows()) return TryGetWindowsPhysicalDiskFingerprint(root);
        if (OperatingSystem.IsMacOS()) return TryGetMacWholeDiskFingerprint(root);
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

            foreach (ManagementObject partition in results)
            {
                using (partition)
                {
                    var diskIndex = partition["DiskIndex"];
                    if (diskIndex is null) continue;
                    var value = Convert.ToUInt32(diskIndex, CultureInfo.InvariantCulture);
                    return $"windows-physical-disk:{value}";
                }
            }
        }
        catch
        {
            // Physical identity is a safety signal. Callers fail closed when unavailable.
        }

        return null;
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

    private static string? TryGetMacWholeDiskFingerprint(string root)
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
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(5000))
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                return null;
            }

            var output = stdout.GetAwaiter().GetResult();
            _ = stderr.GetAwaiter().GetResult();
            if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output)) return null;

            var document = XDocument.Parse(output, LoadOptions.None);
            var dictionary = document.Root?.Elements().FirstOrDefault(element => element.Name.LocalName == "dict");
            if (dictionary is null) return null;

            var parentWholeDisk = GetPlistString(dictionary, "ParentWholeDisk");
            if (!string.IsNullOrWhiteSpace(parentWholeDisk))
            {
                return "mac-whole-disk:" + parentWholeDisk;
            }

            if (GetPlistBoolean(dictionary, "WholeDisk") == true)
            {
                var deviceIdentifier = GetPlistString(dictionary, "DeviceIdentifier");
                if (!string.IsNullOrWhiteSpace(deviceIdentifier))
                {
                    return "mac-whole-disk:" + deviceIdentifier;
                }
            }
        }
        catch
        {
            // Physical identity is a safety signal. Callers fail closed when unavailable.
        }

        return null;
    }

    private static string? GetPlistString(XElement dictionary, string key)
    {
        var elements = dictionary.Elements().ToArray();
        for (var index = 0; index + 1 < elements.Length; index++)
        {
            if (elements[index].Name.LocalName != "key" || elements[index].Value != key) continue;
            return elements[index + 1].Name.LocalName == "string"
                ? elements[index + 1].Value
                : null;
        }

        return null;
    }

    private static bool? GetPlistBoolean(XElement dictionary, string key)
    {
        var elements = dictionary.Elements().ToArray();
        for (var index = 0; index + 1 < elements.Length; index++)
        {
            if (elements[index].Name.LocalName != "key" || elements[index].Value != key) continue;
            return elements[index + 1].Name.LocalName switch
            {
                "true" => true,
                "false" => false,
                _ => null
            };
        }

        return null;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetVolumeNameForVolumeMountPoint(
        string lpszVolumeMountPoint,
        StringBuilder lpszVolumeName,
        uint cchBufferLength);
}
