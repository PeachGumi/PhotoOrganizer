using System.Diagnostics;
using System.Runtime.Versioning;
using System.Security;
using Microsoft.Win32;

namespace PhotoOrganizer.App;

public interface IStartupRegistrationService
{
    bool IsSupported { get; }
    bool IsEnabled();
    bool SetEnabled(bool enabled, out string? error);
}

public sealed class StartupRegistrationService : IStartupRegistrationService
{
    private const string WindowsRunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string WindowsValueName = "PhotoOrganizer";
    private const string MacLabel = "com.peachgumi.photoorganizer";

    public bool IsSupported => OperatingSystem.IsWindows() || OperatingSystem.IsMacOS();

    public bool IsEnabled()
    {
        try
        {
            if (OperatingSystem.IsWindows()) return IsWindowsEnabled();
            if (OperatingSystem.IsMacOS()) return File.Exists(GetMacPlistPath());
        }
        catch
        {
            // Settings UI must fail closed rather than claiming startup is enabled.
        }

        return false;
    }

    public bool SetEnabled(bool enabled, out string? error)
    {
        try
        {
            if (OperatingSystem.IsWindows()) return SetWindows(enabled, out error);
            if (OperatingSystem.IsMacOS()) return SetMac(enabled, out error);
            error = "Login startup is not supported on this operating system.";
            return false;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    [SupportedOSPlatform("windows")]
    private static bool IsWindowsEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(WindowsRunKey, writable: false);
        return key?.GetValue(WindowsValueName) is string value && !string.IsNullOrWhiteSpace(value);
    }

    [SupportedOSPlatform("windows")]
    private static bool SetWindows(bool enabled, out string? error)
    {
        using var key = Registry.CurrentUser.CreateSubKey(WindowsRunKey, writable: true);
        if (key is null)
        {
            error = "Unable to open the Windows login startup registry key.";
            return false;
        }

        if (!enabled)
        {
            key.DeleteValue(WindowsValueName, throwOnMissingValue: false);
            error = null;
            return true;
        }

        var executable = GetRunnableExecutable();
        if (executable is null)
        {
            error = "The packaged application executable could not be identified.";
            return false;
        }

        key.SetValue(WindowsValueName, $"\"{executable}\"", RegistryValueKind.String);
        error = null;
        return true;
    }

    [SupportedOSPlatform("macos")]
    private static bool SetMac(bool enabled, out string? error)
    {
        var plistPath = GetMacPlistPath();
        if (!enabled)
        {
            TryBootout(plistPath);
            if (File.Exists(plistPath)) File.Delete(plistPath);
            error = null;
            return true;
        }

        var executable = GetRunnableExecutable();
        if (executable is null)
        {
            error = "The packaged application executable could not be identified.";
            return false;
        }

        var directory = Path.GetDirectoryName(plistPath)!;
        Directory.CreateDirectory(directory);
        var escapedExecutable = SecurityElement.Escape(executable) ?? executable;
        var plist = $"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>{MacLabel}</string>
  <key>ProgramArguments</key>
  <array>
    <string>{escapedExecutable}</string>
  </array>
  <key>RunAtLoad</key><true/>
</dict>
</plist>
""";

        var temporary = plistPath + ".tmp-" + Guid.NewGuid().ToString("N");
        File.WriteAllText(temporary, plist);
        File.Move(temporary, plistPath, overwrite: true);
        TryBootstrap(plistPath);
        error = null;
        return true;
    }

    private static string? GetRunnableExecutable()
    {
        var executable = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(executable) || !File.Exists(executable)) return null;
        if (string.Equals(Path.GetFileNameWithoutExtension(executable), "dotnet", StringComparison.OrdinalIgnoreCase)) return null;
        return executable;
    }

    private static string GetMacPlistPath() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "Library",
        "LaunchAgents",
        MacLabel + ".plist");

    [SupportedOSPlatform("macos")]
    private static void TryBootstrap(string plistPath) => RunLaunchctl("bootstrap", $"gui/{geteuid()}", plistPath);

    [SupportedOSPlatform("macos")]
    private static void TryBootout(string plistPath) => RunLaunchctl("bootout", $"gui/{geteuid()}", plistPath);

    [SupportedOSPlatform("macos")]
    private static void RunLaunchctl(params string[] arguments)
    {
        if (!File.Exists("/bin/launchctl")) return;
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "/bin/launchctl",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };
            foreach (var argument in arguments) process.StartInfo.ArgumentList.Add(argument);
            if (!process.Start()) return;
            process.WaitForExit(3000);
        }
        catch
        {
            // The LaunchAgent file still controls the next-login state.
        }
    }

    [DllImport("libSystem.B.dylib")]
    [SupportedOSPlatform("macos")]
    private static extern uint geteuid();
}
