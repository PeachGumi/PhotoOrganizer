using System.Runtime.InteropServices;
using System.Runtime.Versioning;
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
            var startInBackground = enabled && new BackgroundPreferencesStore().Load().StartInBackground;
            if (OperatingSystem.IsWindows()) return SetWindows(enabled, startInBackground, out error);
            if (OperatingSystem.IsMacOS()) return SetMac(enabled, startInBackground, out error);
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
    private static bool SetWindows(bool enabled, bool startInBackground, out string? error)
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

        var command = StartupRegistrationFormatter.BuildWindowsCommand(executable, startInBackground);
        key.SetValue(WindowsValueName, command, RegistryValueKind.String);
        error = null;
        return true;
    }

    [SupportedOSPlatform("macos")]
    private static bool SetMac(bool enabled, bool startInBackground, out string? error)
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
        var plist = StartupRegistrationFormatter.BuildMacLaunchAgentPlist(
            MacLabel,
            executable,
            startInBackground);

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
        _ = BoundedProcessRunner.Run(
            "/bin/launchctl",
            arguments,
            TimeSpan.FromSeconds(3));
        // Failure is intentionally non-fatal here: the LaunchAgent file still controls
        // the next-login state, while the settings UI will report the persisted state.
    }

    [DllImport("libSystem.B.dylib")]
    [SupportedOSPlatform("macos")]
    private static extern uint geteuid();
}
