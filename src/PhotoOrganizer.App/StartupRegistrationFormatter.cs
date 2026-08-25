using System.Security;

namespace PhotoOrganizer.App;

internal static class StartupRegistrationFormatter
{
    public static string BuildWindowsCommand(string executable, bool startInBackground)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(executable);
        return $"\"{executable}\"" + (startInBackground ? " --background" : string.Empty);
    }

    public static string BuildMacLaunchAgentPlist(
        string label,
        string executable,
        bool startInBackground)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(label);
        ArgumentException.ThrowIfNullOrWhiteSpace(executable);

        var escapedLabel = SecurityElement.Escape(label) ?? label;
        var escapedExecutable = SecurityElement.Escape(executable) ?? executable;
        var arguments = new List<string>
        {
            $"    <string>{escapedExecutable}</string>"
        };
        if (startInBackground) arguments.Add("    <string>--background</string>");

        return string.Join('\n',
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
            "<!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">",
            "<plist version=\"1.0\">",
            "<dict>",
            $"  <key>Label</key><string>{escapedLabel}</string>",
            "  <key>ProgramArguments</key>",
            "  <array>",
            string.Join('\n', arguments),
            "  </array>",
            "  <key>RunAtLoad</key><true/>",
            "</dict>",
            "</plist>",
            string.Empty);
    }
}
