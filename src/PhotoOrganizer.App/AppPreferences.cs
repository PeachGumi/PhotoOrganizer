using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public sealed record AppPreferences(
    string DestinationPath,
    string[] RawExtensions)
{
    public static AppPreferences Default => new(
        Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
        MediaClassifier.DefaultRawExtensions.ToArray());
}

public sealed class AppPreferencesStore
{
    private readonly string _path = JsonSettingsFile.GetPath("settings.json");

    public AppPreferences Load()
    {
        var loaded = JsonSettingsFile.Load(_path, AppPreferences.Default);
        var defaults = AppPreferences.Default;
        var destination = string.IsNullOrWhiteSpace(loaded.DestinationPath)
            ? defaults.DestinationPath
            : loaded.DestinationPath;
        var raw = loaded.RawExtensions is { Length: > 0 }
            ? loaded.RawExtensions
            : defaults.RawExtensions;
        return loaded with { DestinationPath = destination, RawExtensions = raw };
    }

    public bool Save(AppPreferences preferences, out string? error) =>
        JsonSettingsFile.Save(_path, preferences, out error);
}
