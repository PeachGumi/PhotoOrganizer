namespace PhotoOrganizer.App;

public sealed record BackgroundPreferences(bool StartInBackground)
{
    public static BackgroundPreferences Default => new(false);
}

public sealed class BackgroundPreferencesStore
{
    private readonly string _path = JsonSettingsFile.GetPath("background-settings.json");

    public BackgroundPreferences Load() =>
        JsonSettingsFile.Load(_path, BackgroundPreferences.Default);

    public bool Save(bool startInBackground, out string? error) =>
        JsonSettingsFile.Save(_path, new BackgroundPreferences(startInBackground), out error);
}
