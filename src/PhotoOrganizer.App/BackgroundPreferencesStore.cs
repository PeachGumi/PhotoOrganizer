using System.Text.Json;

namespace PhotoOrganizer.App;

public sealed record BackgroundPreferences(bool StartInBackground)
{
    public static BackgroundPreferences Default => new(false);
}

public sealed class BackgroundPreferencesStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path;

    public BackgroundPreferencesStore()
    {
        var directory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PhotoOrganizer");
        _path = Path.Combine(directory, "background-settings.json");
    }

    public BackgroundPreferences Load()
    {
        try
        {
            if (!File.Exists(_path)) return BackgroundPreferences.Default;
            return JsonSerializer.Deserialize<BackgroundPreferences>(File.ReadAllText(_path))
                ?? BackgroundPreferences.Default;
        }
        catch
        {
            return BackgroundPreferences.Default;
        }
    }

    public bool Save(bool startInBackground, out string? error)
    {
        try
        {
            var directory = Path.GetDirectoryName(_path)!;
            Directory.CreateDirectory(directory);
            var temporary = _path + ".tmp-" + Guid.NewGuid().ToString("N");
            File.WriteAllText(
                temporary,
                JsonSerializer.Serialize(new BackgroundPreferences(startInBackground), JsonOptions));
            File.Move(temporary, _path, overwrite: true);
            error = null;
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }
}
