using System.Text.Json;
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
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path;

    public AppPreferencesStore()
    {
        var directory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PhotoOrganizer");
        _path = Path.Combine(directory, "settings.json");
    }

    public AppPreferences Load()
    {
        try
        {
            if (!File.Exists(_path)) return AppPreferences.Default;
            var loaded = JsonSerializer.Deserialize<AppPreferences>(File.ReadAllText(_path));
            if (loaded is null) return AppPreferences.Default;

            var destination = string.IsNullOrWhiteSpace(loaded.DestinationPath)
                ? AppPreferences.Default.DestinationPath
                : loaded.DestinationPath;
            var raw = loaded.RawExtensions is { Length: > 0 }
                ? loaded.RawExtensions
                : AppPreferences.Default.RawExtensions;
            return loaded with { DestinationPath = destination, RawExtensions = raw };
        }
        catch
        {
            return AppPreferences.Default;
        }
    }

    public bool Save(AppPreferences preferences, out string? error)
    {
        try
        {
            var directory = Path.GetDirectoryName(_path)!;
            Directory.CreateDirectory(directory);
            var temporary = _path + ".tmp-" + Guid.NewGuid().ToString("N");
            File.WriteAllText(temporary, JsonSerializer.Serialize(preferences, JsonOptions));
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
