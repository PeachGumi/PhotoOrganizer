using PhotoOrganizer.App;

namespace PhotoOrganizer.App.Tests;

internal sealed class TestPreferencesStore : IAppPreferencesStore
{
    public TestPreferencesStore(AppPreferences? initial = null)
    {
        Preferences = initial ?? AppPreferences.Default;
    }

    public AppPreferences Preferences { get; private set; }

    public AppPreferences Load() => Preferences;

    public bool Save(AppPreferences preferences, out string? error)
    {
        Preferences = preferences;
        error = null;
        return true;
    }
}
