using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class JsonSettingsFileTests
{
    [TestMethod]
    public void Load_MissingFileReturnsFallback()
    {
        using var temp = new TempDirectory();
        var fallback = new TestSettings("fallback", 1);

        var loaded = JsonSettingsFile.Load(Path.Combine(temp.Path, "missing.json"), fallback);

        Assert.AreEqual(fallback, loaded);
    }

    [TestMethod]
    public void SaveAndLoad_RoundTripsSettings()
    {
        using var temp = new TempDirectory();
        var path = Path.Combine(temp.Path, "nested", "settings.json");
        var expected = new TestSettings("value", 42);

        var saved = JsonSettingsFile.Save(path, expected, out var error);
        var loaded = JsonSettingsFile.Load(path, new TestSettings("fallback", 0));

        Assert.IsTrue(saved, error);
        Assert.AreEqual(expected, loaded);
        Assert.IsFalse(Directory.EnumerateFiles(Path.GetDirectoryName(path)!, "*.tmp-*", SearchOption.TopDirectoryOnly).Any());
    }

    private sealed record TestSettings(string Name, int Count);

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PhotoOrganizerSettingsTests-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
