using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class MediaDateResolverTests
{
    [TestMethod]
    public void ResolveDateSummary_EnumeratesFilesOnceAndReturnsEarliestDistinctKeys()
    {
        using var temp = new TempDirectory();
        var first = AddFile(temp.Path, "first.jpg", new DateTime(2026, 1, 3, 10, 0, 0));
        var second = AddFile(temp.Path, "second.jpg", new DateTime(2026, 1, 2, 10, 0, 0));
        var sameDay = AddFile(temp.Path, "same-day.jpg", new DateTime(2026, 1, 3, 18, 0, 0));
        var paths = new[] { first, second, sameDay };
        var enumerationCount = 0;
        IEnumerable<string> Files()
        {
            enumerationCount++;
            foreach (var path in paths) yield return path;
        }

        var result = new MediaDateResolver().ResolveDateSummary(Files());

        Assert.AreEqual(1, enumerationCount);
        Assert.AreEqual("2026-01-02", result.EarliestDateKey);
        CollectionAssert.AreEqual(
            new[] { "2026-01-02", "2026-01-03" },
            result.DateKeys.ToArray());
    }

    [TestMethod]
    public void ResolveDateSummary_EmptyInputHasNoDistinctKeys()
    {
        var dateBefore = DateTime.Now.ToString("yyyy-MM-dd");
        var result = new MediaDateResolver().ResolveDateSummary(Array.Empty<string>());
        var dateAfter = DateTime.Now.ToString("yyyy-MM-dd");

        Assert.IsTrue(
            result.EarliestDateKey == dateBefore || result.EarliestDateKey == dateAfter,
            $"Unexpected fallback date: {result.EarliestDateKey}");
        Assert.AreEqual(0, result.DateKeys.Count);
    }

    private static string AddFile(string root, string name, DateTime timestamp)
    {
        var path = Path.Combine(root, name);
        File.WriteAllText(path, "not-an-image");
        File.SetLastWriteTime(path, timestamp);
        return path;
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"PhotoOrganizerDateResolverTests-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch { }
        }
    }
}
