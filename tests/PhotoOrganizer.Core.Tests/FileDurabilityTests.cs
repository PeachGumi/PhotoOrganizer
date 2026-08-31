using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class FileDurabilityTests
{
    [TestMethod]
    public void EnsureDurable_MacReadOnlyFile_SucceedsWithoutChangingPermissions()
    {
        if (!OperatingSystem.IsMacOS()) return;

        var root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerDurability-{Guid.NewGuid():N}");
        var path = Path.Combine(root, "locked.jpg");
        Directory.CreateDirectory(root);
        File.WriteAllText(path, "camera-bytes");
        File.SetUnixFileMode(path, UnixFileMode.UserRead);

        try
        {
            var before = File.GetUnixFileMode(path);
            var result = new PlatformFileDurabilityService().EnsureDurable(path);

            Assert.IsTrue(result.Success, result.Error);
            Assert.AreEqual(before, File.GetUnixFileMode(path));
            Assert.AreEqual("camera-bytes", File.ReadAllText(path));
        }
        finally
        {
            File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite);
            Directory.Delete(root, recursive: true);
        }
    }
}
