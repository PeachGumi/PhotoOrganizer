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

    [TestMethod]
    public void EnsureDurable_LocalVolume_ReportsFullSyncLevel()
    {
        if (!OperatingSystem.IsMacOS()) return;

        var root = Path.Combine(Path.GetTempPath(), $"PhotoOrganizerDurabilityLevel-{Guid.NewGuid():N}");
        var path = Path.Combine(root, "photo.jpg");
        Directory.CreateDirectory(root);
        File.WriteAllText(path, "camera-bytes");

        try
        {
            var result = new PlatformFileDurabilityService().EnsureDurable(path);

            Assert.IsTrue(result.Success, result.Error);
            Assert.AreEqual(DurabilityLevel.FullSync, result.Level);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [TestMethod]
    public void IsOperationUnsupported_OnlyTreatsCapabilityErrorsAsUnsupported()
    {
        // Network mounts (SMB/NFS) answer ENOTSUP or EOPNOTSUPP for F_FULLFSYNC and
        // renamex_np(RENAME_EXCL), and implement neither.
        Assert.IsTrue(PlatformFileDurabilityService.IsOperationUnsupported(45), "ENOTSUP");
        Assert.IsTrue(PlatformFileDurabilityService.IsOperationUnsupported(102), "EOPNOTSUPP");

        // Real failures must keep failing the copy instead of being downgraded.
        Assert.IsFalse(PlatformFileDurabilityService.IsOperationUnsupported(4), "EINTR is retried");
        Assert.IsFalse(PlatformFileDurabilityService.IsOperationUnsupported(1), "EPERM");
        Assert.IsFalse(PlatformFileDurabilityService.IsOperationUnsupported(5), "EIO");
        Assert.IsFalse(PlatformFileDurabilityService.IsOperationUnsupported(13), "EACCES");
        Assert.IsFalse(PlatformFileDurabilityService.IsOperationUnsupported(17), "EEXIST is a collision");
    }
}
