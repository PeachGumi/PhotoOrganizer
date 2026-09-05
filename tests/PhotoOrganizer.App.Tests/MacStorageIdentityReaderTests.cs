using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class MacStorageIdentityReaderTests
{
    [TestMethod]
    public void NativeIdentity_MatchesFreshDiskutilIdentity()
    {
        if (!OperatingSystem.IsMacOS())
        {
            Assert.Inconclusive("Disk Arbitration requires macOS.");
            return;
        }
        var output = BoundedProcessRunner.Run("/usr/sbin/diskutil", ["info", "-plist", "/"], TimeSpan.FromSeconds(5));
        Assert.IsNotNull(output);
        Assert.AreEqual(0, output.ExitCode);
        var expected = MacDiskutilInfoParser.Parse(output.StandardOutput);
        Assert.IsNotNull(expected);
        for (var index = 0; index < 3; index++)
        {
            Assert.AreEqual(expected, MacStorageIdentityReader.Read("/"));
        }
    }

    [TestMethod]
    public void MissingMount_DoesNotManufactureIdentity()
    {
        if (!OperatingSystem.IsMacOS())
        {
            Assert.Inconclusive("Disk Arbitration requires macOS.");
            return;
        }
        Assert.IsNull(MacStorageIdentityReader.Read($"/Volumes/PhotoOrganizer-missing-{Guid.NewGuid():N}"));
    }
}
