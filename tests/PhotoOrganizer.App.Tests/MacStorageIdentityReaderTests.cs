using System.Xml.Linq;
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
        var properties = XDocument.Parse(output.StandardOutput).Root!.Element("dict")!;
        var stores = properties.Elements("key")
            .FirstOrDefault(key => key.Value == "APFSPhysicalStores")?
            .ElementsAfterSelf().FirstOrDefault()?.Elements("dict").ToArray();
        if (stores is { Length: > 0 })
        {
            if (stores.Length != 1)
            {
                Assert.IsNull(MacStorageIdentityReader.Read("/"),
                    "A volume backed by multiple APFS physical stores must fail closed.");
                return;
            }
            var store = stores[0].Elements("key").Single(key => key.Value == "APFSPhysicalStore")
                .ElementsAfterSelf().First().Value;
            var physicalOutput = BoundedProcessRunner.Run("/usr/sbin/diskutil", ["info", "-plist", store], TimeSpan.FromSeconds(5));
            Assert.IsNotNull(physicalOutput);
            Assert.AreEqual(0, physicalOutput.ExitCode);
            var physical = MacDiskutilInfoParser.Parse(physicalOutput.StandardOutput);
            Assert.IsNotNull(physical);
            expected = expected with { PhysicalDeviceFingerprint = physical.PhysicalDeviceFingerprint };
        }
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

    [TestMethod]
    public void FileBackedDiskImage_CannotProvePhysicalIndependence()
    {
        if (!OperatingSystem.IsMacOS())
        {
            Assert.Inconclusive("Disk images require macOS.");
            return;
        }

        var directory = Path.Combine(Path.GetTempPath(), "PhotoOrganizer-Identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var image = Path.Combine(directory, "test.dmg");
        var mount = Path.Combine(directory, "mount");
        Directory.CreateDirectory(mount);
        var attached = false;
        try
        {
            var create = BoundedProcessRunner.Run("/usr/bin/hdiutil",
                ["create", "-size", "16m", "-fs", "HFS+", "-volname", "PhotoOrganizer-Identity-Test", image], TimeSpan.FromSeconds(30));
            Assert.IsNotNull(create);
            Assert.AreEqual(0, create.ExitCode, create.StandardError);
            var attach = BoundedProcessRunner.Run("/usr/bin/hdiutil",
                ["attach", "-nobrowse", "-mountpoint", mount, image], TimeSpan.FromSeconds(30));
            Assert.IsNotNull(attach);
            Assert.AreEqual(0, attach.ExitCode, attach.StandardError);
            attached = true;

            var identity = MacStorageIdentityReader.Read(mount);
            Assert.IsNotNull(identity, "The test image should have a volume identity.");
            Assert.IsNull(identity.PhysicalDeviceFingerprint,
                "A file-backed image must not count as an independent physical storage device.");
        }
        finally
        {
            if (attached)
            {
                for (var attempt = 0; attempt < 3 && attached; attempt++)
                {
                    var detach = BoundedProcessRunner.Run("/usr/bin/hdiutil", ["detach", mount], TimeSpan.FromSeconds(30));
                    attached = detach?.ExitCode != 0;
                }
            }
            // Never recursively clean up a directory that may still be mounted.
            if (!attached) Directory.Delete(directory, recursive: true);
        }
    }
}
