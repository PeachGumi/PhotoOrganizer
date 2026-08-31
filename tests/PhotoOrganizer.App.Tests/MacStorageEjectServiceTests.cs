using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class MacStorageEjectServiceTests
{
    [TestMethod]
    public void WholeDiskFingerprint_ProducesDiskutilDeviceIdentifier()
    {
        Assert.AreEqual("disk9", MacStorageEjectService.GetWholeDiskIdentifier("mac-whole-disk:disk9"));
    }

    [TestMethod]
    [DataRow(null)]
    [DataRow("")]
    [DataRow("windows-physical-disk:2")]
    [DataRow("mac-whole-disk:disk9s1")]
    [DataRow("mac-whole-disk:disk9; reboot")]
    public void UnsafeOrNonWholeDiskFingerprint_IsRejected(string? fingerprint)
    {
        Assert.IsNull(MacStorageEjectService.GetWholeDiskIdentifier(fingerprint));
    }
}
