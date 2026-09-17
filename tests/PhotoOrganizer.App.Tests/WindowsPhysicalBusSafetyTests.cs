using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class WindowsPhysicalBusSafetyTests
{
    [TestMethod]
    [DataRow(0)] // Unknown
    [DataRow(8)] // RAID
    [DataRow(9)] // iSCSI
    [DataRow(14)] // Virtual
    [DataRow(15)] // File-backed VHD
    [DataRow(16)] // Storage Spaces
    [DataRow(65535)] // Unrecognized
    public void LogicalOrUnknownDisk_CannotProvePhysicalIndependence(int busType)
    {
        Assert.IsFalse(PlatformStorageVolumeProvider.IsIndependentPhysicalBusType((ushort)busType));
    }

    [TestMethod]
    [DataRow(1)] // SCSI
    [DataRow(7)] // USB
    [DataRow(11)] // SATA
    [DataRow(12)] // SD
    [DataRow(13)] // MMC
    [DataRow(17)] // NVMe
    public void PhysicalStorageBus_IsRecognized(int busType)
    {
        Assert.IsTrue(PlatformStorageVolumeProvider.IsIndependentPhysicalBusType((ushort)busType));
    }
}
