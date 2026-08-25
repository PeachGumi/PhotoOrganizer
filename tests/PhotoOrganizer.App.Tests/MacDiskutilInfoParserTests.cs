using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class MacDiskutilInfoParserTests
{
    [TestMethod]
    public void VolumeUuid_IsPreferredOverAllFallbacks()
    {
        var result = Parse(
            String("VolumeUUID", "volume-uuid"),
            String("APFSVolumeUUID", "apfs-uuid"),
            String("DiskUUID", "disk-uuid"),
            String("MediaUUID", "media-uuid"),
            String("DeviceIdentifier", "disk9s1"),
            String("ParentWholeDisk", "disk9"));

        AssertIdentity(result, "mac-volume:volume-uuid", "mac-whole-disk:disk9");
    }

    [TestMethod]
    public void ApfsVolumeUuid_IsUsedWhenVolumeUuidMissing()
    {
        var result = Parse(
            String("APFSVolumeUUID", "apfs-uuid"),
            String("DeviceIdentifier", "disk3s2"),
            String("ParentWholeDisk", "disk3"));

        AssertIdentity(result, "mac-volume:apfs-uuid", "mac-whole-disk:disk3");
    }

    [TestMethod]
    public void DiskUuid_IsUsedWhenHigherPriorityIdsMissing()
    {
        var result = Parse(
            String("DiskUUID", "disk-uuid"),
            String("DeviceIdentifier", "disk4s1"),
            String("ParentWholeDisk", "disk4"));

        AssertIdentity(result, "mac-volume:disk-uuid", "mac-whole-disk:disk4");
    }

    [TestMethod]
    public void MediaUuid_IsUsedWhenHigherPriorityIdsMissing()
    {
        var result = Parse(
            String("MediaUUID", "media-uuid"),
            String("DeviceIdentifier", "disk5s1"),
            String("ParentWholeDisk", "disk5"));

        AssertIdentity(result, "mac-volume:media-uuid", "mac-whole-disk:disk5");
    }

    [TestMethod]
    public void DeviceIdentifier_IsLastResortVolumeIdentity()
    {
        var result = Parse(
            String("DeviceIdentifier", "disk6s1"),
            String("ParentWholeDisk", "disk6"));

        AssertIdentity(result, "mac-volume:disk6s1", "mac-whole-disk:disk6");
    }

    [TestMethod]
    public void ParentWholeDisk_IsPreferredPhysicalIdentity()
    {
        var result = Parse(
            String("VolumeUUID", "volume"),
            String("DeviceIdentifier", "disk7"),
            String("ParentWholeDisk", "disk-parent"),
            Boolean("WholeDisk", true));

        AssertIdentity(result, "mac-volume:volume", "mac-whole-disk:disk-parent");
    }

    [TestMethod]
    public void WholeDiskTrue_UsesDeviceIdentifierAsPhysicalIdentity()
    {
        var result = Parse(
            String("VolumeUUID", "whole-volume"),
            String("DeviceIdentifier", "disk8"),
            Boolean("WholeDisk", true));

        AssertIdentity(result, "mac-volume:whole-volume", "mac-whole-disk:disk8");
    }

    [TestMethod]
    public void WholeDiskFalse_DoesNotInventPhysicalIdentity()
    {
        var result = Parse(
            String("VolumeUUID", "partition-volume"),
            String("DeviceIdentifier", "disk8s1"),
            Boolean("WholeDisk", false));

        AssertIdentity(result, "mac-volume:partition-volume", null);
    }

    [TestMethod]
    public void MissingPhysicalIdentity_PreservesVolumeButFailsPhysicalProofUpstream()
    {
        var result = Parse(String("VolumeUUID", "volume-only"));

        AssertIdentity(result, "mac-volume:volume-only", null);
    }

    [TestMethod]
    public void WhitespaceValues_AreTrimmedAndBlankFallbacksAreSkipped()
    {
        var result = Parse(
            String("VolumeUUID", "   "),
            String("APFSVolumeUUID", "  apfs-trimmed  "),
            String("ParentWholeDisk", "  disk10  "));

        AssertIdentity(result, "mac-volume:apfs-trimmed", "mac-whole-disk:disk10");
    }

    [TestMethod]
    public void WrongValueType_IsIgnoredAndFallsBackSafely()
    {
        var result = Parse(
            Integer("VolumeUUID", 123),
            String("DiskUUID", "valid-disk-uuid"),
            String("ParentWholeDisk", "disk11"));

        AssertIdentity(result, "mac-volume:valid-disk-uuid", "mac-whole-disk:disk11");
    }

    [TestMethod]
    public void WrongWholeDiskType_DoesNotTreatPartitionAsWholeDisk()
    {
        var result = Parse(
            String("VolumeUUID", "volume"),
            String("DeviceIdentifier", "disk12s1"),
            String("WholeDisk", "true"));

        AssertIdentity(result, "mac-volume:volume", null);
    }

    [TestMethod]
    public void MissingVolumeIdentity_ReturnsNull()
    {
        var result = Parse(
            String("ParentWholeDisk", "disk13"),
            Boolean("WholeDisk", false));

        Assert.IsNull(result);
    }

    [TestMethod]
    public void EmptyOutput_ReturnsNull()
    {
        Assert.IsNull(MacDiskutilInfoParser.Parse(string.Empty));
        Assert.IsNull(MacDiskutilInfoParser.Parse("   \n\t"));
        Assert.IsNull(MacDiskutilInfoParser.Parse(null));
    }

    [TestMethod]
    public void MalformedXml_ReturnsNull()
    {
        Assert.IsNull(MacDiskutilInfoParser.Parse("<plist><dict><key>VolumeUUID</key>"));
    }

    [TestMethod]
    public void MissingDictionary_ReturnsNull()
    {
        Assert.IsNull(MacDiskutilInfoParser.Parse("<plist version=\"1.0\"><array/></plist>"));
    }

    [TestMethod]
    public void DirectDictionaryRoot_IsAccepted()
    {
        var result = MacDiskutilInfoParser.Parse(
            "<dict><key>VolumeUUID</key><string>direct</string>" +
            "<key>ParentWholeDisk</key><string>disk14</string></dict>");

        AssertIdentity(result, "mac-volume:direct", "mac-whole-disk:disk14");
    }

    [TestMethod]
    public void UnknownKeys_DoNotChangeIdentitySelection()
    {
        var result = Parse(
            String("Unrelated", "value"),
            Boolean("Writable", true),
            String("VolumeUUID", "known"),
            String("ParentWholeDisk", "disk15"));

        AssertIdentity(result, "mac-volume:known", "mac-whole-disk:disk15");
    }

    private static MacStorageIdentity? Parse(params string[] entries) =>
        MacDiskutilInfoParser.Parse(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<plist version=\"1.0\"><dict>" +
            string.Concat(entries) +
            "</dict></plist>");

    private static string String(string key, string value) =>
        $"<key>{key}</key><string>{value}</string>";

    private static string Integer(string key, int value) =>
        $"<key>{key}</key><integer>{value}</integer>";

    private static string Boolean(string key, bool value) =>
        $"<key>{key}</key><{(value ? "true" : "false")}/>";

    private static void AssertIdentity(
        MacStorageIdentity? actual,
        string expectedVolume,
        string? expectedPhysical)
    {
        Assert.IsNotNull(actual);
        Assert.AreEqual(expectedVolume, actual.VolumeFingerprint);
        Assert.AreEqual(expectedPhysical, actual.PhysicalDeviceFingerprint);
    }
}
