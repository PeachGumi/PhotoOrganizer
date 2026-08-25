using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class PlatformStorageIdentityTests
{
    [TestMethod]
    public void SystemVolume_ResolvesVolumeAndPhysicalDeviceIdentity()
    {
        var provider = new PlatformStorageVolumeProvider();
        var path = OperatingSystem.IsWindows()
            ? Path.GetPathRoot(Environment.SystemDirectory)!
            : "/";

        var volume = provider.ResolveVolumeForPath(path);

        Assert.IsNotNull(volume, $"Unable to resolve mounted volume for {path}.");
        Assert.IsFalse(string.IsNullOrWhiteSpace(volume.Fingerprint), "Mounted-volume identity is missing.");
        Assert.IsFalse(
            string.IsNullOrWhiteSpace(volume.PhysicalDeviceFingerprint),
            $"Physical-device identity is missing for {path} on {Environment.OSVersion}.");
    }

    [TestMethod]
    public void SystemVolume_RepeatedResolutionKeepsSameIdentity()
    {
        var provider = new PlatformStorageVolumeProvider();
        var path = OperatingSystem.IsWindows()
            ? Path.GetPathRoot(Environment.SystemDirectory)!
            : "/";

        var first = provider.ResolveVolumeForPath(path);
        var second = provider.ResolveVolumeForPath(path);

        Assert.IsNotNull(first);
        Assert.IsNotNull(second);
        Assert.AreEqual(first.Fingerprint, second.Fingerprint);
        Assert.AreEqual(first.PhysicalDeviceFingerprint, second.PhysicalDeviceFingerprint);
    }
}
