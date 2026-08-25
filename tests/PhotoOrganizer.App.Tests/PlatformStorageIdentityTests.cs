using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.Core;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class PlatformStorageIdentityTests
{
    [TestMethod]
    public void SystemVolume_ResolvesVolumeAndPhysicalDeviceIdentity()
    {
        var provider = new PlatformStorageVolumeProvider();
        var path = GetSystemVolumePath();

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
        var path = GetSystemVolumePath();

        var first = provider.ResolveVolumeForPath(path);
        var second = provider.ResolveVolumeForPath(path);

        Assert.IsNotNull(first);
        Assert.IsNotNull(second);
        Assert.AreEqual(first.Fingerprint, second.Fingerprint);
        Assert.AreEqual(first.PhysicalDeviceFingerprint, second.PhysicalDeviceFingerprint);
    }

    [TestMethod]
    public void SystemVolume_AppearsInMountedVolumeSnapshotAndIsMarkedSystem()
    {
        var provider = new PlatformStorageVolumeProvider();
        var resolved = provider.ResolveVolumeForPath(GetSystemVolumePath());

        Assert.IsNotNull(resolved);
        var snapshot = provider.GetMountedVolumes();
        var sameRoot = snapshot.SingleOrDefault(volume =>
            string.Equals(
                PathSafety.Normalize(volume.RootPath),
                PathSafety.Normalize(resolved.RootPath),
                provider.PathComparison));

        Assert.IsNotNull(sameRoot);
        Assert.IsTrue(sameRoot.IsSystem);
        Assert.AreEqual(resolved.Fingerprint, sameRoot.Fingerprint);
        Assert.AreEqual(resolved.PhysicalDeviceFingerprint, sameRoot.PhysicalDeviceFingerprint);
    }

    [TestMethod]
    public void SystemVolume_UsesPlatformSpecificIdentityPrefixes()
    {
        var provider = new PlatformStorageVolumeProvider();
        var volume = provider.ResolveVolumeForPath(GetSystemVolumePath());

        Assert.IsNotNull(volume);
        if (OperatingSystem.IsWindows())
        {
            StringAssert.StartsWith(volume.Fingerprint, "windows-volume:");
            StringAssert.StartsWith(volume.PhysicalDeviceFingerprint!, "windows-physical-disk:");
        }
        else if (OperatingSystem.IsMacOS())
        {
            StringAssert.StartsWith(volume.Fingerprint, "mac-volume:");
            StringAssert.StartsWith(volume.PhysicalDeviceFingerprint!, "mac-whole-disk:");
        }
    }

    [TestMethod]
    public void ExistingFile_ResolvesToSameIdentityAsItsDirectory()
    {
        var provider = new PlatformStorageVolumeProvider();
        var directory = CreateTemporaryDirectory();
        var file = Path.Combine(directory, "probe.txt");
        File.WriteAllText(file, "identity-probe");

        try
        {
            var parentVolume = provider.ResolveVolumeForPath(directory);
            var fileVolume = provider.ResolveVolumeForPath(file);

            Assert.IsNotNull(parentVolume);
            Assert.IsNotNull(fileVolume);
            Assert.AreEqual(parentVolume.Fingerprint, fileVolume.Fingerprint);
            Assert.AreEqual(parentVolume.PhysicalDeviceFingerprint, fileVolume.PhysicalDeviceFingerprint);
            Assert.IsTrue(string.Equals(parentVolume.RootPath, fileVolume.RootPath, provider.PathComparison));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [TestMethod]
    public void MissingDestinationLeaf_UsesNearestExistingVolume()
    {
        var provider = new PlatformStorageVolumeProvider();
        var directory = CreateTemporaryDirectory();

        try
        {
            var missing = Path.Combine(directory, "future", "event", "JPG");
            var parentVolume = provider.ResolveVolumeForPath(directory);
            var missingVolume = provider.ResolveVolumeForPath(missing);

            Assert.IsNotNull(parentVolume);
            Assert.IsNotNull(missingVolume);
            Assert.AreEqual(parentVolume.Fingerprint, missingVolume.Fingerprint);
            Assert.AreEqual(parentVolume.PhysicalDeviceFingerprint, missingVolume.PhysicalDeviceFingerprint);
            Assert.IsTrue(string.Equals(parentVolume.RootPath, missingVolume.RootPath, provider.PathComparison));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [TestMethod]
    public void BlankPath_CannotResolveVolume()
    {
        var provider = new PlatformStorageVolumeProvider();

        Assert.IsNull(provider.ResolveVolumeForPath(string.Empty));
        Assert.IsNull(provider.ResolveVolumeForPath("   "));
    }

    [TestMethod]
    public void MountedVolumeSnapshot_NeverContainsBlankFingerprintOrRoot()
    {
        var provider = new PlatformStorageVolumeProvider();
        var snapshot = provider.GetMountedVolumes();

        Assert.IsTrue(snapshot.Count > 0);
        foreach (var volume in snapshot)
        {
            Assert.IsFalse(string.IsNullOrWhiteSpace(volume.RootPath));
            Assert.IsFalse(string.IsNullOrWhiteSpace(volume.Fingerprint));
        }
    }

    private static string GetSystemVolumePath() => OperatingSystem.IsWindows()
        ? Path.GetPathRoot(Environment.SystemDirectory)!
        : "/";

    private static string CreateTemporaryDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), "PhotoOrganizer-AppTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }
}
