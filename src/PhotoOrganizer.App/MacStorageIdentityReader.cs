using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;

namespace PhotoOrganizer.App;

/// <summary>Reads fresh OS identities without launching diskutil or caching mount state.</summary>
[SupportedOSPlatform("macos")]
internal static class MacStorageIdentityReader
{
    private const string DiskArbitration = "/System/Library/Frameworks/DiskArbitration.framework/DiskArbitration";
    private const string CoreFoundation = "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation";
    private const uint Utf8 = 0x08000100;

    public static MacStorageIdentity? Read(string root)
    {
        var owned = new List<IntPtr>();
        var library = IntPtr.Zero;
        IntPtr Own(IntPtr value) { if (value != IntPtr.Zero) owned.Add(value); return value; }
        try
        {
            library = NativeLibrary.Load(DiskArbitration);
            var session = Own(DASessionCreate(IntPtr.Zero));
            if (session == IntPtr.Zero) return null;
            var bytes = Encoding.UTF8.GetBytes(root);
            var url = Own(CFURLCreateFromFileSystemRepresentation(IntPtr.Zero, bytes, bytes.Length, true));
            if (url == IntPtr.Zero) return null;
            var disk = Own(DADiskCreateFromVolumePath(IntPtr.Zero, session, url));
            if (disk == IntPtr.Zero) return null;
            var description = Own(DADiskCopyDescription(disk));
            if (description == IntPtr.Zero) return null;

            string? ReadUuid(string symbol)
            {
                var key = Marshal.ReadIntPtr(NativeLibrary.GetExport(library, symbol));
                var value = CFDictionaryGetValue(description, key);
                if (value == IntPtr.Zero || CFGetTypeID(value) != CFUUIDGetTypeID()) return null;
                var text = Own(CFUUIDCreateString(IntPtr.Zero, value));
                if (text == IntPtr.Zero) return null;
                var buffer = new byte[128];
                if (!CFStringGetCString(text, buffer, buffer.Length, Utf8)) return null;
                var end = Array.IndexOf(buffer, (byte)0);
                return end > 0 ? Encoding.UTF8.GetString(buffer, 0, end) : null;
            }

            var volume = ReadUuid("kDADiskDescriptionVolumeUUIDKey")
                ?? ReadUuid("kDADiskDescriptionMediaUUIDKey")
                ?? Marshal.PtrToStringUTF8(DADiskGetBSDName(disk));
            if (string.IsNullOrWhiteSpace(volume)) return null;
            var wholeDisk = Own(DADiskCopyWholeDisk(disk));
            var physical = wholeDisk == IntPtr.Zero ? null : Marshal.PtrToStringUTF8(DADiskGetBSDName(wholeDisk));
            return new MacStorageIdentity("mac-volume:" + volume,
                string.IsNullOrWhiteSpace(physical) ? null : "mac-whole-disk:" + physical);
        }
        catch
        {
            // An unavailable framework or identity is not proof that a volume is safe.
            return null;
        }
        finally
        {
            for (var index = owned.Count - 1; index >= 0; index--) CFRelease(owned[index]);
            if (library != IntPtr.Zero) NativeLibrary.Free(library);
        }
    }

    [DllImport(DiskArbitration)] private static extern IntPtr DASessionCreate(IntPtr allocator);
    [DllImport(DiskArbitration)] private static extern IntPtr DADiskCreateFromVolumePath(IntPtr allocator, IntPtr session, IntPtr path);
    [DllImport(DiskArbitration)] private static extern IntPtr DADiskCopyDescription(IntPtr disk);
    [DllImport(DiskArbitration)] private static extern IntPtr DADiskCopyWholeDisk(IntPtr disk);
    [DllImport(DiskArbitration)] private static extern IntPtr DADiskGetBSDName(IntPtr disk);
    [DllImport(CoreFoundation)] private static extern IntPtr CFURLCreateFromFileSystemRepresentation(
        IntPtr allocator, byte[] bytes, nint length, [MarshalAs(UnmanagedType.I1)] bool isDirectory);
    [DllImport(CoreFoundation)] private static extern IntPtr CFDictionaryGetValue(IntPtr dictionary, IntPtr key);
    [DllImport(CoreFoundation)] private static extern nuint CFGetTypeID(IntPtr value);
    [DllImport(CoreFoundation)] private static extern nuint CFUUIDGetTypeID();
    [DllImport(CoreFoundation)] private static extern IntPtr CFUUIDCreateString(IntPtr allocator, IntPtr uuid);
    [DllImport(CoreFoundation)]
    [return: MarshalAs(UnmanagedType.I1)]
    private static extern bool CFStringGetCString(IntPtr text, byte[] buffer, nint bufferSize, uint encoding);
    [DllImport(CoreFoundation)] private static extern void CFRelease(IntPtr value);
}
