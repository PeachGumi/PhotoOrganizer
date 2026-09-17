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
    private const string IOKit = "/System/Library/Frameworks/IOKit.framework/IOKit";
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
            var physical = ReadPhysicalDisk(disk);
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

    private static string? ReadPhysicalDisk(IntPtr disk)
    {
        var current = DADiskCopyIOMedia(disk);
        var owned = new List<IntPtr>();
        IntPtr Key(string value)
        {
            var key = CFStringCreateWithCString(IntPtr.Zero, value, Utf8);
            if (key == IntPtr.Zero) throw new InvalidOperationException("Unable to create an IOKit property key.");
            owned.Add(key);
            return key;
        }
        IntPtr Property(IntPtr key)
        {
            var value = IORegistryEntryCreateCFProperty(current, key, IntPtr.Zero, 0);
            if (value != IntPtr.Zero) owned.Add(value);
            return value;
        }
        static bool IsTrue(IntPtr value) => value != IntPtr.Zero
            && CFGetTypeID(value) == CFBooleanGetTypeID() && CFBooleanGetValue(value);
        static string? ReadString(IntPtr value)
        {
            if (value == IntPtr.Zero || CFGetTypeID(value) != CFStringGetTypeID()) return null;
            var buffer = new byte[256];
            if (!CFStringGetCString(value, buffer, buffer.Length, Utf8)) return null;
            var end = Array.IndexOf(buffer, (byte)0);
            return end > 0 ? Encoding.UTF8.GetString(buffer, 0, end) : null;
        }

        try
        {
            var wholeKey = Key("Whole");
            var nameKey = Key("BSD Name");
            var compositeKey = Key("APFSComposited");
            var protocolKey = Key("Protocol Characteristics");
            var interconnectKey = Key("Physical Interconnect");
            string? wholeDisk = null;

            // APFS containers are themselves marked Whole. Follow the unique I/O
            // ancestry through them to the actual block device; different APFS
            // containers on the same disk must have the same physical fingerprint.
            for (var depth = 0; current != 0 && depth < 128; depth++)
            {
                if (IsTrue(Property(compositeKey))) return null;
                if (IOObjectConformsTo(current, "IOMedia") && IsTrue(Property(wholeKey)))
                {
                    wholeDisk = ReadString(Property(nameKey));
                    if (wholeDisk is null) return null;
                }

                if (IOObjectConformsTo(current, "IOBlockStorageDevice"))
                {
                    var protocol = Property(protocolKey);
                    if (protocol == IntPtr.Zero || CFGetTypeID(protocol) != CFDictionaryGetTypeID()) return null;
                    var interconnect = ReadString(CFDictionaryGetValue(protocol, interconnectKey));
                    // File-backed images and RAM disks cannot prove an independent
                    // physical backup, even though they have their own BSD disk ID.
                    return string.IsNullOrWhiteSpace(interconnect) || interconnect == "Virtual Interface"
                        ? null
                        : wholeDisk;
                }

                if (IORegistryEntryGetParentIterator(current, "IOService", out var iterator) != 0) return null;
                uint parent = 0;
                try
                {
                    parent = IOIteratorNext(iterator);
                    var other = IOIteratorNext(iterator);
                    var ambiguous = other != 0;
                    if (other != 0) IOObjectRelease(other);
                    if (parent == 0 || ambiguous || !IOIteratorIsValid(iterator)) return null;
                    IOObjectRelease(current);
                    current = parent;
                    parent = 0;
                }
                finally
                {
                    if (parent != 0) IOObjectRelease(parent);
                    IOObjectRelease(iterator);
                }
            }
            return null;
        }
        finally
        {
            if (current != 0) IOObjectRelease(current);
            for (var index = owned.Count - 1; index >= 0; index--) CFRelease(owned[index]);
        }
    }

    [DllImport(DiskArbitration)] private static extern IntPtr DASessionCreate(IntPtr allocator);
    [DllImport(DiskArbitration)] private static extern IntPtr DADiskCreateFromVolumePath(IntPtr allocator, IntPtr session, IntPtr path);
    [DllImport(DiskArbitration)] private static extern IntPtr DADiskCopyDescription(IntPtr disk);
    [DllImport(DiskArbitration)] private static extern uint DADiskCopyIOMedia(IntPtr disk);
    [DllImport(DiskArbitration)] private static extern IntPtr DADiskGetBSDName(IntPtr disk);
    [DllImport(CoreFoundation)]
    private static extern IntPtr CFURLCreateFromFileSystemRepresentation(
        IntPtr allocator, byte[] bytes, nint length, [MarshalAs(UnmanagedType.I1)] bool isDirectory);
    [DllImport(CoreFoundation)] private static extern IntPtr CFDictionaryGetValue(IntPtr dictionary, IntPtr key);
    [DllImport(CoreFoundation)] private static extern nuint CFGetTypeID(IntPtr value);
    [DllImport(CoreFoundation)] private static extern nuint CFUUIDGetTypeID();
    [DllImport(CoreFoundation)] private static extern nuint CFStringGetTypeID();
    [DllImport(CoreFoundation)] private static extern nuint CFBooleanGetTypeID();
    [DllImport(CoreFoundation)] private static extern nuint CFDictionaryGetTypeID();
    [DllImport(CoreFoundation)]
    [return: MarshalAs(UnmanagedType.I1)]
    private static extern bool CFBooleanGetValue(IntPtr value);
    [DllImport(CoreFoundation)] private static extern IntPtr CFStringCreateWithCString(IntPtr allocator, string text, uint encoding);
    [DllImport(CoreFoundation)] private static extern IntPtr CFUUIDCreateString(IntPtr allocator, IntPtr uuid);
    [DllImport(CoreFoundation)]
    [return: MarshalAs(UnmanagedType.I1)]
    private static extern bool CFStringGetCString(IntPtr text, byte[] buffer, nint bufferSize, uint encoding);
    [DllImport(CoreFoundation)] private static extern void CFRelease(IntPtr value);
    [DllImport(IOKit)] private static extern bool IOObjectConformsTo(uint value, string className);
    [DllImport(IOKit)] private static extern int IOObjectRelease(uint value);
    [DllImport(IOKit)] private static extern int IORegistryEntryGetParentIterator(uint entry, string plane, out uint iterator);
    [DllImport(IOKit)] private static extern uint IOIteratorNext(uint iterator);
    [DllImport(IOKit)] private static extern bool IOIteratorIsValid(uint iterator);
    [DllImport(IOKit)] private static extern IntPtr IORegistryEntryCreateCFProperty(uint entry, IntPtr key, IntPtr allocator, uint options);
}
