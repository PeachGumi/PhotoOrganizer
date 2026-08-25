using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PhotoOrganizer.Core;

public sealed record DurabilityResult(bool Success, string? Error = null);

public enum FinalizeFileStatus
{
    Committed,
    DestinationExists,
    Failed
}

public sealed record FinalizeFileResult(
    FinalizeFileStatus Status,
    bool FinalPathCreated,
    string? Error = null);

public interface IFileDurabilityService
{
    FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc);

    DurabilityResult EnsureDurable(string filePath);
}

public sealed partial class PlatformFileDurabilityService : IFileDurabilityService
{
    private const uint MoveFileWriteThrough = 0x00000008;
    private const int FFullFsync = 51;
    private const int OpenReadOnly = 0;
    private const int OpenReadWrite = 2;
    private const int InterruptedSystemCall = 4;

    public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc)
    {
        var moved = false;

        try
        {
            if (OperatingSystem.IsWindows())
            {
                if (!MoveFileEx(temporaryPath, finalPath, MoveFileWriteThrough))
                {
                    var error = Marshal.GetLastPInvokeError();
                    if (File.Exists(finalPath) || Directory.Exists(finalPath))
                    {
                        return new FinalizeFileResult(FinalizeFileStatus.DestinationExists, false);
                    }

                    return new FinalizeFileResult(
                        FinalizeFileStatus.Failed,
                        false,
                        $"MOVEFILE_WRITE_THROUGH final move failed ({error}): {new Win32Exception(error).Message}");
                }
            }
            else
            {
                try
                {
                    File.Move(temporaryPath, finalPath, overwrite: false);
                }
                catch (IOException) when (File.Exists(finalPath) || Directory.Exists(finalPath))
                {
                    return new FinalizeFileResult(FinalizeFileStatus.DestinationExists, false);
                }
            }

            moved = true;
            File.SetLastWriteTimeUtc(finalPath, lastWriteUtc);

            var durability = EnsureDurable(finalPath);
            if (!durability.Success)
            {
                return new FinalizeFileResult(
                    FinalizeFileStatus.Failed,
                    true,
                    durability.Error ?? "Finalized file could not be committed durably.");
            }

            return new FinalizeFileResult(FinalizeFileStatus.Committed, true);
        }
        catch (Exception ex)
        {
            return new FinalizeFileResult(
                FinalizeFileStatus.Failed,
                moved,
                $"Durable finalization failed: {ex.Message}");
        }
    }

    public DurabilityResult EnsureDurable(string filePath)
    {
        try
        {
            var fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath))
            {
                return new DurabilityResult(false, "Destination file no longer exists while establishing durability.");
            }

            if (OperatingSystem.IsWindows())
            {
                using var stream = new FileStream(
                    fullPath,
                    FileMode.Open,
                    FileAccess.ReadWrite,
                    FileShare.Read,
                    bufferSize: 1,
                    FileOptions.WriteThrough);
                stream.Flush(flushToDisk: true);
                return new DurabilityResult(true);
            }

            if (OperatingSystem.IsMacOS())
            {
                var parent = Path.GetDirectoryName(fullPath);
                if (string.IsNullOrWhiteSpace(parent))
                {
                    return new DurabilityResult(false, "Destination parent directory could not be resolved for durable commit.");
                }

                // Push the rename/directory-entry metadata to the device first. F_FULLFSYNC
                // on the finalized file then drains the device queue, covering that prior fsync
                // as well as the finalized file's data and timestamp metadata.
                var directorySync = SynchronizeMacDirectory(parent);
                if (!directorySync.Success)
                {
                    return directorySync;
                }

                return FullFsyncMacFile(fullPath);
            }

            // The shipping targets are Windows and macOS. Keep non-shipping Unix CI
            // functional with the strongest portable .NET flush available.
            using (var stream = new FileStream(
                       fullPath,
                       FileMode.Open,
                       FileAccess.ReadWrite,
                       FileShare.Read,
                       bufferSize: 1,
                       FileOptions.WriteThrough))
            {
                stream.Flush(flushToDisk: true);
            }

            return new DurabilityResult(true);
        }
        catch (Exception ex)
        {
            return new DurabilityResult(false, $"Destination durability flush failed: {ex.Message}");
        }
    }

    private static DurabilityResult SynchronizeMacDirectory(string directoryPath)
    {
        var descriptor = OpenMac(directoryPath, OpenReadOnly);
        if (descriptor < 0)
        {
            var error = Marshal.GetLastPInvokeError();
            return MacFailure("open parent directory", error);
        }

        try
        {
            while (FsyncMac(descriptor) != 0)
            {
                var error = Marshal.GetLastPInvokeError();
                if (error == InterruptedSystemCall)
                {
                    continue;
                }

                return MacFailure("fsync parent directory", error);
            }

            return new DurabilityResult(true);
        }
        finally
        {
            _ = CloseMac(descriptor);
        }
    }

    private static DurabilityResult FullFsyncMacFile(string filePath)
    {
        // Use a writable descriptor for a write durability operation. If the
        // destination has become read-only, fail closed rather than approving SD reuse.
        var descriptor = OpenMac(filePath, OpenReadWrite);
        if (descriptor < 0)
        {
            var error = Marshal.GetLastPInvokeError();
            return MacFailure("open finalized file for durable synchronization", error);
        }

        try
        {
            while (FcntlMac(descriptor, FFullFsync) != 0)
            {
                var error = Marshal.GetLastPInvokeError();
                if (error == InterruptedSystemCall)
                {
                    continue;
                }

                return MacFailure("F_FULLFSYNC finalized file", error);
            }

            return new DurabilityResult(true);
        }
        finally
        {
            _ = CloseMac(descriptor);
        }
    }

    private static DurabilityResult MacFailure(string operation, int error) =>
        new(false, $"macOS {operation} failed (errno {error}).");

    [LibraryImport("kernel32.dll", EntryPoint = "MoveFileExW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool MoveFileEx(string existingFileName, string newFileName, uint flags);

    [LibraryImport("/usr/lib/libSystem.B.dylib", EntryPoint = "open", SetLastError = true, StringMarshalling = StringMarshalling.Utf8)]
    private static partial int OpenMac(string path, int flags);

    [LibraryImport("/usr/lib/libSystem.B.dylib", EntryPoint = "fsync", SetLastError = true)]
    private static partial int FsyncMac(int fileDescriptor);

    [LibraryImport("/usr/lib/libSystem.B.dylib", EntryPoint = "fcntl", SetLastError = true)]
    private static partial int FcntlMac(int fileDescriptor, int command);

    [LibraryImport("/usr/lib/libSystem.B.dylib", EntryPoint = "close", SetLastError = true)]
    private static partial int CloseMac(int fileDescriptor);
}
