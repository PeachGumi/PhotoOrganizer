using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PhotoOrganizer.Core;

public enum DurableCommitStatus
{
    Committed,
    DestinationExists
}

public interface IDurableFileCommitter
{
    DurableCommitStatus CommitNoReplace(string temporaryPath, string finalPath);
    void EnsureDurable(string filePath);
}

public sealed class PlatformDurableFileCommitter : IDurableFileCommitter
{
    private const uint MoveFileWriteThrough = 0x00000008;
    private const int ErrorFileExists = 80;
    private const int ErrorAlreadyExists = 183;
    private const int FFullFsync = 51;
    private const int OpenReadOnly = 0;

    public DurableCommitStatus CommitNoReplace(string temporaryPath, string finalPath)
    {
        if (OperatingSystem.IsWindows())
        {
            return CommitWindows(temporaryPath, finalPath);
        }

        if (OperatingSystem.IsMacOS())
        {
            return CommitMacOS(temporaryPath, finalPath);
        }

        // Photo Organizer currently ships only for Windows/macOS. Keep any other
        // development host fail-closed rather than silently weakening durability.
        throw new PlatformNotSupportedException("Durable destination commit is implemented only for Windows and macOS.");
    }

    public void EnsureDurable(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("Destination file disappeared before durability verification.", filePath);
        }

        if (OperatingSystem.IsWindows())
        {
            FlushWindowsFile(filePath);
            return;
        }

        if (OperatingSystem.IsMacOS())
        {
            // Synchronize the name/directory entry before the full device-cache flush.
            SyncDirectory(Path.GetDirectoryName(filePath)
                ?? throw new IOException("Destination file has no parent directory."));
            FullSyncFile(filePath);
            return;
        }

        throw new PlatformNotSupportedException("Durability verification is implemented only for Windows and macOS.");
    }

    private static DurableCommitStatus CommitWindows(string temporaryPath, string finalPath)
    {
        // No REPLACE_EXISTING flag is supplied. MOVEFILE_WRITE_THROUGH makes the
        // rename/move operation wait until it is committed to disk before returning.
        if (!MoveFileEx(temporaryPath, finalPath, MoveFileWriteThrough))
        {
            var error = Marshal.GetLastPInvokeError();
            if (error is ErrorFileExists or ErrorAlreadyExists)
            {
                return DurableCommitStatus.DestinationExists;
            }

            throw new IOException(
                $"Durable Windows finalization failed ({error}): {new Win32Exception(error).Message}");
        }

        FlushWindowsFile(finalPath);
        return DurableCommitStatus.Committed;
    }

    private static DurableCommitStatus CommitMacOS(string temporaryPath, string finalPath)
    {
        try
        {
            File.Move(temporaryPath, finalPath, overwrite: false);
        }
        catch (IOException) when (File.Exists(finalPath))
        {
            return DurableCommitStatus.DestinationExists;
        }

        // Failure after rename is intentionally fail-closed without deleting the
        // finalized bytes. A later import can verify or create another safe copy.
        SyncDirectory(Path.GetDirectoryName(finalPath)
            ?? throw new IOException("Final destination has no parent directory."));
        FullSyncFile(finalPath);
        return DurableCommitStatus.Committed;
    }

    private static void FlushWindowsFile(string path)
    {
        // GENERIC_WRITE is required for FlushFileBuffers. FileStream.Flush(true)
        // issues that OS durability request while leaving file bytes unchanged.
        using var file = new FileStream(
            path,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.Read,
            bufferSize: 1,
            FileOptions.WriteThrough);
        file.Flush(flushToDisk: true);
    }

    private static void FullSyncFile(string path)
    {
        using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 1,
            FileOptions.SequentialScan);
        var descriptor = stream.SafeFileHandle.DangerousGetHandle().ToInt32();
        if (Fcntl(descriptor, FFullFsync) != 0)
        {
            throw CreateMacIOException("F_FULLFSYNC failed for destination file");
        }
    }

    private static void SyncDirectory(string directoryPath)
    {
        var descriptor = Open(directoryPath, OpenReadOnly);
        if (descriptor < 0)
        {
            throw CreateMacIOException("Unable to open destination directory for durable sync");
        }

        try
        {
            if (Fsync(descriptor) != 0)
            {
                throw CreateMacIOException("Unable to synchronize destination directory entry");
            }
        }
        finally
        {
            _ = Close(descriptor);
        }
    }

    private static IOException CreateMacIOException(string message)
    {
        var error = Marshal.GetLastPInvokeError();
        return new IOException($"{message} ({error}): {new Win32Exception(error).Message}");
    }

    [DllImport("kernel32.dll", EntryPoint = "MoveFileExW", SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool MoveFileEx(string existingFileName, string newFileName, uint flags);

    [DllImport("libSystem.B.dylib", EntryPoint = "fcntl", SetLastError = true)]
    private static extern int Fcntl(int fileDescriptor, int command);

    [DllImport("libSystem.B.dylib", EntryPoint = "open", SetLastError = true)]
    private static extern int Open(string path, int flags);

    [DllImport("libSystem.B.dylib", EntryPoint = "fsync", SetLastError = true)]
    private static extern int Fsync(int fileDescriptor);

    [DllImport("libSystem.B.dylib", EntryPoint = "close", SetLastError = true)]
    private static extern int Close(int fileDescriptor);
}
