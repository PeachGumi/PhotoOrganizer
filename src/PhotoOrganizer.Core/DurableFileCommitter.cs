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

        // Flush the finalized file handle as an independent final barrier. Metadata
        // such as the source modification time is applied to the temporary file before
        // this method is called, so no application metadata writes occur afterward.
        using var final = new FileStream(
            finalPath,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.Read,
            bufferSize: 1,
            FileOptions.WriteThrough);
        final.Flush(flushToDisk: true);
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

        // First synchronize the directory entry created by the rename. Then use
        // F_FULLFSYNC on the finalized file. Apple's fsync(2) documentation notes
        // that plain fsync can leave data in drive write caches; F_FULLFSYNC also
        // asks the device to flush buffered data to permanent storage. Failure is
        // fatal to reuse approval, but the finalized file is deliberately preserved.
        SyncDirectory(Path.GetDirectoryName(finalPath)
            ?? throw new IOException("Final destination has no parent directory."));
        FullSyncFile(finalPath);
        return DurableCommitStatus.Committed;
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
            throw CreateMacIOException("F_FULLFSYNC failed for finalized destination file");
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
