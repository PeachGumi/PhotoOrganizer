using PhotoOrganizer.Core;

namespace PhotoOrganizer.FaultInjection.Worker;

public static class WorkerAssemblyMarker;

public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        if (args.Length != 4)
        {
            Console.Error.WriteLine("Usage: <mode> <source> <destination-directory> <checkpoint-file>");
            return 64;
        }

        var mode = args[0];
        var source = Path.GetFullPath(args[1]);
        var destination = Path.GetFullPath(args[2]);
        var checkpoint = Path.GetFullPath(args[3]);

        Directory.CreateDirectory(destination);
        Directory.CreateDirectory(Path.GetDirectoryName(checkpoint)!);

        IFileDurabilityService durability = mode switch
        {
            "before-finalize" => new CheckpointDurabilityService(checkpoint, CheckpointMode.BeforeFinalize),
            "after-finalize-before-durable" => new CheckpointDurabilityService(checkpoint, CheckpointMode.AfterFinalizeBeforeDurable),
            "after-durable-before-return" => new CheckpointDurabilityService(checkpoint, CheckpointMode.AfterDurableBeforeReturn),
            "duplicate-durability" => new CheckpointDurabilityService(checkpoint, CheckpointMode.DuplicateDurability),
            _ => throw new ArgumentOutOfRangeException(nameof(args), mode, "Unknown fault-injection mode.")
        };

        var result = await new SafeCopyService(durability).CopyAsync(source, destination).ConfigureAwait(false);
        Console.WriteLine($"{result.Status}|{result.DestinationPath}|{result.Error}");
        return result.Status == CopyStatus.Failed ? 2 : 0;
    }

    private enum CheckpointMode
    {
        BeforeFinalize,
        AfterFinalizeBeforeDurable,
        AfterDurableBeforeReturn,
        DuplicateDurability
    }

    private sealed class CheckpointDurabilityService : IFileDurabilityService
    {
        private readonly string _checkpointPath;
        private readonly CheckpointMode _mode;
        private readonly PlatformFileDurabilityService _platform = new();

        public CheckpointDurabilityService(string checkpointPath, CheckpointMode mode)
        {
            _checkpointPath = checkpointPath;
            _mode = mode;
        }

        public FinalizeFileResult FinalizeNewFile(string temporaryPath, string finalPath, DateTime lastWriteUtc)
        {
            if (_mode == CheckpointMode.BeforeFinalize)
            {
                StopAtCheckpoint("before-finalize");
            }

            if (_mode == CheckpointMode.AfterFinalizeBeforeDurable)
            {
                try
                {
                    File.Move(temporaryPath, finalPath, overwrite: false);
                    File.SetLastWriteTimeUtc(finalPath, lastWriteUtc);
                    StopAtCheckpoint("after-finalize-before-durable");
                    throw new UnreachableException();
                }
                catch (IOException) when (File.Exists(finalPath) || Directory.Exists(finalPath))
                {
                    return new FinalizeFileResult(FinalizeFileStatus.DestinationExists, false);
                }
            }

            if (_mode == CheckpointMode.AfterDurableBeforeReturn)
            {
                var result = _platform.FinalizeNewFile(temporaryPath, finalPath, lastWriteUtc);
                if (result.Status == FinalizeFileStatus.Committed)
                {
                    StopAtCheckpoint("after-durable-before-return");
                }

                return result;
            }

            return _platform.FinalizeNewFile(temporaryPath, finalPath, lastWriteUtc);
        }

        public DurabilityResult EnsureDurable(string filePath)
        {
            if (_mode == CheckpointMode.DuplicateDurability)
            {
                StopAtCheckpoint("duplicate-durability");
            }

            return _platform.EnsureDurable(filePath);
        }

        private void StopAtCheckpoint(string checkpoint)
        {
            File.WriteAllText(_checkpointPath, checkpoint);
            using var gate = new ManualResetEventSlim(false);
            gate.Wait();
        }
    }
}
