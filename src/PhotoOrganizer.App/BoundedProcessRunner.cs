using System.Diagnostics;

namespace PhotoOrganizer.App;

internal sealed record BoundedProcessResult(
    int ExitCode,
    string StandardOutput,
    string StandardError,
    bool TimedOut);

internal static class BoundedProcessRunner
{
    public static BoundedProcessResult? Run(
        string fileName,
        IEnumerable<string> arguments,
        TimeSpan timeout)
    {
        if (string.IsNullOrWhiteSpace(fileName)) return null;
        if (timeout <= TimeSpan.Zero) return null;

        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            foreach (var argument in arguments)
            {
                process.StartInfo.ArgumentList.Add(argument);
            }

            if (!process.Start()) return null;

            // Drain both pipes concurrently before waiting. This prevents a child with
            // enough stderr/stdout output to fill an OS pipe from deadlocking the parent.
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            var timeoutMilliseconds = timeout.TotalMilliseconds >= int.MaxValue
                ? int.MaxValue
                : Math.Max(1, (int)Math.Ceiling(timeout.TotalMilliseconds));

            if (!process.WaitForExit(timeoutMilliseconds))
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                try { process.WaitForExit(1000); } catch { }
                return new BoundedProcessResult(-1, string.Empty, string.Empty, TimedOut: true);
            }

            return new BoundedProcessResult(
                process.ExitCode,
                stdout.GetAwaiter().GetResult(),
                stderr.GetAwaiter().GetResult(),
                TimedOut: false);
        }
        catch
        {
            return null;
        }
    }
}
