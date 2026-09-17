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
            var timeoutMilliseconds = timeout.TotalMilliseconds >= int.MaxValue
                ? int.MaxValue
                : Math.Max(1, (int)Math.Ceiling(timeout.TotalMilliseconds));
            using var deadline = new CancellationTokenSource(timeoutMilliseconds);
            var stdout = process.StandardOutput.ReadToEndAsync(deadline.Token);
            var stderr = process.StandardError.ReadToEndAsync(deadline.Token);

            try
            {
                if (!process.WaitForExit(timeoutMilliseconds))
                {
                    deadline.Cancel();
                    throw new OperationCanceledException(deadline.Token);
                }

                // A child can exit while its descendants still hold the inherited
                // pipes open. The same deadline must cover EOF, not just process exit.
                var output = Task.WhenAll(stdout, stderr)
                    .WaitAsync(deadline.Token).GetAwaiter().GetResult();
                return new BoundedProcessResult(
                    process.ExitCode,
                    output[0],
                    output[1],
                    TimedOut: false);
            }
            catch (OperationCanceledException) when (deadline.IsCancellationRequested)
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                try { process.WaitForExit(1000); } catch { }
                return new BoundedProcessResult(-1, string.Empty, string.Empty, TimedOut: true);
            }
        }
        catch
        {
            return null;
        }
    }
}
