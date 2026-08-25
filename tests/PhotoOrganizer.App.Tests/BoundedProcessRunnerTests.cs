using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class BoundedProcessRunnerTests
{
    [TestMethod]
    public void SuccessfulCommand_CapturesStdoutAndExitCode()
    {
        var result = RunShell(
            OperatingSystem.IsWindows() ? "echo hello" : "printf hello",
            TimeSpan.FromSeconds(5));

        Assert.IsNotNull(result);
        Assert.IsFalse(result.TimedOut);
        Assert.AreEqual(0, result.ExitCode);
        Assert.AreEqual("hello", result.StandardOutput.Trim());
    }

    [TestMethod]
    public void NonZeroCommand_CapturesStderrAndExitCode()
    {
        var script = OperatingSystem.IsWindows()
            ? "echo problem 1>&2 & exit /b 7"
            : "printf problem >&2; exit 7";

        var result = RunShell(script, TimeSpan.FromSeconds(5));

        Assert.IsNotNull(result);
        Assert.IsFalse(result.TimedOut);
        Assert.AreEqual(7, result.ExitCode);
        StringAssert.Contains(result.StandardError, "problem");
    }

    [TestMethod]
    public void SlowCommand_TimesOutAndReturnsPromptly()
    {
        var script = OperatingSystem.IsWindows()
            ? "ping 127.0.0.1 -n 6 >nul"
            : "sleep 5";
        var stopwatch = Stopwatch.StartNew();

        var result = RunShell(script, TimeSpan.FromMilliseconds(150));

        stopwatch.Stop();
        Assert.IsNotNull(result);
        Assert.IsTrue(result.TimedOut);
        Assert.IsTrue(
            stopwatch.Elapsed < TimeSpan.FromSeconds(4),
            $"Timed command took too long to return: {stopwatch.Elapsed}.");
    }

    [TestMethod]
    public void LargeStderr_DoesNotDeadlockRedirectedPipes()
    {
        var script = OperatingSystem.IsWindows()
            ? "for /L %i in (1,1,12000) do @echo error-line 1>&2"
            : "i=0; while [ $i -lt 12000 ]; do echo error-line >&2; i=$((i+1)); done";

        var result = RunShell(script, TimeSpan.FromSeconds(15));

        Assert.IsNotNull(result);
        Assert.IsFalse(result.TimedOut);
        Assert.AreEqual(0, result.ExitCode);
        Assert.IsTrue(result.StandardError.Length > 64 * 1024);
        StringAssert.Contains(result.StandardError, "error-line");
    }

    [TestMethod]
    public void MissingExecutable_ReturnsNullInsteadOfThrowing()
    {
        var result = BoundedProcessRunner.Run(
            Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "does-not-exist"),
            [],
            TimeSpan.FromSeconds(1));

        Assert.IsNull(result);
    }

    [TestMethod]
    public void NonPositiveTimeout_IsRejected()
    {
        var result = BoundedProcessRunner.Run(
            OperatingSystem.IsWindows() ? "cmd.exe" : "/bin/sh",
            [],
            TimeSpan.Zero);

        Assert.IsNull(result);
    }

    private static BoundedProcessResult? RunShell(string script, TimeSpan timeout)
    {
        if (OperatingSystem.IsWindows())
        {
            return BoundedProcessRunner.Run(
                "cmd.exe",
                ["/d", "/s", "/c", script],
                timeout);
        }

        return BoundedProcessRunner.Run(
            "/bin/sh",
            ["-c", script],
            timeout);
    }
}
