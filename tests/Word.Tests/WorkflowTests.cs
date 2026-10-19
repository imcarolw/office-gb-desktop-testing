using System.Diagnostics;
using NUnit.Framework;
using OfficeGbAutomation.Common;

namespace Word.Tests;

[TestFixture, Category("Infrastructure")]
public sealed class WorkflowTests
{
    [Test]
    public async Task Catalog_routing_and_evidence_validation_work_without_Office()
    {
        var start = new ProcessStartInfo("pwsh")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            WorkingDirectory = Path.GetTempPath()
        };
        start.ArgumentList.Add("-NoProfile");
        start.ArgumentList.Add("-File");
        start.ArgumentList.Add(Path.Combine(Configuration.SolutionRoot(), "scripts", "Test-Workflow.ps1"));
        using var process = Process.Start(start)
            ?? throw new InvalidOperationException("Could not start PowerShell 7 for workflow checks.");
        var output = process.StandardOutput.ReadToEndAsync();
        var errors = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        try
        {
            await process.WaitForExitAsync(timeout.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("PowerShell workflow checks exceeded two minutes.");
        }
        var log = await output + Environment.NewLine + await errors;
        TestContext.Progress.WriteLine(log);
        Assert.That(process.ExitCode, Is.Zero, log);
    }
}
