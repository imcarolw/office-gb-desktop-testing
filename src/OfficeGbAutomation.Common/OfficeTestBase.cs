using System.Diagnostics;
using System.Reflection;
using NUnit.Framework;

namespace OfficeGbAutomation.Common;

[Apartment(ApartmentState.STA)]
[NonParallelizable]
public abstract class OfficeTestBase
{
    protected OfficeSettings Settings { get; private set; } = null!;
    protected GbSamples Samples { get; private set; } = null!;

    [Test, Category("Infrastructure")]
    public void Catalog_matches_discovered_interactive_tests()
    {
        var catalog = Configuration.Load<CatalogCase[]>("test-catalog.json");
        var assemblyName = GetType().Assembly.GetName().Name + ".csproj";
        var entries = catalog.Where(c => Path.GetFileName(c.Project) == assemblyName).ToArray();
        var methods = GetType().GetMethods().Where(m =>
            m.GetCustomAttribute<TestAttribute>() is not null &&
            m.GetCustomAttributes<CategoryAttribute>().Any(c => c.Name == "Interactive")).ToArray();
        Assert.That(entries.Select(c => c.TestName), Is.EquivalentTo(methods.Select(m => m.Name)),
            "The shared catalog must match NUnit's interactive test methods.");
        Assert.That(catalog.Select(c => c.CaseId).Distinct().Count(), Is.EqualTo(catalog.Length));
        foreach (var entry in entries)
        {
            var method = methods.Single(m => m.Name == entry.TestName);
            Assert.That(method.GetCustomAttributes<CategoryAttribute>().Select(c => c.Name),
                Does.Contain($"Case{entry.CaseId}"));
            Assert.That(entry.ExpectedSteps, Is.GreaterThan(0));
        }
    }

    protected void RunCase(string appName, int caseId, int expectedSteps,
        Action<OfficeSession, Evidence> execute)
    {
        var entry = Configuration.Load<CatalogCase[]>("test-catalog.json").Single(c => c.CaseId == caseId);
        Assert.That(entry.Application, Is.EqualTo(appName));
        Assert.That(entry.ExpectedSteps, Is.EqualTo(expectedSteps), "Catalog step count differs from the implementation.");
        Assert.That(entry.TestName, Is.EqualTo(TestContext.CurrentContext.Test.MethodName));
        if (Environment.GetEnvironmentVariable("RUN_OFFICE_UI_TESTS") != "1")
            Assert.Ignore("Set RUN_OFFICE_UI_TESTS=1 to run on an unlocked, interactive desktop.");
        Settings = Configuration.Load<OfficeSettings>("office-test-settings.json");
        Samples = Configuration.Load<GbSamples>("gb-samples.json");
        Settings.Validate();
        Settings = Settings.ResolvePaths(Configuration.SolutionRoot());
        Samples.Validate();
        var runId = Environment.GetEnvironmentVariable("OFFICE_UI_RUN_ID")
            ?? $"{DateTime.Now:yyyyMMdd-HHmmss-fff}-{Guid.NewGuid():N}";
        if (runId.Length == 0 || runId.Any(c => !char.IsAsciiLetterOrDigit(c) && c != '-'))
            throw new InvalidDataException("OFFICE_UI_RUN_ID may contain only ASCII letters, digits, and hyphens.");
        var directory = Path.Combine(Settings.ArtifactRoot, runId, appName, caseId.ToString());
        using var dpi = new DesktopDpiScope();
        var session = new OfficeSession(appName, Settings);
        var evidence = new Evidence(directory, appName, caseId, session.Capture, expectedSteps);
        Exception? failure = null;
        FileStream? desktopLock = null;
        try
        {
            Directory.CreateDirectory(Settings.ArtifactRoot);
            desktopLock = new FileStream(Path.Combine(Settings.ArtifactRoot, "desktop.lock"),
                FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            foreach (var executable in new[] { "WINWORD", "EXCEL", "POWERPNT" })
            {
                var processes = Process.GetProcessesByName(executable);
                try
                {
                    if (processes.Length > 0)
                        throw new InvalidOperationException(
                            $"Close all Office applications before testing; {executable} is already running.");
                }
                finally
                {
                    foreach (var process in processes)
                        process.Dispose();
                }
            }
            evidence.Result.Observations["OfficeExecutable"] =
                Settings.ResolveExecutable(OfficeSession.ExecutableName(appName));
            execute(session, evidence);
            Assert.That(evidence.Result.Steps.Count, Is.EqualTo(expectedSteps),
                "Every numbered workbook step must be implemented.");
            Assert.That(evidence.Result.Steps.All(s => s.Screenshots.Count > 0), Is.True,
                "Every step must have screenshot evidence.");
        }
        catch (Exception error)
        {
            failure = error;
            try
            {
                session.DumpUi(Path.Combine(directory, "failure-ui.txt"));
            }
            catch (Exception diagnosticError)
            {
                evidence.Result.Observations["DiagnosticError"] = diagnosticError.ToString();
            }
            if (error is CaseBlockedException)
                Assert.Inconclusive(error.Message);
            throw;
        }
        finally
        {
            try
            {
                session.Dispose();
            }
            catch (Exception cleanupError)
            {
                evidence.Result.Observations["CleanupError"] = cleanupError.ToString();
                if (failure is null)
                {
                    failure = cleanupError;
                    throw;
                }
                TestContext.Progress.WriteLine($"Cleanup also failed: {cleanupError}");
            }
            finally
            {
                if (session.StartupDiagnostics.Count > 0)
                    evidence.Result.Observations["StartupHandling"] = string.Join(Environment.NewLine, session.StartupDiagnostics);
                evidence.Complete(failure);
                desktopLock?.Dispose();
                TestContext.Progress.WriteLine($"Case {caseId}: {evidence.Result.Status}. Evidence: {directory}");
                TestContext.AddTestAttachment(Path.Combine(directory, "result.json"));
            }
        }
    }
}
