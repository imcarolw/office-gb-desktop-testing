using System.Text.Json;

namespace OfficeGbAutomation.Common;

public sealed class StepResult
{
    public int Number { get; init; }
    public string Action { get; init; } = "";
    public DateTime StartedUtc { get; init; } = DateTime.UtcNow;
    public DateTime? CompletedUtc { get; set; }
    public string Status { get; set; } = "Running";
    public List<string> Screenshots { get; } = [];
    public string? Error { get; set; }
    public string? ScreenshotError { get; set; }
}

public sealed class CaseResult
{
    public string Application { get; init; } = "";
    public int CaseId { get; init; }
    public int ExpectedStepCount { get; init; }
    public string SourceSheet { get; init; } = "";
    public DateTime StartedUtc { get; init; } = DateTime.UtcNow;
    public DateTime? CompletedUtc { get; set; }
    public string Status { get; set; } = "Running";
    public string VisualStatus { get; } = "ReviewRequired";
    public string? Error { get; set; }
    public List<StepResult> Steps { get; } = [];
    public Dictionary<string, string> Observations { get; } = [];
}

public sealed class Evidence
{
    private readonly Action<string> capture;
    private StepResult? current;
    public string DirectoryPath { get; }
    public CaseResult Result { get; }

    public Evidence(string directoryPath, string application, int caseId, Action<string> capture, int expectedStepCount = 0)
    {
        DirectoryPath = directoryPath;
        this.capture = capture;
        Directory.CreateDirectory(directoryPath);
        Result = new CaseResult
        {
            Application = application, CaseId = caseId, SourceSheet = $"{application} for Win32",
            ExpectedStepCount = expectedStepCount
        };
        Persist();
    }

    public void Step(int number, string action, Action execute)
    {
        if (current is not null || number != Result.Steps.Count + 1)
            throw new InvalidOperationException("Steps must execute sequentially, starting at 1.");
        current = new StepResult { Number = number, Action = action };
        Result.Steps.Add(current);
        Persist();
        Exception? actionError = null;
        try
        {
            execute();
            current.Status = "AutomatedChecksPassed";
        }
        catch (Exception error)
        {
            actionError = error;
            current.Status = error is CaseBlockedException ? "Blocked" : "Failed";
            current.Error = error.ToString();
            throw;
        }
        finally
        {
            try
            {
                Snapshot(actionError is null ? "after" : "failure");
            }
            catch (Exception captureError)
            {
                current.Status = "Failed";
                current.ScreenshotError = captureError.ToString();
                if (actionError is null)
                    throw;
                NUnit.Framework.TestContext.Progress.WriteLine($"Screenshot capture also failed: {captureError}");
            }
            finally
            {
                current.CompletedUtc = DateTime.UtcNow;
                current = null;
                Persist();
            }
        }
    }

    public void Snapshot(string checkpoint)
    {
        if (current is null)
            throw new InvalidOperationException("Screenshots must belong to a running step.");
        if (string.IsNullOrWhiteSpace(checkpoint) ||
            checkpoint.Any(c => !char.IsAsciiLetterOrDigit(c) && c != '-'))
            throw new ArgumentException("Checkpoint names may contain only ASCII letters, digits, and hyphens.");
        var name = $"step-{current.Number:D2}-{current.Screenshots.Count + 1:D2}-{checkpoint}.png";
        var path = Path.Combine(DirectoryPath, name);
        capture(path);
        if (!File.Exists(path) || new FileInfo(path).Length == 0)
            throw new IOException($"Screenshot capture did not produce a nonempty file: {path}");
        current.Screenshots.Add(name);
        Persist();
    }

    public void Complete(Exception? error)
    {
        Result.CompletedUtc = DateTime.UtcNow;
        Result.Status = error is null ? "NeedsVisualReview"
            : error is CaseBlockedException ? "Blocked" : "Failed";
        Result.Error = error?.ToString();
        Persist();
    }

    public void Persist()
    {
        var path = Path.Combine(DirectoryPath, "result.json");
        var temporary = path + ".tmp";
        File.WriteAllText(temporary, JsonSerializer.Serialize(Result, Configuration.JsonOptions));
        File.Move(temporary, path, overwrite: true);
    }
}
