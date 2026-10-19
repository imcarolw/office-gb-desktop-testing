using System.Text.Json;
using Microsoft.Win32;

namespace OfficeGbAutomation.Common;

public sealed class OfficeSettings
{
    public string OfficeDirectory { get; init; } = "";
    public string ArtifactRoot { get; init; } = "";
    public int UiTimeoutSeconds { get; init; }
    public string[] ExpectedPowerPointFonts { get; init; } = [];

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(ArtifactRoot))
            throw new InvalidDataException("ArtifactRoot must be a nonempty relative or absolute path.");
        if (UiTimeoutSeconds < 1 || UiTimeoutSeconds > 300)
            throw new InvalidDataException("UiTimeoutSeconds must be between 1 and 300.");
        if (ExpectedPowerPointFonts.Length == 0 || ExpectedPowerPointFonts.Any(string.IsNullOrWhiteSpace))
            throw new InvalidDataException("ExpectedPowerPointFonts must contain nonempty font names.");
    }

    public OfficeSettings ResolvePaths(string solutionRoot) => new()
    {
        OfficeDirectory = string.IsNullOrWhiteSpace(OfficeDirectory) ? ""
            : Path.GetFullPath(OfficeDirectory, solutionRoot),
        ArtifactRoot = Path.GetFullPath(ArtifactRoot, solutionRoot),
        UiTimeoutSeconds = UiTimeoutSeconds,
        ExpectedPowerPointFonts = ExpectedPowerPointFonts
    };

    public string ResolveExecutable(string executableName)
    {
        if (!string.IsNullOrWhiteSpace(OfficeDirectory))
        {
            var path = Path.Combine(OfficeDirectory, executableName + ".EXE");
            return File.Exists(path) ? path
                : throw new FileNotFoundException($"Configured Office executable does not exist: '{path}'.", path);
        }

        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var registry = RegistryKey.OpenBaseKey(hive, view);
            using var key = registry.OpenSubKey(
                $@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{executableName}.EXE");
            if (key?.GetValue("") is not string registeredPath)
                continue;
            var path = Environment.ExpandEnvironmentVariables(registeredPath.Trim().Trim('"'));
            if (Path.IsPathFullyQualified(path) && File.Exists(path))
                return path;
        }
        throw new FileNotFoundException(
            $"Could not discover {executableName}.EXE through Windows App Paths. Set officeDirectory in office-test-settings.json.");
    }
}

public sealed class GbSamples
{
    public string Text { get; init; } = "";
    public string SearchText { get; init; } = "";
    public string FileNameStem { get; init; } = "";
    public string SheetName { get; init; } = "";

    public void Validate()
    {
        if (new[] { Text, SearchText, FileNameStem, SheetName }.Any(string.IsNullOrWhiteSpace))
            throw new InvalidDataException("All GB sample fields must be nonempty.");
        if (FileNameStem.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            throw new InvalidDataException("FileNameStem contains invalid filename characters.");
        if (SheetName.Length > 31 || SheetName.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
            throw new InvalidDataException("SheetName must be a valid Excel worksheet name.");
        foreach (var value in new[] { Text, SearchText, FileNameStem, SheetName })
            UnicodeInput.Validate(value);
    }
}

public static class Configuration
{
    public static string SolutionRoot()
    {
        var configured = Environment.GetEnvironmentVariable("OFFICE_TEST_ROOT");
        if (!string.IsNullOrWhiteSpace(configured))
        {
            if (!Path.IsPathFullyQualified(configured) || !Directory.Exists(configured))
                throw new InvalidDataException("OFFICE_TEST_ROOT must name an existing absolute directory.");
            return Path.GetFullPath(configured);
        }
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
            if (File.Exists(Path.Combine(directory.FullName, "OfficeGbAutomation.slnx")))
                return directory.FullName;
        throw new DirectoryNotFoundException(
            "Cannot locate OfficeGbAutomation.slnx above the test output. Set OFFICE_TEST_ROOT when deploying tests separately.");
    }

    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    public static T Load<T>(string fileName) =>
        JsonSerializer.Deserialize<T>(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, fileName)), JsonOptions)
        ?? throw new InvalidDataException($"Configuration '{fileName}' was null.");
}
