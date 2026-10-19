using System.Text.Json;
using System.Xml.Linq;
using NUnit.Framework;
using OfficeGbAutomation.Common;

namespace Word.Tests;

[TestFixture, Category("Infrastructure")]
public sealed class InfrastructureTests
{
    private string directory = "";

    [SetUp]
    public void SetUp() => directory = Path.Combine(Path.GetTempPath(), "OfficeGbEvidence-" + Guid.NewGuid().ToString("N"));

    [TearDown]
    public void TearDown()
    {
        if (Directory.Exists(directory))
            Directory.Delete(directory, recursive: true);
    }

    private Evidence Create(Action<string>? capture = null) =>
        new(directory, "Word", 5842, capture ?? (p => File.WriteAllBytes(p, [1, 2, 3])));

    [Test]
    public void Every_successful_step_has_persistent_screenshot_and_visual_review_status()
    {
        var evidence = Create();
        evidence.Step(1, "Launch", () => { });
        evidence.Step(2, "Type", () => evidence.Snapshot("typed"));
        evidence.Complete(null);
        using var result = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "result.json")));
        Assert.Multiple(() =>
        {
            Assert.That(result.RootElement.GetProperty("Status").GetString(), Is.EqualTo("NeedsVisualReview"));
            Assert.That(result.RootElement.GetProperty("Steps").GetArrayLength(), Is.EqualTo(2));
            Assert.That(evidence.Result.Steps[0].Screenshots, Has.Count.EqualTo(1));
            Assert.That(evidence.Result.Steps[1].Screenshots, Has.Count.EqualTo(2));
            Assert.That(Directory.GetFiles(directory, "*.png"), Has.Length.EqualTo(3));
        });
    }

    [Test]
    public void Failed_action_keeps_original_error_and_failure_screenshot()
    {
        var evidence = Create();
        var error = Assert.Throws<InvalidOperationException>(() =>
            evidence.Step(1, "Launch", () => throw new InvalidOperationException("action failed")));
        evidence.Complete(error);
        Assert.Multiple(() =>
        {
            Assert.That(evidence.Result.Status, Is.EqualTo("Failed"));
            Assert.That(evidence.Result.Steps[0].Error, Does.Contain("action failed"));
            Assert.That(evidence.Result.Steps[0].Screenshots.Single(), Does.Contain("failure"));
        });
    }

    [Test]
    public void Capture_failure_is_not_reported_as_success()
    {
        var evidence = Create(_ => throw new IOException("capture failed"));
        var error = Assert.Throws<IOException>(() => evidence.Step(1, "Launch", () => { }));
        evidence.Complete(error);
        Assert.That(evidence.Result.Steps[0].ScreenshotError, Does.Contain("capture failed"));
        Assert.That(evidence.Result.Status, Is.EqualTo("Failed"));
    }

    [Test]
    public void Blocked_case_is_explicit_and_retains_its_checkpoint_screenshot()
    {
        var evidence = Create();
        var error = Assert.Throws<CaseBlockedException>(() =>
            evidence.Step(1, "Unsupported command", () => throw new CaseBlockedException("command not available")));
        evidence.Complete(error);
        Assert.Multiple(() =>
        {
            Assert.That(evidence.Result.Status, Is.EqualTo("Blocked"));
            Assert.That(evidence.Result.Steps[0].Status, Is.EqualTo("Blocked"));
            Assert.That(evidence.Result.Steps[0].Screenshots, Has.Count.EqualTo(1));
        });
    }

    [Test]
    public void Capture_failure_does_not_mask_action_failure()
    {
        var evidence = Create(_ => throw new IOException("capture failed"));
        Assert.Throws<InvalidOperationException>(() =>
            evidence.Step(1, "Launch", () => throw new InvalidOperationException("action failed")));
        Assert.That(evidence.Result.Steps[0].ScreenshotError, Does.Contain("capture failed"));
        Assert.That(evidence.Result.Steps[0].Error, Does.Contain("action failed"));
    }

    [Test]
    public void Missing_capture_file_is_an_explicit_failure()
    {
        var evidence = Create(_ => { });
        Assert.Throws<IOException>(() => evidence.Step(1, "Launch", () => { }));
    }

    [Test]
    public void Steps_must_be_sequential()
    {
        var evidence = Create();
        Assert.Throws<InvalidOperationException>(() => evidence.Step(2, "Skipped launch", () => { }));
    }

    [Test]
    public void Snapshot_requires_an_active_step()
    {
        var evidence = Create();
        Assert.Throws<InvalidOperationException>(() => evidence.Snapshot("outside"));
    }

    [TestCase(0xD840, "")]
    [TestCase(0xDC00, "")]
    [TestCase(0xD840, "a")]
    public void Unpaired_surrogates_are_rejected(int codeUnit, string suffix) =>
        Assert.Throws<ArgumentException>(() => UnicodeInput.Validate(((char)codeUnit).ToString() + suffix));

    [Test]
    public void Supplementary_GB_samples_and_configuration_are_valid()
    {
        var samples = Configuration.Load<GbSamples>("gb-samples.json");
        var settings = Configuration.Load<OfficeSettings>("office-test-settings.json");
        Assert.DoesNotThrow(samples.Validate);
        Assert.DoesNotThrow(settings.Validate);
        Assert.That(samples.Text, Does.Contain("\ud840\udc00"));
        Assert.That(samples.SheetName.Length, Is.LessThanOrEqualTo(31));
    }

    [Test]
    public void Invalid_worksheet_names_are_rejected() =>
        Assert.Throws<InvalidDataException>(() => new GbSamples
        {
            Text = "text", SearchText = "search", FileNameStem = "file", SheetName = "invalid/name"
        }.Validate());

    [Test]
    public void Relative_paths_resolve_against_the_supplied_root()
    {
        var settings = new OfficeSettings { ArtifactRoot = "artifacts", OfficeDirectory = "custom-office" }
            .ResolvePaths(directory);
        Assert.That(settings.ArtifactRoot, Is.EqualTo(Path.Combine(directory, "artifacts")));
        Assert.That(settings.OfficeDirectory, Is.EqualTo(Path.Combine(directory, "custom-office")));
    }

    [Test]
    public void Absolute_paths_are_preserved_and_empty_Office_directory_enables_discovery()
    {
        var settings = new OfficeSettings { ArtifactRoot = directory }.ResolvePaths(Path.GetTempPath());
        Assert.That(settings.ArtifactRoot, Is.EqualTo(directory));
        Assert.That(settings.OfficeDirectory, Is.Empty);
    }

    [Test]
    public void Invalid_explicit_Office_directory_does_not_fall_back_to_discovery()
    {
        var settings = new OfficeSettings { ArtifactRoot = "artifacts", OfficeDirectory = "missing-office" }
            .ResolvePaths(directory);
        Assert.Throws<FileNotFoundException>(() => settings.ResolveExecutable("WINWORD"));
    }

    [Test]
    public void Direct_test_execution_finds_the_solution_root()
    {
        var root = Configuration.SolutionRoot();
        Assert.That(File.Exists(Path.Combine(root, "OfficeGbAutomation.slnx")), Is.True);
        var settings = Configuration.Load<OfficeSettings>("office-test-settings.json").ResolvePaths(root);
        Assert.That(settings.ArtifactRoot, Is.EqualTo(Path.Combine(root, "artifacts")));
    }

    [TestCase("+mj-ea", "DengXian Light")]
    [TestCase("+mn-ea", "DengXian")]
    public void Theme_font_resolution_uses_Han_script_not_empty_generic_East_Asian_font(string token, string expected)
    {
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "fonts.xml");
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        new XDocument(new XElement(a + "fontScheme",
            new XElement(a + "majorFont", new XElement(a + "ea", new XAttribute("typeface", "")),
                new XElement(a + "font", new XAttribute("script", "Hans"), new XAttribute("typeface", "DengXian Light"))),
            new XElement(a + "minorFont", new XElement(a + "ea", new XAttribute("typeface", "")),
                new XElement(a + "font", new XAttribute("script", "Hans"), new XAttribute("typeface", "DengXian")))))
            .Save(path);
        Assert.That(ThemeFontResolver.Resolve(token, path, "Hans"), Is.EqualTo(expected));
        Assert.Throws<InvalidDataException>(() => ThemeFontResolver.Resolve(token, path, "Unknown"));
    }

    [Test]
    public void Explicit_font_name_is_not_replaced_by_a_theme_default() =>
        Assert.That(ThemeFontResolver.Resolve("CustomFont", "unused.xml", "Hans"), Is.EqualTo("CustomFont"));

    [Test]
    public void Empty_font_name_is_an_explicit_error() =>
        Assert.Throws<InvalidDataException>(() => ThemeFontResolver.Resolve("", "unused.xml", "Hans"));
}
