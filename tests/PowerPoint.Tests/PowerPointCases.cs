using System.Drawing;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using NUnit.Framework;
using OfficeGbAutomation.Common;

namespace PowerPoint.Tests;

public sealed class PowerPointCases : OfficeTestBase
{
    [Test, Category("Interactive"), Category("Case5921")]
    public void Case5921_Find_and_replace_preserve_GB_characters() =>
        RunCase("PowerPoint", 5921, 12, (office, evidence) =>
        {
            evidence.Step(1, "Launch PowerPoint.", office.Launch);
            evidence.Step(2, "Create a blank document.", office.NewBlank);
            evidence.Step(3, "Input GB18030 characters in the main input field.", () =>
            {
                ClickPlaceholder(office, 1);
                UnicodeInput.Type(Samples.Text);
                Assert.That(TitleText(office), Is.EqualTo(Samples.Text));
            });
            evidence.Step(4, "Go to Home -> Editing -> Find.", () => OpenEditingDialog(office, "Find"));
            evidence.Step(5, "Input existing GB characters in Find what.", () =>
                office.Enter(DialogField(office, "Find what:", "Find what"), Samples.SearchText));
            evidence.Step(6, "Click Find Next.", () => FindNext(office));
            evidence.Step(7, "Verify the GB characters can be found.", () => VerifySelection(office, Samples.SearchText));
            evidence.Step(8, "Switch directly from Find to Replace, or use Home -> Editing -> Replace.", () =>
            {
                var combinedPane = office.Elements().FirstOrDefault(e =>
                    OfficeSession.Type(e) == ControlType.Custom && OfficeSession.Name(e) == "Find and Replace");
                if (combinedPane is not null)
                {
                    DialogField(office, "Replace with:", "Replace with", "Replace what:", "Replace what").Focus();
                    evidence.Result.Observations["ReplaceNavigation"] =
                        "The modern combined Find and Replace pane already exposes Replace; focus its replacement field without closing the pane.";
                    return;
                }
                CloseFindPanel(office);
                ClickPlaceholder(office, 1);
                Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.HOME);
                OpenEditingDialog(office, "Replace");
            });
            evidence.Step(9, "Input existing GB characters in Find what; click Find Next.", () =>
            {
                office.Enter(DialogField(office, "Find what:", "Find what"), Samples.SearchText);
                evidence.Snapshot("find-query");
                FindNext(office);
            });
            evidence.Step(10, "Verify the GB characters can be found.", () => VerifySelection(office, Samples.SearchText));
            evidence.Step(11, "Input GB characters in Replace with; click Replace.", () =>
            {
                office.Enter(DialogField(office, "Replace with:", "Replace with", "Replace what:", "Replace what"),
                    Samples.FileNameStem);
                evidence.Snapshot("replacement-query");
                office.Click("Replace button", "Replace", "\u66ff\u6362");
            });
            evidence.Step(12, "Verify characters were replaced and display correctly.", () =>
            {
                CloseFindPanel(office);
                var expected = Samples.Text.Replace(Samples.SearchText, Samples.FileNameStem, StringComparison.Ordinal);
                office.Wait(() => TitleText(office) == expected, "exact GB replacement in title");
                evidence.Result.Observations["ExpectedAfterReplacement"] = expected;
                evidence.Result.Observations["ActualAfterReplacement"] = TitleText(office);
                evidence.Result.Observations["ModernUiMapping"] =
                    "Current Office uses a combined Find and Replace pane; Enter in its Find field is the Find Next equivalent.";
            });
        });

    private static void OpenEditingDialog(OfficeSession office, string command)
    {
        office.Click("Home ribbon tab", "Home", "\u5f00\u59cb");
        var collapsed = office.Elements().FirstOrDefault(e => OfficeSession.Type(e) == ControlType.MenuItem
            && OfficeSession.Name(e) == "Editing");
        if (collapsed is not null)
            OfficeSession.Activate(collapsed);
        var combined = office.Elements().FirstOrDefault(e => OfficeSession.Id(e) == "FindDialog");
        if (combined is not null)
            OfficeSession.Activate(combined);
        else
            office.Click(command + " command", command,
                command == "Find" ? "\u67e5\u627e" : "\u66ff\u6362");
        DialogField(office, "Find what:", "Find what");
    }

    private static AutomationElement DialogField(OfficeSession office, params string[] names)
    {
        AutomationElement? field = null;
        office.Wait(() =>
        {
            field = office.Elements().FirstOrDefault(e => OfficeSession.Type(e) == ControlType.Edit
                && names.Contains(OfficeSession.Name(e), StringComparer.OrdinalIgnoreCase));
            if (field is not null)
                return true;
            var pane = office.Elements().FirstOrDefault(e => OfficeSession.Type(e) == ControlType.Custom
                && OfficeSession.Name(e) == "Find and Replace");
            if (pane is null)
                return false;
            var fields = pane.FindAllDescendants().Where(e => OfficeSession.Type(e) == ControlType.Edit)
                .OrderBy(e => e.BoundingRectangle.Top).ToArray();
            if (fields.Length != 2)
                throw new InvalidOperationException("Modern Find and Replace pane must expose exactly two edit fields.");
            field = fields[names[0].StartsWith("Find", StringComparison.Ordinal) ? 0 : 1];
            return true;
        }, string.Join("/", names) + " field");
        return field!;
    }

    private static void FindNext(OfficeSession office)
    {
        var next = office.Elements().FirstOrDefault(e => OfficeSession.Type(e) == ControlType.Button
            && OfficeSession.Name(e).Replace("&", "") == "Find Next");
        if (next is not null)
            OfficeSession.Activate(next);
        else
        {
            DialogField(office, "Find what:", "Find what").Focus();
            Keyboard.Type(VirtualKeyShort.ENTER);
        }
    }

    private static void CloseFindPanel(OfficeSession office)
    {
        var pane = office.Elements().FirstOrDefault(e => OfficeSession.Type(e) == ControlType.Window
            && OfficeSession.Class(e) == "MsoWorkPane" && OfficeSession.Name(e) == "Find and Replace");
        if (pane is null)
        {
            Keyboard.Type(VirtualKeyShort.ESCAPE);
            return;
        }
        var host = office.Find(e => OfficeSession.Class(e) == "NetUINativeHWNDHost"
            && OfficeSession.Name(e) == "Find and Replace", "Find and Replace pane host");
        var close = host.FindAllDescendants().FirstOrDefault(e =>
            OfficeSession.Type(e) == ControlType.Button && OfficeSession.Name(e) == "Close pane")
            ?? throw new InvalidOperationException("Find and Replace pane has no Close button.");
        OfficeSession.Activate(close);
        office.Wait(() => !office.Elements().Any(e => OfficeSession.Class(e) == "NetUINativeHWNDHost"
            && OfficeSession.Name(e) == "Find and Replace"), "Find and Replace pane to close");
    }

    private static string TitleText(OfficeSession office) => office.ReadModel((scope, app) =>
    {
        dynamic presentation = scope.Track(app.ActivePresentation);
        dynamic slides = scope.Track(presentation.Slides);
        dynamic slide = scope.Track(slides[1]);
        dynamic shapes = scope.Track(slide.Shapes);
        dynamic title = scope.Track(shapes.Title);
        dynamic frame = scope.Track(title.TextFrame);
        dynamic range = scope.Track(frame.TextRange);
        return (string)range.Text;
    });

    private static void VerifySelection(OfficeSession office, string expected) =>
        office.Wait(() => office.ReadModel((scope, app) =>
        {
            dynamic window = scope.Track(app.ActiveWindow);
            dynamic selection = scope.Track(window.Selection);
            if ((int)selection.Type != 3)
                return false;
            dynamic range = scope.Track(selection.TextRange);
            return (string)range.Text == expected;
        }), "exact GB text selected by Find Next");

    [Test, Category("Interactive"), Category("Case6295")]
    public void Case6295_Title_and_subtitle_use_expected_default_East_Asian_font() =>
        RunCase("PowerPoint", 6295, 4, (office, evidence) =>
        {
            evidence.Step(1, "Launch PowerPoint.", office.Launch);
            evidence.Step(2, "New a blank presentation.", office.NewBlank);
            evidence.Step(3, "Input GB18030 characters in title and subtitle; check default font.", () =>
            {
                ClickPlaceholder(office, 1);
                UnicodeInput.Type(Samples.Text);
                evidence.Snapshot("title");
                Keyboard.Type(VirtualKeyShort.ESCAPE);
                ClickPlaceholder(office, 4);
                UnicodeInput.Type(Samples.Text);
                Keyboard.Type(VirtualKeyShort.ESCAPE);
                evidence.Snapshot("subtitle");
                CheckTextAndFonts(office, evidence, assertFont: false);
            });
            evidence.Step(4, "Check GB18030 rendering with Deng Xian or Deng Xian Light.", () =>
            {
                office.Window.Focus();
                CheckTextAndFonts(office, evidence);
                evidence.Result.Observations["VisualReview"] =
                    "Inspect title and subtitle for missing glyphs, clipping, and incorrect glyphs; do not change fonts.";
            });
        });

    private static void ClickPlaceholder(OfficeSession office, int placeholderType)
    {
        var point = office.ReadModel((scope, app) =>
        {
            dynamic window = scope.Track(app.ActiveWindow);
            dynamic presentation = scope.Track(app.ActivePresentation);
            dynamic slides = scope.Track(presentation.Slides);
            dynamic slide = scope.Track(slides[1]);
            dynamic shapes = scope.Track(slide.Shapes);
            for (var i = 1; i <= (int)shapes.Count; i++)
            {
                dynamic shape = scope.Track(shapes[i]);
                if ((int)shape.Type != 14)
                    continue;
                dynamic format = scope.Track(shape.PlaceholderFormat);
                if ((int)format.Type != placeholderType && !(placeholderType == 1 && (int)format.Type == 3))
                    continue;
                var x = (int)window.PointsToScreenPixelsX((float)shape.Left + (float)shape.Width / 2);
                var y = (int)window.PointsToScreenPixelsY((float)shape.Top + (float)shape.Height / 2);
                return new Point(x, y);
            }
            throw new InvalidOperationException($"Blank slide has no placeholder of type {placeholderType}.");
        });
        if (!office.Window.BoundingRectangle.Contains(point))
            throw new InvalidOperationException("Placeholder screen coordinates are outside the owned PowerPoint window.");
        office.Window.Focus();
        Mouse.Click(point);
    }

    private void CheckTextAndFonts(OfficeSession office, Evidence evidence, bool assertFont = true)
    {
        var values = office.ReadModel((scope, app) =>
        {
            dynamic presentation = scope.Track(app.ActivePresentation);
            dynamic slides = scope.Track(presentation.Slides);
            dynamic slide = scope.Track(slides[1]);
            dynamic shapes = scope.Track(slide.Shapes);
            var result = new List<(string Text, string Font)>();
            for (var i = 1; i <= (int)shapes.Count; i++)
            {
                dynamic shape = scope.Track(shapes[i]);
                if ((int)shape.HasTextFrame == 0)
                    continue;
                dynamic frame = scope.Track(shape.TextFrame);
                dynamic range = scope.Track(frame.TextRange);
                var text = (string)range.Text;
                if (text != Samples.Text)
                    continue;
                dynamic font = scope.Track(range.Font);
                var eastAsianFont = (string)font.NameFarEast;
                evidence.Result.Observations[$"Shape{i}RawEastAsianFont"] = eastAsianFont;
                if (eastAsianFont.StartsWith("+", StringComparison.Ordinal))
                {
                    dynamic master = scope.Track(slide.Master);
                    dynamic theme = scope.Track(master.Theme);
                    dynamic scheme = scope.Track(theme.ThemeFontScheme);
                    var themePath = Path.Combine(evidence.DirectoryPath, $"shape-{i}-default-fonts.xml");
                    if (!File.Exists(themePath))
                        scheme.Save(themePath);
                    eastAsianFont = ThemeFontResolver.Resolve(eastAsianFont, themePath, "Hans");
                }
                result.Add((text, eastAsianFont));
            }
            return result;
        });
        evidence.Result.Observations["ObservedEastAsianFonts"] = string.Join(", ", values.Select(v => v.Font));
        Assert.That(values.Count, Is.EqualTo(2), "Both title and subtitle must contain the exact sample.");
        if (assertFont)
            foreach (var value in values)
                Assert.That(Settings.ExpectedPowerPointFonts, Does.Contain(value.Font),
                    $"Default East Asian font '{value.Font}' differs from the case expectation; no font was forced.");
    }
}
