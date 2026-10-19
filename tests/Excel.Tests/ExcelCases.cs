using System.Drawing;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using NUnit.Framework;
using OfficeGbAutomation.Common;

namespace Excel.Tests;

public sealed class ExcelCases : OfficeTestBase
{
    [Test, Category("Interactive"), Category("Case6841")]
    public void Case6841_Advanced_options_display_GB_workbook_and_sheet_names() =>
        RunCase("Excel", 6841, 8, (office, evidence) =>
        {
            var path = Path.Combine(evidence.DirectoryPath, Samples.FileNameStem + ".xlsx");
            evidence.Step(1, "Launch Excel and create a blank worksheet.", () =>
            {
                ComScope.CreateExcelFixture(path);
                office.Launch();
                office.NewBlank();
            });
            evidence.Step(2, "Click File -> Open.", () => office.Backstage("Open"));
            evidence.Step(3, "From PC, choose a file with a GB18030 filename.", () => office.OpenFile(path));
            evidence.Step(4, "Save; check workbook name in AutoRecover exceptions for.", () =>
            {
                Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_S);
                office.Backstage("Options");
                office.Click("Save options category", "Save", "\u4fdd\u5b58");
                InspectDropdown(office, evidence, Path.GetFileName(path),
                    "AutoRecover exceptions for:", "AutoRecover exceptions for",
                    "\u81ea\u52a8\u6062\u590d\u4f8b\u5916\u9879");
            });
            evidence.Step(5, "Change sheet names; check Display options for this worksheet.", () =>
            {
                Keyboard.Type(VirtualKeyShort.ESCAPE);
                var original = office.ReadModel((scope, app) =>
                {
                    dynamic workbook = scope.Track(app.ActiveWorkbook);
                    dynamic sheets = scope.Track(workbook.Worksheets);
                    dynamic sheet = scope.Track(sheets[1]);
                    return (string)sheet.Name;
                });
                var tab = office.Named("worksheet tab", original);
                tab.DoubleClick();
                Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
                UnicodeInput.Type(Samples.SheetName);
                Keyboard.Type(VirtualKeyShort.ENTER);
                office.Wait(() => office.Elements().Any(e => OfficeSession.Name(e) == Samples.SheetName), "renamed worksheet tab");
                evidence.Snapshot("renamed-sheet");
                Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_S);
                OpenAdvanced(office);
                InspectDropdown(office, evidence, Samples.SheetName,
                    "Display options for this worksheet:", "Display options for this worksheet",
                    "\u6b64\u5de5\u4f5c\u8868\u7684\u663e\u793a\u9009\u9879");
            });
            evidence.Step(6, "Click File -> Options -> Advanced.", () =>
            {
                Keyboard.Type(VirtualKeyShort.ESCAPE);
                OpenAdvanced(office);
            });
            evidence.Step(7, "Open Image Size and Quality / Current workbook / workbook display dropdowns.", () =>
            {
                InspectDropdown(office, evidence, Path.GetFileName(path),
                    "Image Size and Quality", "\u56fe\u50cf\u5927\u5c0f\u548c\u8d28\u91cf");
                InspectDropdown(office, evidence, Path.GetFileName(path),
                    "Display options for this workbook:", "Display options for this workbook",
                    "\u6b64\u5de5\u4f5c\u7c3f\u7684\u663e\u793a\u9009\u9879");
            });
            evidence.Step(8, "Check GB18030 characters in the dropdown list.", () =>
            {
                InspectDropdown(office, evidence, Samples.SheetName,
                    "Display options for this worksheet:", "Display options for this worksheet",
                    "\u6b64\u5de5\u4f5c\u8868\u7684\u663e\u793a\u9009\u9879");
            });
            evidence.Result.Observations["StepOrdering"] =
                "Step 4 uses Options -> Save for AutoRecover exceptions; step 5 uses Advanced for sheet display names; step 6 opens Advanced again.";
        });

    [Test, Category("Interactive"), Category("Case5881")]
    public void Case5881_Horizontal_and_vertical_text_boxes_display_GB_characters() =>
        RunCase("Excel", 5881, 4, (office, evidence) =>
        {
            evidence.Step(1, "Launch Excel and select blank workbook.", () =>
            {
                office.Launch();
                office.NewBlank();
            });
            evidence.Step(2, "Click Insert -> Text -> Text Box dropdown.", () => OpenTextBoxMenu(office));
            evidence.Step(3, "Draw Horizontal Text Box; fill GB18030 characters.", () =>
            {
                office.Click("horizontal text box command",
                    "Draw Horizontal Text Box", "Horizontal Text Box",
                    "\u7ed8\u5236\u6a2a\u6392\u6587\u672c\u6846", "\u6a2a\u6392\u6587\u672c\u6846");
                DrawTextBox(office, evidence, 0.25, 1);
                evidence.Snapshot("horizontal-empty");
                FocusDrawnTextBox(office, evidence, 1);
                UnicodeInput.Type(Samples.Text);
                Keyboard.Type(VirtualKeyShort.ESCAPE);
                AssertTextBoxes(office, evidence, 1);
            });
            evidence.Step(4, "Draw Vertical Text Box; fill GB18030 characters.", () =>
            {
                OpenTextBoxMenu(office);
                office.Click("vertical text box command",
                    "Vertical Text Box", "Draw Vertical Text Box",
                    "\u7ed8\u5236\u7ad6\u6392\u6587\u672c\u6846", "\u7ad6\u6392\u6587\u672c\u6846");
                DrawTextBox(office, evidence, 0.55, 2);
                evidence.Snapshot("vertical-empty");
                FocusDrawnTextBox(office, evidence, 2);
                UnicodeInput.Type(Samples.Text);
                Keyboard.Type(VirtualKeyShort.ESCAPE);
                AssertTextBoxes(office, evidence, 2);
                var orientations = office.ReadModel((scope, app) =>
                {
                    dynamic sheet = scope.Track(app.ActiveSheet);
                    dynamic shapes = scope.Track(sheet.Shapes);
                    var values = new List<int>();
                    for (var i = 1; i <= (int)shapes.Count; i++)
                    {
                        dynamic shape = scope.Track(shapes[i]);
                        dynamic frame = scope.Track(shape.TextFrame);
                        values.Add((int)frame.Orientation);
                    }
                    return values;
                });
                evidence.Result.Observations["TextBoxOrientations"] = string.Join(", ", orientations);
                Assert.That(orientations, Does.Contain(1), "Horizontal text box orientation is required.");
                Assert.That(orientations.Any(o => o is 2 or 3 or 4 or 5), Is.True,
                    "A vertical text box orientation is required.");
            });
        });

    private static void OpenAdvanced(OfficeSession office)
    {
        office.Backstage("Options");
        office.Click("Advanced options category", "Advanced", "\u9ad8\u7ea7");
        office.Find(e => OfficeSession.Type(e) == ControlType.ComboBox, "Advanced options dropdown");
    }

    private static void OpenTextBoxMenu(OfficeSession office)
    {
        office.Click("Insert ribbon tab", "Insert", "\u63d2\u5165");
        var collapsedTextGroup = office.Elements().FirstOrDefault(e =>
            OfficeSession.Type(e) == ControlType.MenuItem &&
            new[] { "Text", "\u6587\u672c" }.Contains(OfficeSession.Name(e), StringComparer.OrdinalIgnoreCase));
        if (collapsedTextGroup is not null)
            OfficeSession.Activate(collapsedTextGroup);
        var textBox = office.Named("Text Box dropdown", "Text Box", "\u6587\u672c\u6846");
        if (textBox.Patterns.ExpandCollapse.IsSupported)
            textBox.Patterns.ExpandCollapse.Pattern.Expand();
        else
        {
            var arrow = textBox.FindFirstDescendant(cf => cf.ByControlType(ControlType.MenuItem))
                ?? throw new CaseBlockedException(
                    "This Excel build exposes Text Box as a single horizontal command, not the horizontal/vertical dropdown required by case 5881.");
            OfficeSession.Activate(arrow);
        }
    }

    private static void DrawTextBox(OfficeSession office, Evidence evidence, double xFraction, int count)
    {
        var grid = Rectangle.Intersect(office.Window.BoundingRectangle,
            office.Find(e => OfficeSession.Type(e) == ControlType.DataGrid,
                "worksheet drawing surface").BoundingRectangle);
        evidence.Result.Observations["DrawingSurface"] = grid.ToString();
        var start = new Point(grid.Left + (int)(grid.Width * xFraction), grid.Top + (int)(grid.Height * 0.25));
        var end = new Point(start.X + (int)(grid.Width * 0.22), start.Y + (int)(grid.Height * 0.35));
        Mouse.MoveTo(start);
        Mouse.Down(MouseButton.Left);
        try
        {
            Thread.Sleep(100);
            for (var i = 1; i <= 10; i++)
            {
                Mouse.MoveTo(new Point(start.X + (end.X - start.X) * i / 10,
                    start.Y + (end.Y - start.Y) * i / 10));
                Thread.Sleep(25);
            }
        }
        finally { Mouse.Up(MouseButton.Left); }
        office.Wait(() => office.ReadModel((scope, app) =>
        {
            dynamic sheet = scope.Track(app.ActiveSheet);
            dynamic shapes = scope.Track(sheet.Shapes);
            return (int)shapes.Count == count;
        }), $"drawing of text box {count} to finish");
    }

    private static void FocusDrawnTextBox(OfficeSession office, Evidence evidence, int index)
    {
        var name = office.ReadModel((scope, app) =>
        {
            dynamic workbook = scope.Track(app.ActiveWorkbook);
            dynamic sheet = scope.Track(app.ActiveSheet);
            dynamic shapes = scope.Track(sheet.Shapes);
            dynamic shape = scope.Track(shapes[index]);
            evidence.Result.Observations["DisplayDrawingObjects"] = ((int)workbook.DisplayDrawingObjects).ToString();
            evidence.Result.Observations[$"Shape{index}Visible"] = ((int)shape.Visible).ToString();
            return (string)shape.Name;
        });
        var image = office.NamedControl(ControlType.Image, "drawn text box " + name, name);
        evidence.Result.Observations[$"Shape{index}ScreenBounds"] = image.BoundingRectangle.ToString();
        if (image.IsOffscreen)
            throw new CaseBlockedException($"Drawn text box '{name}' is offscreen; check Excel object-display settings.");
        image.DoubleClick();
    }

    private void AssertTextBoxes(OfficeSession office, Evidence evidence, int expectedCount)
    {
        office.Wait(() => office.ReadModel((scope, app) =>
        {
            dynamic sheet = scope.Track(app.ActiveSheet);
            dynamic shapes = scope.Track(sheet.Shapes);
            var texts = new List<string>();
            for (var i = 1; i <= (int)shapes.Count; i++)
            {
                dynamic shape = scope.Track(shapes[i]);
                if ((int)shape.Type != 17)
                    continue;
                dynamic frame = scope.Track(shape.TextFrame);
                dynamic characters = scope.Track(frame.Characters());
                texts.Add((string)characters.Text);
                evidence.Result.Observations[$"TextBox{i}"] =
                    $"Text: {characters.Text}; Left: {shape.Left}; Top: {shape.Top}; Width: {shape.Width}; Height: {shape.Height}";
            }
            return texts.Count == expectedCount && texts.All(t => t == Samples.Text);
        }), $"{expectedCount} text boxes with exact GB text");
    }

    private static void InspectDropdown(OfficeSession office, Evidence evidence, string expected, params string[] labels)
    {
        var combo = office.Find(e => OfficeSession.Type(e) == ControlType.ComboBox && labels.Any(l =>
            OfficeSession.Name(e).TrimEnd(':', ' ', '\uff1a').Equals(l.TrimEnd(':', ' ', '\uff1a'),
                StringComparison.OrdinalIgnoreCase)), string.Join("/", labels));
        if (combo.Patterns.ScrollItem.IsSupported)
            combo.Patterns.ScrollItem.Pattern.ScrollIntoView();
        combo.Focus();
        combo.AsComboBox().Expand();
        office.Wait(() => office.Elements().Any(e =>
            OfficeSession.Type(e) == ControlType.ListItem && OfficeSession.Name(e) == expected && !e.IsOffscreen),
            $"'{expected}' as a visible list item in expanded dropdown");
        evidence.Snapshot("dropdown");
        Keyboard.Type(VirtualKeyShort.ESCAPE);
    }
}
