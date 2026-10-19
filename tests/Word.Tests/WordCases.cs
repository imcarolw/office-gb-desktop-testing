using System.IO.Compression;
using System.Xml.Linq;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using NUnit.Framework;
using OfficeGbAutomation.Common;

namespace Word.Tests;

public sealed class WordCases : OfficeTestBase
{
    [Test, Category("Interactive"), Category("Case5842")]
    public void Case5842_Template_search_displays_GB_characters() =>
        RunCase("Word", 5842, 4, (office, evidence) =>
        {
            evidence.Step(1, "Launch Word.", office.Launch);
            evidence.Step(2, "Select the Blank document.", office.NewBlank);
            evidence.Step(3, "Click File -> New.", () => office.Backstage("New"));
            evidence.Step(4, "Fill GB18030 characters in templates search field.", () =>
                office.Enter(office.TemplateSearch(), Samples.SearchText));
        });

    [Test, Category("Interactive"), Category("Case5844")]
    public void Case5844_Save_As_preserves_GB_content_and_filename() =>
        RunCase("Word", 5844, 6, (office, evidence) =>
        {
            var name = $"{Samples.FileNameStem}-{Guid.NewGuid():N}.docx";
            var documentPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), name);
            evidence.Result.Observations["DocumentPath"] = documentPath;
            evidence.Step(1, "Launch Word.", office.Launch);
            evidence.Step(2, "Select the Blank document.", office.NewBlank);
            evidence.Step(3, "Fill GB18030 characters into the blank edit area.", () =>
            {
                office.Find(e => OfficeSession.Class(e) == "_WwG", "Word document surface").Focus();
                UnicodeInput.Type(Samples.Text);
                office.Wait(() => office.ReadModel((scope, app) =>
                {
                    dynamic document = scope.Track(app.ActiveDocument);
                    dynamic content = scope.Track(document.Content);
                    return ((string)content.Text).TrimEnd('\r', '\a') == Samples.Text;
                }), "GB text in Word document");
            });
            evidence.Step(4, "Click File -> Save As -> Browse -> Desktop.", () =>
            {
                office.Backstage("Save As");
                evidence.Snapshot("save-as");
                office.BrowseToDesktop();
            });
            evidence.Step(5, "Fill GB18030 characters in File name text box.", () =>
                office.Enter(office.FileNameField(), name));
            evidence.Step(6, "Click Save button.", () =>
            {
                office.ClickDialogButton("Save", "\u4fdd\u5b58(S)", "\u4fdd\u5b58");
                office.Wait(() => File.Exists(documentPath) &&
                    office.Window.Title.Contains(Path.GetFileNameWithoutExtension(name), StringComparison.Ordinal),
                    "saved Word document and title");
                office.Wait(() => office.ReadModel((scope, app) =>
                {
                    dynamic document = scope.Track(app.ActiveDocument);
                    return (bool)document.Saved;
                }), "Word save operation to finish");
                Keyboard.Type(VirtualKeyShort.ESCAPE);
                using var file = new FileStream(documentPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var archive = new ZipArchive(file, ZipArchiveMode.Read, leaveOpen: true);
                var part = archive.GetEntry("word/document.xml")
                    ?? throw new InvalidDataException("Saved file contains no Word document part.");
                using var stream = part.Open();
                var xml = XDocument.Load(stream);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var text = string.Concat(xml.Descendants(w + "t").Select(t => t.Value));
                Assert.That(text, Is.EqualTo(Samples.Text), "Saved Unicode content changed.");
                Assert.That(Path.GetFileName(documentPath), Is.EqualTo(name));
                file.Position = 0;
                using var copy = File.Create(Path.Combine(evidence.DirectoryPath, name));
                file.CopyTo(copy);
            });
        });
}
