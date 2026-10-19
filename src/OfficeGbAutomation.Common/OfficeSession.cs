using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Capturing;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;

namespace OfficeGbAutomation.Common;

public sealed class OfficeSession : IDisposable
{
    private readonly OfficeSettings settings;
    private readonly string appName;
    private Application? application;
    public List<string> StartupDiagnostics { get; } = [];
    public UIA3Automation Automation { get; } = new();
    public Application Application => application ?? throw new InvalidOperationException("Office has not launched.");
    private string MainWindowClass => appName switch
    {
        "Word" => "OpusApp", "Excel" => "XLMAIN", _ => "PPTFrameClass"
    };
    public Window Window => MainWindow()
        ?? throw new InvalidOperationException("Office has no main window.");

    private Window? MainWindow() => Application.GetAllTopLevelWindows(Automation)
        .FirstOrDefault(w => Class(w) == MainWindowClass);

    public static string Name(AutomationElement element) => element.Properties.Name.ValueOrDefault ?? "";
    public static string Class(AutomationElement element) => element.Properties.ClassName.ValueOrDefault ?? "";
    public static string Id(AutomationElement element) => element.Properties.AutomationId.ValueOrDefault ?? "";
    public static ControlType Type(AutomationElement element) => element.Properties.ControlType.ValueOrDefault;

    public OfficeSession(string appName, OfficeSettings settings)
    {
        this.appName = appName;
        this.settings = settings;
    }

    public static string ExecutableName(string appName) => appName switch
    {
        "Word" => "WINWORD",
        "Excel" => "EXCEL",
        "PowerPoint" => "POWERPNT",
        _ => throw new ArgumentOutOfRangeException(nameof(appName))
    };

    public void Launch()
    {
        var start = new ProcessStartInfo(settings.ResolveExecutable(ExecutableName(appName)))
        {
            UseShellExecute = false
        };
        if (appName == "Word")
        {
            start.ArgumentList.Add("/x");
            start.ArgumentList.Add("/n");
            start.ArgumentList.Add("/q");
        }
        else if (appName == "Excel")
            start.ArgumentList.Add("/x");
        application = Application.Launch(start);
        Wait(() => MainWindow() is not null, "Office main window (not splash screen)");
        Wait(() =>
        {
            Window.Focus();
            if (Window.Patterns.Window.IsSupported && Window.Patterns.Window.Pattern.CanMaximize.Value)
                Window.Patterns.Window.Pattern.SetWindowVisualState(WindowVisualState.Maximized);
            return true;
        }, "Office main window to accept focus");
    }

    public void Wait(Func<bool> condition, string description)
    {
        OfficeStartup.WaitFor(condition, description, TimeSpan.FromSeconds(settings.UiTimeoutSeconds),
            LogStartup);
    }

    private void LogStartup(string message)
    {
        StartupDiagnostics.Add(message);
        NUnit.Framework.TestContext.Progress.WriteLine(message);
    }

    public AutomationElement Find(Func<AutomationElement, bool> predicate, string description)
    {
        AutomationElement? found = null;
        Wait(() =>
        {
            found = Elements().FirstOrDefault(e => e.IsAvailable && predicate(e));
            return found is not null;
        }, description);
        return found!;
    }

    public IEnumerable<AutomationElement> Elements() =>
        Application.GetAllTopLevelWindows(Automation)
            .SelectMany(w => new[] { (AutomationElement)w }.Concat(w.FindAllDescendants()));

    public AutomationElement Named(string description, params string[] names) =>
        Find(e => names.Contains(Name(e), StringComparer.OrdinalIgnoreCase), description);

    public AutomationElement NamedControl(ControlType type, string description, params string[] names) =>
        Find(e => Type(e) == type && names.Contains(Name(e), StringComparer.OrdinalIgnoreCase), description);

    public void Click(string description, params string[] names) => Activate(Named(description, names));

    public static void Activate(AutomationElement element)
    {
        if (element.Patterns.ScrollItem.IsSupported)
            element.Patterns.ScrollItem.Pattern.ScrollIntoView();
        if (element.Patterns.Invoke.IsSupported)
            element.Patterns.Invoke.Pattern.Invoke();
        else if (element.Patterns.ExpandCollapse.IsSupported)
            element.Patterns.ExpandCollapse.Pattern.Expand();
        else if (element.Patterns.SelectionItem.IsSupported)
            element.Patterns.SelectionItem.Pattern.Select();
        else
        {
            if (element.IsOffscreen)
                throw new InvalidOperationException($"'{Name(element)}' is offscreen and has no activation pattern.");
            element.Click();
        }
    }

    public void NewBlank()
    {
        var labels = appName switch
        {
            "Word" => new[] { "Blank document", "\u7a7a\u767d\u6587\u6863" },
            "Excel" => new[] { "Blank workbook", "\u7a7a\u767d\u5de5\u4f5c\u7c3f" },
            _ => new[] { "Blank Presentation", "Blank presentation", "\u7a7a\u767d\u6f14\u793a\u6587\u7a3f" }
        };
        AutomationElement? blank = null;
        Wait(() =>
        {
            DismissCloudOnboarding();
            blank = Elements().FirstOrDefault(e => labels.Contains(Name(e), StringComparer.OrdinalIgnoreCase));
            return true;
        }, "Office start screen inspection");
        if (blank is not null)
            Activate(blank);
        else
        {
            Window.Focus();
            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_N);
        }
        PrepareDocument();
    }

    private void PrepareDocument()
    {
        var stable = Stopwatch.StartNew();
        var wasReady = false;
        var lastState = "Document not yet inspected.";
        try
        {
            Wait(() =>
            {
                DismissCloudOnboarding();
                DismissKnownTaskPanes();
                var elements = Elements().ToArray();
                var file = elements.Any(e => Id(e) == "FileTabButton" && !e.Properties.IsOffscreen.ValueOrDefault);
                var home = elements.Any(e => Type(e) == ControlType.TabItem && !e.Properties.IsOffscreen.ValueOrDefault &&
                    (Name(e) == "Home" || Name(e) == "\u5f00\u59cb"));
                var document = HasEditingDocument();
                lastState = $"Editing document: {document}; File tab: {file}; Home tab: {home}.";
                if (!document || !file || !home)
                {
                    wasReady = false;
                    return false;
                }
                if (!wasReady)
                {
                    stable.Restart();
                    wasReady = true;
                }
                return stable.ElapsedMilliseconds >= 500;
            }, "document editing surface and Home/File ribbon after optional OfficePLUS/Copilot pane dismissal");
        }
        catch (TimeoutException error)
        {
            throw new TimeoutException(
                $"{error.Message} {lastState} Check failure screenshots/UI diagnostics for remaining task panes " +
                "or add-ins hiding Office controls. No add-in was disabled or Office preference changed.", error);
        }
        Window.Focus();
    }

    private bool HasEditingDocument()
    {
        if (appName != "PowerPoint")
            return FindNativeSurfaceHandle(appName == "Word" ? "_WwG" : "EXCEL7") != IntPtr.Zero;
        return ReadModel((scope, app) =>
        {
            dynamic presentations = scope.Track(app.Presentations);
            if ((int)presentations.Count == 0)
                return false;
            dynamic presentation = scope.Track(app.ActivePresentation);
            dynamic slides = scope.Track(presentation.Slides);
            return (int)slides.Count > 0;
        });
    }

    private (string Name, AutomationElement Host)[] KnownTaskPanes()
    {
        var result = new List<(string, AutomationElement)>();
        foreach (var host in Window.FindAllDescendants(cf => cf.ByClassName("MsoWorkPane")))
        {
            if (host.Properties.ProcessId.ValueOrDefault != Application.ProcessId || !IsVisiblyOpen(host))
                continue;
            var marker = host.FindAllDescendants(cf => cf.ByClassName("NetUInetpane"))
                .FirstOrDefault(e => OfficeStartup.IsKnownTaskPane(Class(e), Name(e)) && IsVisiblyOpen(e));
            if (marker is not null)
                result.Add((Name(marker), host));
        }
        return result.ToArray();
    }

    private void DismissKnownTaskPanes()
    {
        foreach (var name in KnownTaskPanes().Select(p => p.Name).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            OfficeStartup.CloseOptionalPane(name,
                () => KnownTaskPanes().Any(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase)),
                () =>
                {
                    var panes = KnownTaskPanes().Where(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase)).ToArray();
                    if (panes.Length == 0)
                        return true;
                    if (panes.Length != 1)
                        throw new InvalidOperationException($"Multiple {name} task panes found; refusing an ambiguous Close action.");
                    return TryCloseTaskPane(name);
                }, Wait, LogStartup);
        }
    }

    private bool TryCloseTaskPane(string name)
    {
        // Office exposes a task pane's header separately from its content window.
        var headers = Window.FindAllDescendants(cf => cf.ByClassName("NetUINativeHWNDHost"))
            .Where(e => OfficeStartup.IsPaneHeader(Class(e), Name(e), name) &&
                e.Properties.ProcessId.ValueOrDefault == Application.ProcessId &&
                IsVisiblyOpen(e)).ToArray();
        if (headers.Length == 0)
            return false;
        if (headers.Length != 1)
            throw new InvalidOperationException($"Multiple {name} task pane headers found; refusing an ambiguous Close action.");
        var header = headers[0];
        var walker = Automation.TreeWalkerFactory.GetRawViewWalker();
        foreach (var button in header.FindAllDescendants(cf => cf.ByControlType(ControlType.Button)))
        {
            if (!IsVisiblyOpen(button) || !button.Properties.IsEnabled.ValueOrDefault ||
                !OfficeStartup.IsPaneCloseButton(Name(button)))
                continue;
            var parent = walker.GetParent(button);
            while (parent is not null && !parent.Equals(header) &&
                Class(parent) is not ("OsfAxControl" or "WebView2Holder" or "Chrome_WidgetWin_1"))
                parent = walker.GetParent(parent);
            if (parent is null || !parent.Equals(header) || !button.Patterns.Invoke.IsSupported)
                continue;
            button.Patterns.Invoke.Pattern.Invoke();
            return true;
        }
        return false;
    }

    private bool IsVisiblyOpen(AutomationElement element)
    {
        bool? offscreen = element.Properties.IsOffscreen.TryGetValue(out var value) ? value : null;
        if (offscreen == true)
            return false;
        if (!element.Properties.BoundingRectangle.TryGetValue(out var bounds))
            return false;
        bool? nativeVisible = null;
        if (element.Properties.NativeWindowHandle.TryGetValue(out var handle) && handle != 0)
        {
            var hwnd = new IntPtr(handle);
            var owner = new IntPtr(Window.Properties.NativeWindowHandle.Value);
            GetWindowThreadProcessId(hwnd, out var processId);
            var className = new StringBuilder(256);
            GetClassName(hwnd, className, className.Capacity);
            if (className.ToString() == Class(element))
            {
                nativeVisible = processId == Application.ProcessId && IsChild(owner, hwnd) &&
                    IsWindowVisible(hwnd) && !IsIconic(owner);
            }
        }
        return OfficeStartup.IsVisiblyOpen(offscreen, nativeVisible, bounds, Window.BoundingRectangle);
    }

    private void DismissCloudOnboarding()
    {
        var dialog = Elements().FirstOrDefault(e => Type(e) == ControlType.Window
            && Name(e) == "Cloud Create Onboarding");
        if (dialog is null)
            return;
        var close = dialog.FindFirstDescendant(cf => cf.ByAutomationId("Close"))?.AsButton()
            ?? throw new InvalidOperationException("Cloud onboarding appeared without its Close button.");
        LogStartup("Closing cloud-save onboarding without changing save preferences.");
        close.Invoke();
        Wait(() => !Elements().Any(e => Name(e) == "Cloud Create Onboarding"),
            "cloud onboarding to close");
    }

    public void Backstage(string page)
    {
        Window.Focus();
        Activate(Find(e => Id(e) == "FileTabButton", "File ribbon tab"));
        Click($"File -> {page}", page, page switch
        {
            "New" => "\u65b0\u5efa",
            "Open" => "\u6253\u5f00",
            "Save As" => "\u53e6\u5b58\u4e3a",
            "Options" => "\u9009\u9879",
            _ => throw new ArgumentOutOfRangeException(nameof(page))
        });
    }

    public void Enter(AutomationElement element, string text)
    {
        element.Focus();
        Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
        UnicodeInput.Type(text);
        Wait(() => ReadText(element) == text, $"exact text in '{element.Name}'");
    }

    public static string ReadText(AutomationElement element)
    {
        if (element.Patterns.Value.IsSupported)
            return element.Patterns.Value.Pattern.Value.Value;
        if (element.Patterns.Text.IsSupported)
            return element.Patterns.Text.Pattern.DocumentRange.GetText(-1).TrimEnd('\r', '\n');
        throw new InvalidOperationException($"'{Name(element)}' exposes neither Value nor Text; exact readback is unavailable.");
    }

    public AutomationElement TemplateSearch()
    {
        return Find(e => Type(e) == ControlType.Edit &&
            (Name(e).Contains("template", StringComparison.OrdinalIgnoreCase)
             || Name(e).Contains("\u6a21\u677f", StringComparison.Ordinal)), "template search field");
    }

    public void OpenFile(string path)
    {
        Click("Browse local files", "Browse", "\u6d4f\u89c8");
        var field = FileNameField();
        Enter(field, path);
        ClickDialogButton("Open", "\u6253\u5f00(O)", "\u6253\u5f00");
        Wait(() => Window.Title.Contains(Path.GetFileNameWithoutExtension(path), StringComparison.Ordinal),
            "opened file in title bar");
        PrepareDocument();
    }

    public AutomationElement FileNameField()
    {
        var dialog = Find(e => Type(e) == ControlType.Window && Class(e) == "#32770", "native file dialog");
        return dialog.FindAllDescendants().FirstOrDefault(e =>
            (Id(e) == "1001" || Id(e) == "1148") && Type(e) == ControlType.Edit)
            ?? throw new InvalidOperationException("Native file dialog has no supported filename edit field.");
    }

    public void ClickDialogButton(params string[] names)
    {
        var button = Find(e => Type(e) is ControlType.Button or ControlType.SplitButton && Id(e) == "1"
            && names.Any(n => Name(e).Replace("&", "").StartsWith(n, StringComparison.OrdinalIgnoreCase)),
            string.Join("/", names) + " dialog button");
        Activate(button);
    }

    public void BrowseToDesktop()
    {
        Click("Browse Save As", "Browse", "\u6d4f\u89c8");
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (string.IsNullOrWhiteSpace(desktop) || !Directory.Exists(desktop))
            throw new DirectoryNotFoundException("A usable Desktop directory is required by Word 5844.");
        Enter(FileNameField(), desktop + "\\");
        Keyboard.Type(VirtualKeyShort.ENTER);
        Wait(() => Elements().Any(e => Name(e).Contains("Desktop", StringComparison.OrdinalIgnoreCase)
            || Name(e).Contains("\u684c\u9762", StringComparison.Ordinal)), "Desktop in Save As dialog");
    }

    public void Capture(string path)
    {
        Thread.Sleep(250);
        var bounds = Window.BoundingRectangle;
        foreach (var window in Application.GetAllTopLevelWindows(Automation))
        {
            var r = window.BoundingRectangle;
            if (r.Width > 0 && r.Height > 0)
                bounds = Rectangle.Union(bounds, r);
        }
        CaptureBounds(bounds, path);
    }

    private static void CaptureBounds(Rectangle bounds, string path)
    {
        if (bounds.Width <= 0 || bounds.Height <= 0)
            throw new InvalidOperationException("Office window bounds are empty; cannot capture screenshot.");
        FlaUI.Core.Capturing.Capture.Rectangle(bounds).ToFile(path);
    }

    public void DumpUi(string path) =>
        File.WriteAllLines(path, Elements().Select(e =>
            $"{Type(e)} | {Id(e)} | {Class(e)} | {Name(e)}"));

    public T ReadModel<T>(Func<ComScope, dynamic, T> read)
    {
        using var retry = new ComMessageFilter(settings.UiTimeoutSeconds);
        using var scope = new ComScope();
        object model;
        if (appName == "PowerPoint")
            model = ComScope.ActiveObject("PowerPoint.Application");
        else
        {
            var handle = NativeSurfaceHandle(appName == "Word" ? "_WwG" : "EXCEL7");
            var id = new Guid("00020400-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(
                handle, 0xFFFFFFF0, ref id, out model));
        }
        dynamic native = scope.Track(model);
        dynamic app = appName == "PowerPoint" ? native : scope.Track(native.Application);
        return read(scope, app);
    }

    private IntPtr NativeSurfaceHandle(string className)
    {
        IntPtr found = IntPtr.Zero;
        Wait(() =>
        {
            found = FindNativeSurfaceHandle(className);
            return found != IntPtr.Zero;
        }, "native " + className + " object model surface");
        return found;
    }

    private IntPtr FindNativeSurfaceHandle(string className)
    {
        var found = IntPtr.Zero;
        EnumChildWindows(new IntPtr(Window.Properties.NativeWindowHandle.Value), (handle, _) =>
        {
            var name = new StringBuilder(256);
            GetClassName(handle, name, name.Capacity);
            if (name.ToString() != className)
                return true;
            found = handle;
            return false;
        }, IntPtr.Zero);
        return found;
    }

    public void Dispose()
    {
        try
        {
            if (application is not null && !application.HasExited)
            {
                using var process = Process.GetProcessById(application.ProcessId);
                process.CloseMainWindow();
                if (!process.WaitForExit(3000))
                {
                    NUnit.Framework.TestContext.Progress.WriteLine(
                        $"Cleanup: closing owned {appName} process {application.ProcessId} after timeout.");
                    application.Kill();
                    if (!process.WaitForExit(5000))
                        throw new TimeoutException($"Owned {appName} process did not exit after cleanup.");
                }
            }
        }
        finally
        {
            application?.Dispose();
            Automation.Dispose();
        }
    }

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid,
        [MarshalAs(UnmanagedType.Interface)] out object model);

    private delegate bool EnumWindowCallback(IntPtr handle, IntPtr parameter);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowCallback callback, IntPtr parameter);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr handle, StringBuilder className, int capacity);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr handle);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsChild(IntPtr parent, IntPtr child);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsIconic(IntPtr handle);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
