using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;

namespace OfficeGbAutomation.Common;

public static class OfficeStartup
{
    public static bool IsKnownTaskPane(string className, string name) =>
        className == "NetUInetpane" &&
        (name.Equals("OfficePLUS", StringComparison.OrdinalIgnoreCase) ||
         name.Equals("Copilot", StringComparison.OrdinalIgnoreCase));

    public static bool IsPaneHeader(string className, string name, string paneName) =>
        className == "NetUINativeHWNDHost" && IsKnownTaskPane("NetUInetpane", paneName) &&
        name.Equals(paneName, StringComparison.OrdinalIgnoreCase);

    public static bool IsPaneCloseButton(string name) =>
        name is "Close pane" or "\u5173\u95ed\u7a97\u683c";

    public static bool IsCopilotChatCloseButton(string paneName, string automationId) =>
        paneName.Equals("Copilot", StringComparison.OrdinalIgnoreCase) &&
        automationId == "sessionPaneCloseButton";

    public static bool IsVisiblyOpen(bool? isOffscreen, bool? nativeWindowVisible,
        Rectangle bounds, Rectangle officeBounds)
    {
        var visibleBounds = Rectangle.Intersect(bounds, officeBounds);
        return isOffscreen != true && nativeWindowVisible != false &&
            (isOffscreen == false || nativeWindowVisible == true) &&
            bounds.Width > 0 && bounds.Height > 0 &&
            visibleBounds.Width > 0 && visibleBounds.Height > 0;
    }

    public static void WaitFor(Func<bool> condition, string description, TimeSpan timeout,
        Action<string> log)
    {
        var timer = Stopwatch.StartNew();
        COMException? lastError = null;
        var reported = new HashSet<int>();
        while (timer.Elapsed < timeout)
        {
            try
            {
                if (condition())
                    return;
            }
            catch (COMException error) when (error.HResult is
                unchecked((int)0x80040201) or unchecked((int)0x80010001) or unchecked((int)0x8001010A))
            {
                lastError = error;
                if (reported.Add(error.HResult))
                    log($"Transient UI Automation failure while waiting for {description}: 0x{error.HResult:X8}. Retrying within the timeout.");
            }
            Thread.Sleep(150);
        }
        throw new TimeoutException($"Timed out waiting for {description}." +
            (lastError is null ? "" : $" Last UI Automation error: 0x{lastError.HResult:X8}."), lastError);
    }

    public static void CloseOptionalPane(string name, Func<bool> isOpen, Func<bool> tryClose,
        Action<Func<bool>, string> wait, Action<string> log)
    {
        if (!isOpen())
            return;
        log($"Closing optional {name} task pane in the test-owned Office window; no add-in settings are changed.");
        wait(() => !isOpen() || tryClose(),
            $"{name} task pane header's own Close pane button. " +
            "If unavailable, close the pane manually before running; do not disable add-ins automatically");
        wait(() => !isOpen(),
            $"{name} task pane to disappear after closing. It may be reopening or blocking Office accessibility");
    }
}
