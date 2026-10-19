using System.ComponentModel;
using System.Runtime.InteropServices;

namespace OfficeGbAutomation.Common;

public sealed class DesktopDpiScope : IDisposable
{
    private readonly IntPtr previous;

    public DesktopDpiScope()
    {
        previous = SetThreadDpiAwarenessContext(new IntPtr(-4));
        if (previous == IntPtr.Zero)
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot enable per-monitor desktop coordinates.");
    }

    public void Dispose()
    {
        if (SetThreadDpiAwarenessContext(previous) == IntPtr.Zero)
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot restore desktop DPI awareness.");
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
}
