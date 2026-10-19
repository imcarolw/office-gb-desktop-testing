using System.Runtime.InteropServices;

namespace OfficeGbAutomation.Common;

[ComVisible(true)]
public sealed class ComMessageFilter : IOleMessageFilter, IDisposable
{
    private readonly int timeoutMilliseconds;
    private readonly IOleMessageFilter? previous;

    public ComMessageFilter(int timeoutSeconds)
    {
        timeoutMilliseconds = checked(timeoutSeconds * 1000);
        Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(this, out previous));
    }

    public int HandleInComingCall(int callType, IntPtr caller, int ticks, IntPtr info) => 0;

    // Office may reject readback while it processes input. OLE retries only those
    // transient rejections, with the same bounded timeout as the UI waits.
    public int RetryRejectedCall(IntPtr callee, int ticks, int rejectType) =>
        rejectType is 1 or 2 && ticks < timeoutMilliseconds ? 100 : -1;

    public int MessagePending(IntPtr callee, int ticks, int pendingType) => 2;

    public void Dispose() => Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(previous, out _));

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        [MarshalAs(UnmanagedType.Interface)] IOleMessageFilter? filter,
        [MarshalAs(UnmanagedType.Interface)] out IOleMessageFilter? previous);
}

[ComVisible(true), Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int callType, IntPtr caller, int ticks, IntPtr info);
    [PreserveSig] int RetryRejectedCall(IntPtr callee, int ticks, int rejectType);
    [PreserveSig] int MessagePending(IntPtr callee, int ticks, int pendingType);
}
