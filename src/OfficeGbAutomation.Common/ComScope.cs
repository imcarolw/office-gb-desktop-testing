using System.Runtime.InteropServices;

namespace OfficeGbAutomation.Common;

public sealed class ComScope : IDisposable
{
    private readonly List<object> objects = [];

    public dynamic Track(object value)
    {
        if (Marshal.IsComObject(value) && !objects.Any(o => ReferenceEquals(o, value)))
            objects.Add(value);
        return value;
    }

    public static object ActiveObject(string progId)
    {
        Marshal.ThrowExceptionForHR(CLSIDFromProgID(progId, out var classId));
        Marshal.ThrowExceptionForHR(GetActiveObject(ref classId, IntPtr.Zero, out var value));
        return value;
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid classId);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid classId, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object value);

    public void Dispose()
    {
        for (var i = objects.Count - 1; i >= 0; i--)
            Marshal.FinalReleaseComObject(objects[i]);
        objects.Clear();
    }

    public static void CreateExcelFixture(string path)
    {
        using var scope = new ComScope();
        var type = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel COM registration was not found.");
        dynamic app = scope.Track(Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Could not create Excel fixture instance."));
        try
        {
            app.Visible = false;
            app.DisplayAlerts = false;
            dynamic workbooks = scope.Track(app.Workbooks);
            dynamic workbook = scope.Track(workbooks.Add());
            workbook.SaveAs(path, 51);
            workbook.Close(false);
        }
        finally
        {
            app.Quit();
        }
    }
}
