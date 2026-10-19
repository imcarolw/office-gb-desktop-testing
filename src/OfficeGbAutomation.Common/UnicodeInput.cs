using System.ComponentModel;
using System.Runtime.InteropServices;

namespace OfficeGbAutomation.Common;

public static class UnicodeInput
{
    public static void Validate(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        for (var i = 0; i < text.Length; i++)
        {
            if (char.IsHighSurrogate(text[i]))
            {
                if (++i >= text.Length || !char.IsLowSurrogate(text[i]))
                    throw new ArgumentException("Text contains an unpaired high surrogate.", nameof(text));
            }
            else if (char.IsLowSurrogate(text[i]))
                throw new ArgumentException("Text contains an unpaired low surrogate.", nameof(text));
        }
    }

    public static void Type(string text)
    {
        Validate(text);
        foreach (var character in text)
        {
            var inputs = new[] { Create(character, false), Create(character, true) };
            if (SendInput(2, inputs, Marshal.SizeOf<Input>()) != 2)
                throw new Win32Exception(Marshal.GetLastWin32Error(),
                    $"Windows could not send Unicode code unit U+{(int)character:X4}.");
        }
    }

    private static Input Create(char character, bool keyUp) => new()
    {
        Type = 1,
        Data = new InputUnion
        {
            Keyboard = new KeyboardInput
            {
                ScanCode = character,
                Flags = 0x0004u | (keyUp ? 0x0002u : 0u)
            }
        }
    };

    [StructLayout(LayoutKind.Sequential)]
    private struct Input { public uint Type; public InputUnion Data; }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)] public KeyboardInput Keyboard;
        [FieldOffset(0)] public MouseInput Mouse;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardInput
    {
        public ushort VirtualKey;
        public ushort ScanCode;
        public uint Flags;
        public uint Time;
        public UIntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput
    {
        public int X;
        public int Y;
        public uint MouseData;
        public uint Flags;
        public uint Time;
        public UIntPtr ExtraInfo;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint count, Input[] inputs, int size);
}
