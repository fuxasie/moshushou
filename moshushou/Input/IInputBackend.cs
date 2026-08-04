using System;

namespace moshushou.Input
{
    public enum InputKey : byte
    {
        A = 0x04,
        F = 0x09,
        S = 0x16,
        V = 0x19,
        Enter = 0x28,
        Escape = 0x29,
        Backspace = 0x2A,
        Delete = 0x4C,
        DownArrow = 0x51,

        LeftControl = 0xE0,
        LeftShift = 0xE1,
        LeftAlt = 0xE2,
        LeftWindows = 0xE3,
        RightControl = 0xE4,
        RightShift = 0xE5,
        RightAlt = 0xE6,
        RightWindows = 0xE7
    }

    public enum InputMouseButton : byte
    {
        Left = 0x01,
        Right = 0x02,
        Middle = 0x04,
        X1 = 0x08,
        X2 = 0x10
    }

    public interface IInputBackend : IDisposable
    {
        string Name { get; }
        bool IsAvailable { get; }

        void KeyDown(InputKey key);
        void KeyUp(InputKey key);
        void KeyPress(InputKey key);
        void KeyChord(InputKey modifier, InputKey key);

        void MoveMouseAbsolute(int screenX, int screenY);
        void MouseButton(InputMouseButton button, bool pressed);

        void ReleaseAll();
    }

    public sealed class InputBackendUnavailableException : InvalidOperationException
    {
        public InputBackendUnavailableException(string message)
            : base(message)
        {
        }

        public InputBackendUnavailableException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
