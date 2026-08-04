using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WindowsInput;
using WindowsInput.Native;

namespace moshushou.Input
{
    public sealed class SendInputBackend : IInputBackend
    {
        private readonly object _sync = new();
        private readonly InputSimulator _simulator = new();
        private readonly HashSet<InputKey> _pressedKeys = new();
        private byte _pressedMouseButtons;
        private bool _disposed;

        public string Name => "SendInput";
        public bool IsAvailable => !_disposed;

        public void KeyDown(InputKey key)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                _simulator.Keyboard.KeyDown(ToVirtualKeyCode(key));
                _pressedKeys.Add(key);
            }
        }

        public void KeyUp(InputKey key)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                _simulator.Keyboard.KeyUp(ToVirtualKeyCode(key));
                _pressedKeys.Remove(key);
            }
        }

        public void KeyPress(InputKey key)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                _simulator.Keyboard.KeyPress(ToVirtualKeyCode(key));
            }
        }

        public void KeyChord(InputKey modifier, InputKey key)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                _simulator.Keyboard.ModifiedKeyStroke(
                    ToVirtualKeyCode(modifier),
                    ToVirtualKeyCode(key));
            }
        }

        public void MoveMouseAbsolute(int screenX, int screenY)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                if (!SetCursorPos(screenX, screenY))
                {
                    throw new InvalidOperationException($"SetCursorPos failed: {Marshal.GetLastWin32Error()}");
                }
            }
        }

        public void MouseButton(InputMouseButton button, bool pressed)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                uint flag = ToMouseEventFlag(button, pressed);
                mouse_event(flag, 0, 0, 0, UIntPtr.Zero);

                if (pressed)
                {
                    _pressedMouseButtons |= (byte)button;
                }
                else
                {
                    _pressedMouseButtons &= (byte)~(byte)button;
                }
            }
        }

        public void ReleaseAll()
        {
            lock (_sync)
            {
                if (_disposed)
                {
                    return;
                }

                foreach (InputKey key in _pressedKeys)
                {
                    try
                    {
                        _simulator.Keyboard.KeyUp(ToVirtualKeyCode(key));
                    }
                    catch
                    {
                    }
                }

                _pressedKeys.Clear();

                foreach (InputMouseButton button in Enum.GetValues<InputMouseButton>())
                {
                    if ((_pressedMouseButtons & (byte)button) != 0)
                    {
                        try
                        {
                            mouse_event(ToMouseEventFlag(button, false), 0, 0, 0, UIntPtr.Zero);
                        }
                        catch
                        {
                        }
                    }
                }

                _pressedMouseButtons = 0;
            }
        }

        public void Dispose()
        {
            ReleaseAll();
            _disposed = true;
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        private static VirtualKeyCode ToVirtualKeyCode(InputKey key)
        {
            return key switch
            {
                InputKey.A => VirtualKeyCode.VK_A,
                InputKey.F => VirtualKeyCode.VK_F,
                InputKey.S => VirtualKeyCode.VK_S,
                InputKey.V => VirtualKeyCode.VK_V,
                InputKey.Enter => VirtualKeyCode.RETURN,
                InputKey.Escape => VirtualKeyCode.ESCAPE,
                InputKey.Backspace => VirtualKeyCode.BACK,
                InputKey.Delete => VirtualKeyCode.DELETE,
                InputKey.DownArrow => VirtualKeyCode.DOWN,
                InputKey.LeftControl => VirtualKeyCode.LCONTROL,
                InputKey.RightControl => VirtualKeyCode.RCONTROL,
                InputKey.LeftShift => VirtualKeyCode.LSHIFT,
                InputKey.RightShift => VirtualKeyCode.RSHIFT,
                InputKey.LeftAlt => VirtualKeyCode.LMENU,
                InputKey.RightAlt => VirtualKeyCode.RMENU,
                InputKey.LeftWindows => VirtualKeyCode.LWIN,
                InputKey.RightWindows => VirtualKeyCode.RWIN,
                _ => throw new ArgumentOutOfRangeException(nameof(key), key, null)
            };
        }

        private static uint ToMouseEventFlag(InputMouseButton button, bool pressed)
        {
            return (button, pressed) switch
            {
                (InputMouseButton.Left, true) => MOUSEEVENTF_LEFTDOWN,
                (InputMouseButton.Left, false) => MOUSEEVENTF_LEFTUP,
                (InputMouseButton.Right, true) => MOUSEEVENTF_RIGHTDOWN,
                (InputMouseButton.Right, false) => MOUSEEVENTF_RIGHTUP,
                (InputMouseButton.Middle, true) => MOUSEEVENTF_MIDDLEDOWN,
                (InputMouseButton.Middle, false) => MOUSEEVENTF_MIDDLEUP,
                _ => throw new NotSupportedException($"SendInput backend does not support mouse button {button}.")
            };
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
    }
}
