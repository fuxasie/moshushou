using System.Text.Json;
using System.Runtime.InteropServices;
using moshushou.Input;

string? resultPath = args.Length >= 2 && args[0] == "--result" ? args[1] : null;
bool mouseTest = args.Contains("--mouse-test", StringComparer.OrdinalIgnoreCase);
bool clickTest = args.Contains("--click-test", StringComparer.OrdinalIgnoreCase);

var result = new SmokeTestResult
{
    Timestamp = DateTimeOffset.Now,
    WindowsVersion = Environment.OSVersion.Version.ToString()
};

try
{
    if (mouseTest || clickTest)
    {
        using var backend = new VirtualHidBackend();
        if (!backend.IsAvailable)
        {
            throw new InvalidOperationException("No compatible Virtual HID control collection was found.");
        }

        if (!NativeMethods.GetCursorPos(out NativeMethods.POINT before))
        {
            throw new InvalidOperationException($"GetCursorPos failed: {Marshal.GetLastWin32Error()}");
        }

        bool buttonPressed = false;
        try
        {
            int targetX;
            int targetY;
            if (clickTest)
            {
                IntPtr consoleWindow = NativeMethods.GetConsoleWindow();
                if (consoleWindow == IntPtr.Zero ||
                    !NativeMethods.GetWindowRect(consoleWindow, out NativeMethods.RECT consoleRect))
                {
                    throw new InvalidOperationException("The smoke-test console window is unavailable.");
                }

                NativeMethods.SetForegroundWindow(consoleWindow);
                targetX = consoleRect.Left + (consoleRect.Right - consoleRect.Left) / 2;
                targetY = consoleRect.Top + (consoleRect.Bottom - consoleRect.Top) / 2;
            }
            else
            {
                int left = NativeMethods.GetSystemMetrics(76);
                int top = NativeMethods.GetSystemMetrics(77);
                int width = Math.Max(1, NativeMethods.GetSystemMetrics(78));
                int height = Math.Max(1, NativeMethods.GetSystemMetrics(79));
                targetX = left + width / 3;
                targetY = top + height / 3;
            }

            backend.MoveMouseAbsolute(targetX, targetY);
            Thread.Sleep(500);
            NativeMethods.GetCursorPos(out NativeMethods.POINT after);

            result.CursorBefore = $"{before.X},{before.Y}";
            result.CursorTarget = $"{targetX},{targetY}";
            result.CursorAfter = $"{after.X},{after.Y}";
            bool moved = Math.Abs(after.X - targetX) <= 2 && Math.Abs(after.Y - targetY) <= 2;
            if (clickTest && moved)
            {
                backend.MouseButton(InputMouseButton.Left, true);
                buttonPressed = true;
                Thread.Sleep(150);
                result.ButtonDownObserved = (NativeMethods.GetAsyncKeyState(0x01) & 0x8000) != 0;

                backend.MouseButton(InputMouseButton.Left, false);
                buttonPressed = false;
                Thread.Sleep(150);
                result.ButtonReleasedObserved = (NativeMethods.GetAsyncKeyState(0x01) & 0x8000) == 0;
                result.Success = result.ButtonDownObserved && result.ButtonReleasedObserved;
                result.Message = result.Success
                    ? "Virtual HID left-button down and up states were observed successfully."
                    : "The cursor moved, but the expected left-button state transition was not observed.";
            }
            else
            {
                result.Success = moved;
                result.Message = result.Success
                    ? "Virtual HID mouse report moved the cursor to the requested absolute position."
                    : "Virtual HID mouse report was accepted, but the cursor did not reach the requested position.";
            }
        }
        finally
        {
            if (buttonPressed)
            {
                try
                {
                    backend.MouseButton(InputMouseButton.Left, false);
                }
                catch
                {
                    // Continue with cursor restoration even if the device disconnected.
                }
            }
            NativeMethods.SetCursorPos(before.X, before.Y);
        }
    }
    else
    {
        using VirtualHidControlDevice? device = VirtualHidControlDevice.TryOpen();
        if (device is null)
        {
            throw new InvalidOperationException("No compatible Virtual HID control collection was found.");
        }

        result.DevicePath = device.DevicePath;
        result.ProductName = device.ProductName;

        // Neutral keyboard report: validates the complete feeder -> control
        // collection -> UMDF driver path without producing a visible key press.
        device.Send(new byte[] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 });
        result.Success = true;
        result.Message = "Virtual HID control channel opened and accepted a neutral keyboard report.";
    }
}
catch (Exception ex)
{
    result.Success = false;
    result.Message = ex.ToString();
}

string json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
Console.WriteLine(json);
if (!string.IsNullOrWhiteSpace(resultPath))
{
    string fullPath = Path.GetFullPath(resultPath);
    Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
    File.WriteAllText(fullPath, json);
}

return result.Success ? 0 : 1;

internal sealed class SmokeTestResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public string WindowsVersion { get; set; } = string.Empty;
    public DateTimeOffset Timestamp { get; set; }
    public string? ProductName { get; set; }
    public string? DevicePath { get; set; }
    public string? CursorBefore { get; set; }
    public string? CursorTarget { get; set; }
    public string? CursorAfter { get; set; }
    public bool ButtonDownObserved { get; set; }
    public bool ButtonReleasedObserved { get; set; }
}

internal static class NativeMethods
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool GetCursorPos(out POINT point);

    [DllImport("user32.dll")]
    internal static extern int GetSystemMetrics(int index);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetCursorPos(int x, int y);

    [DllImport("kernel32.dll")]
    internal static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool GetWindowRect(IntPtr window, out RECT rect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetForegroundWindow(IntPtr window);

    [DllImport("user32.dll")]
    internal static extern short GetAsyncKeyState(int virtualKey);
}
