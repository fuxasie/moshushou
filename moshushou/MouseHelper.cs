using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using moshushou.Input;

namespace moshushou
{
    /// <summary>
    /// Human-like mouse movement and click helpers.
    /// </summary>
    public static class MouseHelper
    {
        private static IInputBackend? _inputBackend;

        public static void Configure(IInputBackend inputBackend)
        {
            _inputBackend = inputBackend ?? throw new ArgumentNullException(nameof(inputBackend));
        }

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        /// <summary>
        /// Check if the ESC key is currently pressed (Global check).
        /// </summary>
        public static bool IsEscPressed()
        {
            // VK_ESCAPE = 0x1B
            // GetAsyncKeyState returns short. High bit set means key is down.
            return (GetAsyncKeyState(0x1B) & 0x8000) != 0;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        private static readonly Random _random = new Random();

        /// <summary>
        /// Get current cursor position in screen coordinates.
        /// </summary>
        public static Point GetCursorPosition()
        {
            if (GetCursorPos(out POINT point))
            {
                return new Point(point.X, point.Y);
            }

            return Point.Empty;
        }

        /// <summary>
        /// Move to target with smooth trajectory and click.
        /// </summary>
        public static async Task HumanLikeClickAsync(int x, int y, int moveDurationBase = 90)
        {
            int targetX = x + _random.Next(-2, 3);
            int targetY = y + _random.Next(-2, 3);

            await MoveMouseSmoothlyAsync(targetX, targetY, moveDurationBase);
            await Task.Delay(_random.Next(4, 10));

            InputBackend.MouseButton(InputMouseButton.Left, true);
            await Task.Delay(_random.Next(10, 20));
            InputBackend.MouseButton(InputMouseButton.Left, false);
        }

        /// <summary>
        /// Click at current cursor position without moving.
        /// </summary>
        public static async Task LeftClickCurrentAsync()
        {
            if (!GetCursorPos(out POINT point))
            {
                return;
            }

            InputBackend.MouseButton(InputMouseButton.Left, true);
            await Task.Delay(_random.Next(10, 20));
            InputBackend.MouseButton(InputMouseButton.Left, false);
        }

        /// <summary>
        /// Move cursor smoothly with a cubic Bezier path.
        /// </summary>
        public static async Task MoveMouseSmoothlyAsync(int targetX, int targetY, int durationMs)
        {
            GetCursorPos(out POINT startPoint);
            int startX = startPoint.X;
            int startY = startPoint.Y;

            double distance = Math.Sqrt(Math.Pow(targetX - startX, 2) + Math.Pow(targetY - startY, 2));
            if (distance < 4)
            {
                InputBackend.MoveMouseAbsolute(targetX, targetY);
                return;
            }

            int actualDuration = durationMs + (int)(distance * 0.025);
            if (actualDuration > 180) actualDuration = 180;
            if (actualDuration < 20) actualDuration = 20;

            int randomOffset = Math.Max(6, Math.Min((int)(distance * 0.18), 90));

            int p1x = startX + (targetX - startX) / 3 + _random.Next(-randomOffset, randomOffset);
            int p1y = startY + (targetY - startY) / 3 + _random.Next(-randomOffset, randomOffset);

            int p2x = startX + 2 * (targetX - startX) / 3 + _random.Next(-randomOffset, randomOffset);
            int p2y = startY + 2 * (targetY - startY) / 3 + _random.Next(-randomOffset, randomOffset);

            int steps = actualDuration / 12;
            if (steps < 4) steps = 4;

            for (int i = 0; i <= steps; i++)
            {
                double t = (double)i / steps;
                t = t * t * (3f - 2f * t);

                double u = 1 - t;
                double tt = t * t;
                double uu = u * u;
                double uuu = uu * u;
                double ttt = tt * t;

                double x = uuu * startX + 3 * uu * t * p1x + 3 * u * tt * p2x + ttt * targetX;
                double y = uuu * startY + 3 * uu * t * p1y + 3 * u * tt * p2y + ttt * targetY;

                InputBackend.MoveMouseAbsolute((int)x, (int)y);
                if (i < steps)
                {
                    await Task.Delay(_random.Next(1, 3));
                }
            }

            InputBackend.MoveMouseAbsolute(targetX, targetY);
        }

        private static IInputBackend InputBackend =>
            _inputBackend ?? throw new InvalidOperationException("MouseHelper has not been configured with an input backend.");
    }
}
