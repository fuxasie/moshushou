using System;
using System.Collections.Concurrent;

namespace moshushou.Input
{
    /// <summary>
    /// Virtual HID input does not carry LLKHF_INJECTED. Track our own key
    /// transitions so the low-level hook does not mistake them for physical Ctrl.
    /// </summary>
    public static class SyntheticInputTracker
    {
        private sealed record ExpectedTransition(uint VirtualKey, bool IsDown, long ExpiresAt);

        private static readonly ConcurrentQueue<ExpectedTransition> Expected = new();
        private static readonly long ExpirationTicks = TimeSpan.FromSeconds(1).Ticks;

        public static void Expect(InputKey key, bool isDown)
        {
            uint virtualKey = ToVirtualKey(key);
            if (virtualKey == 0)
            {
                return;
            }

            long now = DateTime.UtcNow.Ticks;
            TrimExpired(now);
            Expected.Enqueue(new ExpectedTransition(virtualKey, isDown, now + ExpirationTicks));
        }

        public static bool TryConsume(uint virtualKey, bool isDown)
        {
            long now = DateTime.UtcNow.Ticks;
            TrimExpired(now);

            if (!Expected.TryPeek(out ExpectedTransition? expected))
            {
                return false;
            }

            if (expected.IsDown == isDown && IsSameKey(expected.VirtualKey, virtualKey))
            {
                Expected.TryDequeue(out _);
                return true;
            }

            return false;
        }

        public static void Clear()
        {
            while (Expected.TryDequeue(out _))
            {
            }
        }

        private static void TrimExpired(long now)
        {
            while (Expected.TryPeek(out ExpectedTransition? item) && item.ExpiresAt < now)
            {
                Expected.TryDequeue(out _);
            }
        }

        private static bool IsSameKey(uint expected, uint actual)
        {
            if (expected == actual)
            {
                return true;
            }

            return IsControl(expected) && IsControl(actual);
        }

        private static bool IsControl(uint key) => key is 0x11 or 0xA2 or 0xA3;

        private static uint ToVirtualKey(InputKey key)
        {
            return key switch
            {
                InputKey.A => 0x41,
                InputKey.F => 0x46,
                InputKey.S => 0x53,
                InputKey.V => 0x56,
                InputKey.Enter => 0x0D,
                InputKey.Escape => 0x1B,
                InputKey.Backspace => 0x08,
                InputKey.Delete => 0x2E,
                InputKey.DownArrow => 0x28,
                InputKey.LeftControl => 0xA2,
                InputKey.RightControl => 0xA3,
                InputKey.LeftShift => 0xA0,
                InputKey.RightShift => 0xA1,
                InputKey.LeftAlt => 0xA4,
                InputKey.RightAlt => 0xA5,
                InputKey.LeftWindows => 0x5B,
                InputKey.RightWindows => 0x5C,
                _ => 0
            };
        }
    }
}
