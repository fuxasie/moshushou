using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace moshushou.Input
{
    /// <summary>
    /// Sends keyboard and absolute mouse reports through the FakerInput-compatible
    /// vendor control collection (UsagePage FF00, Usage 01, report id 40).
    /// The matching UMDF2 driver can be replaced without changing application code.
    /// </summary>
    public sealed class VirtualHidBackend : IInputBackend
    {
        private readonly object _sync = new();
        private readonly HashSet<InputKey> _pressedKeys = new();
        private VirtualHidControlDevice? _device;
        private byte _modifiers;
        private byte _mouseButtons;
        private bool _disposed;

        public string Name => "VirtualHID";

        public static bool IsCompatibleDevicePresent()
        {
            using VirtualHidControlDevice? device = VirtualHidControlDevice.TryOpen();
            return device != null;
        }

        public bool IsAvailable
        {
            get
            {
                lock (_sync)
                {
                    return !_disposed && EnsureConnected(throwOnFailure: false);
                }
            }
        }

        public void KeyDown(InputKey key)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                KeyDownCore(key);
            }
        }

        public void KeyUp(InputKey key)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                KeyUpCore(key);
            }
        }

        public void KeyPress(InputKey key)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                bool alreadyPressed = IsPressed(key);
                if (!alreadyPressed)
                {
                    KeyDownCore(key);
                }

                try
                {
                    if (alreadyPressed)
                    {
                        SendKeyboardReport();
                    }
                }
                finally
                {
                    if (!alreadyPressed)
                    {
                        KeyUpCore(key);
                    }
                }
            }
        }

        public void KeyChord(InputKey modifier, InputKey key)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                bool modifierAlreadyPressed = IsPressed(modifier);
                bool keyAlreadyPressed = IsPressed(key);

                if (!modifierAlreadyPressed)
                {
                    KeyDownCore(modifier);
                }

                try
                {
                    if (!keyAlreadyPressed)
                    {
                        KeyDownCore(key);
                    }
                }
                finally
                {
                    if (!keyAlreadyPressed)
                    {
                        KeyUpCore(key);
                    }

                    if (!modifierAlreadyPressed)
                    {
                        KeyUpCore(modifier);
                    }
                }
            }
        }

        public void MoveMouseAbsolute(int screenX, int screenY)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                SendAbsoluteMouseReport(screenX, screenY);
            }
        }

        public void MouseButton(InputMouseButton button, bool pressed)
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                if (pressed)
                {
                    _mouseButtons |= (byte)button;
                }
                else
                {
                    _mouseButtons &= (byte)~(byte)button;
                }

                if (!GetCursorPos(out POINT point))
                {
                    throw new InvalidOperationException($"GetCursorPos failed: {Marshal.GetLastWin32Error()}");
                }

                SendAbsoluteMouseReport(point.X, point.Y);
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

                try
                {
                    foreach (InputKey key in _pressedKeys.ToArray())
                    {
                        SyntheticInputTracker.Expect(key, false);
                    }

                    _pressedKeys.Clear();
                    _modifiers = 0;
                    SendKeyboardReport();

                    _mouseButtons = 0;
                    if (GetCursorPos(out POINT point))
                    {
                        SendAbsoluteMouseReport(point.X, point.Y);
                    }
                }
                catch
                {
                    SyntheticInputTracker.Clear();
                }
            }
        }

        internal void DisconnectForDriverMaintenance()
        {
            lock (_sync)
            {
                if (_disposed)
                {
                    return;
                }

                ReleaseAll();
                _device?.Dispose();
                _device = null;
                SyntheticInputTracker.Clear();
            }
        }

        public void Dispose()
        {
            lock (_sync)
            {
                if (_disposed)
                {
                    return;
                }

                ReleaseAll();
                _disposed = true;
                _device?.Dispose();
                _device = null;
                SyntheticInputTracker.Clear();
            }
        }

        private void KeyDownCore(InputKey key)
        {
            byte previousModifiers = _modifiers;
            if (IsModifier(key))
            {
                byte bit = ModifierBit(key);
                if ((_modifiers & bit) != 0)
                {
                    return;
                }

                _modifiers |= bit;
            }
            else
            {
                if (_pressedKeys.Contains(key))
                {
                    return;
                }

                int nonModifierCount = _pressedKeys.Count(item => !IsModifier(item));
                if (nonModifierCount >= 6)
                {
                    throw new InvalidOperationException("Virtual HID keyboard supports at most six simultaneous non-modifier keys.");
                }
            }

            _pressedKeys.Add(key);
            SyntheticInputTracker.Expect(key, true);
            try
            {
                SendKeyboardReport();
            }
            catch
            {
                _pressedKeys.Remove(key);
                _modifiers = previousModifiers;
                SyntheticInputTracker.Clear();
                throw;
            }
        }

        private void KeyUpCore(InputKey key)
        {
            bool wasPressed = _pressedKeys.Remove(key);
            bool modifierStateChanged = false;
            if (IsModifier(key))
            {
                byte bit = ModifierBit(key);
                modifierStateChanged = (_modifiers & bit) != 0;
                _modifiers &= (byte)~bit;
            }

            if (!wasPressed && !modifierStateChanged)
            {
                // Still send an all-up-compatible report for defensive releases,
                // but do not register an expected hook transition: Windows will
                // not emit KeyUp when this virtual device already had the key up.
                try
                {
                    SendKeyboardReport();
                }
                catch
                {
                    SyntheticInputTracker.Clear();
                    throw;
                }

                return;
            }

            SyntheticInputTracker.Expect(key, false);
            try
            {
                SendKeyboardReport();
            }
            catch
            {
                SyntheticInputTracker.Clear();
                throw;
            }
        }

        private void SendKeyboardReport()
        {
            byte[] report = new byte[9];
            report[0] = 0x01;
            report[1] = _modifiers;

            int reportIndex = 3;
            foreach (InputKey key in _pressedKeys.Where(item => !IsModifier(item)).OrderBy(item => (byte)item))
            {
                report[reportIndex++] = (byte)key;
            }

            SendPayload(report);
        }

        private void SendAbsoluteMouseReport(int screenX, int screenY)
        {
            int virtualLeft = GetSystemMetrics(SM_XVIRTUALSCREEN);
            int virtualTop = GetSystemMetrics(SM_YVIRTUALSCREEN);
            int virtualWidth = Math.Max(1, GetSystemMetrics(SM_CXVIRTUALSCREEN));
            int virtualHeight = Math.Max(1, GetSystemMetrics(SM_CYVIRTUALSCREEN));

            ushort normalizedX = NormalizeCoordinate(screenX, virtualLeft, virtualWidth);
            ushort normalizedY = NormalizeCoordinate(screenY, virtualTop, virtualHeight);

            byte[] report =
            {
                0x04,
                _mouseButtons,
                (byte)(normalizedX & 0xFF),
                (byte)(normalizedX >> 8),
                (byte)(normalizedY & 0xFF),
                (byte)(normalizedY >> 8),
                0x00
            };

            SendPayload(report);
        }

        private void SendPayload(byte[] payload)
        {
            if (payload.Length > VirtualHidControlDevice.MaximumPayloadLength)
            {
                throw new ArgumentOutOfRangeException(nameof(payload));
            }

            EnsureConnected(throwOnFailure: true);
            try
            {
                _device!.Send(payload);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ObjectDisposedException)
            {
                _device?.Dispose();
                _device = null;
                throw new InputBackendUnavailableException("Virtual HID device disconnected while sending an input report.", ex);
            }
        }

        private bool EnsureConnected(bool throwOnFailure)
        {
            if (_device != null && _device.IsOpen)
            {
                return true;
            }

            _device?.Dispose();
            _device = VirtualHidControlDevice.TryOpen();
            if (_device != null)
            {
                return true;
            }

            if (throwOnFailure)
            {
                throw new InputBackendUnavailableException(
                    "No compatible Virtual HID control collection was found. Install the UMDF2 driver first.");
            }

            return false;
        }

        private bool IsPressed(InputKey key) => _pressedKeys.Contains(key);

        private static bool IsModifier(InputKey key) => (byte)key is >= 0xE0 and <= 0xE7;

        private static byte ModifierBit(InputKey key) => (byte)(1 << ((byte)key - 0xE0));

        private static ushort NormalizeCoordinate(int coordinate, int origin, int length)
        {
            if (length <= 1)
            {
                return 0;
            }

            double value = (coordinate - origin) * 32767.0 / (length - 1);
            return (ushort)Math.Clamp((int)Math.Round(value), 0, 32767);
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        private const int SM_XVIRTUALSCREEN = 76;
        private const int SM_YVIRTUALSCREEN = 77;
        private const int SM_CXVIRTUALSCREEN = 78;
        private const int SM_CYVIRTUALSCREEN = 79;
    }

    internal sealed class VirtualHidControlDevice : IDisposable
    {
        public const int MaximumPayloadLength = 63;
        private const int ControlReportLength = 65;
        private const byte ControlReportId = 0x40;
        private const ushort ControlUsagePage = 0xFF00;
        private const ushort ControlUsage = 0x0001;

        private readonly SafeFileHandle _handle;
        private readonly object _writeSync = new();
        private bool _disposed;

        private VirtualHidControlDevice(SafeFileHandle handle, string path, string productName)
        {
            _handle = handle;
            DevicePath = path;
            ProductName = productName;
        }

        public string DevicePath { get; }
        public string ProductName { get; }
        public bool IsOpen => !_disposed && !_handle.IsInvalid && !_handle.IsClosed;

        public static VirtualHidControlDevice? TryOpen()
        {
            HidD_GetHidGuid(out Guid hidGuid);
            IntPtr deviceInfoSet = SetupDiGetClassDevs(
                ref hidGuid,
                IntPtr.Zero,
                IntPtr.Zero,
                DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);

            if (deviceInfoSet == INVALID_HANDLE_VALUE)
            {
                return null;
            }

            try
            {
                for (uint index = 0; ; index++)
                {
                    var interfaceData = new SP_DEVICE_INTERFACE_DATA
                    {
                        cbSize = Marshal.SizeOf<SP_DEVICE_INTERFACE_DATA>()
                    };

                    if (!SetupDiEnumDeviceInterfaces(
                            deviceInfoSet,
                            IntPtr.Zero,
                            ref hidGuid,
                            index,
                            ref interfaceData))
                    {
                        if (Marshal.GetLastWin32Error() == ERROR_NO_MORE_ITEMS)
                        {
                            break;
                        }

                        continue;
                    }

                    string? path = GetDevicePath(deviceInfoSet, ref interfaceData);
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        continue;
                    }

                    SafeFileHandle handle = CreateFile(
                        path,
                        GENERIC_READ | GENERIC_WRITE,
                        FILE_SHARE_READ | FILE_SHARE_WRITE,
                        IntPtr.Zero,
                        OPEN_EXISTING,
                        0,
                        IntPtr.Zero);

                    if (handle.IsInvalid)
                    {
                        handle.Dispose();
                        continue;
                    }

                    if (!TryGetCapabilities(handle, out HIDP_CAPS caps) ||
                        caps.UsagePage != ControlUsagePage ||
                        caps.Usage != ControlUsage ||
                        caps.OutputReportByteLength != ControlReportLength)
                    {
                        handle.Dispose();
                        continue;
                    }

                    string productName = GetProductName(handle);
                    if (!IsSupportedDeviceIdentity(path, productName))
                    {
                        handle.Dispose();
                        continue;
                    }

                    return new VirtualHidControlDevice(handle, path, productName);
                }
            }
            finally
            {
                SetupDiDestroyDeviceInfoList(deviceInfoSet);
            }

            return null;
        }

        public void Send(ReadOnlySpan<byte> payload)
        {
            if (payload.Length > MaximumPayloadLength)
            {
                throw new ArgumentOutOfRangeException(nameof(payload));
            }

            lock (_writeSync)
            {
                ObjectDisposedException.ThrowIf(_disposed, this);
                byte[] controlReport = new byte[ControlReportLength];
                controlReport[0] = ControlReportId;
                controlReport[1] = (byte)payload.Length;
                payload.CopyTo(controlReport.AsSpan(2));
                if (!WriteFile(
                        _handle,
                        controlReport,
                        controlReport.Length,
                        out int bytesWritten,
                        IntPtr.Zero) ||
                    bytesWritten != controlReport.Length)
                {
                    throw new IOException($"Virtual HID WriteFile failed: {Marshal.GetLastWin32Error()}");
                }
            }
        }

        public void Dispose()
        {
            lock (_writeSync)
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _handle.Dispose();
            }
        }

        private static string? GetDevicePath(IntPtr deviceInfoSet, ref SP_DEVICE_INTERFACE_DATA interfaceData)
        {
            SetupDiGetDeviceInterfaceDetail(
                deviceInfoSet,
                ref interfaceData,
                IntPtr.Zero,
                0,
                out uint requiredSize,
                IntPtr.Zero);

            if (requiredSize == 0)
            {
                return null;
            }

            IntPtr detailData = Marshal.AllocHGlobal((int)requiredSize);
            try
            {
                Marshal.WriteInt32(detailData, IntPtr.Size == 8 ? 8 : 6);
                if (!SetupDiGetDeviceInterfaceDetail(
                        deviceInfoSet,
                        ref interfaceData,
                        detailData,
                        requiredSize,
                        out _,
                        IntPtr.Zero))
                {
                    return null;
                }

                return Marshal.PtrToStringUni(IntPtr.Add(detailData, 4));
            }
            finally
            {
                Marshal.FreeHGlobal(detailData);
            }
        }

        private static bool TryGetCapabilities(SafeFileHandle handle, out HIDP_CAPS caps)
        {
            caps = default;
            if (!HidD_GetPreparsedData(handle, out IntPtr preparsedData))
            {
                return false;
            }

            try
            {
                return HidP_GetCaps(preparsedData, out caps) == HIDP_STATUS_SUCCESS;
            }
            finally
            {
                HidD_FreePreparsedData(preparsedData);
            }
        }

        private static string GetProductName(SafeFileHandle handle)
        {
            IntPtr buffer = Marshal.AllocHGlobal(256);
            try
            {
                if (HidD_GetProductString(handle, buffer, 256))
                {
                    return Marshal.PtrToStringUni(buffer) ?? string.Empty;
                }

                return string.Empty;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private static bool IsSupportedDeviceIdentity(string path, string productName)
        {
            return productName.Contains("Moshushou", StringComparison.OrdinalIgnoreCase) ||
                   productName.Contains("FakerInput", StringComparison.OrdinalIgnoreCase) ||
                   path.Contains("vid_18d1&pid_9400", StringComparison.OrdinalIgnoreCase);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SP_DEVICE_INTERFACE_DATA
        {
            public int cbSize;
            public Guid InterfaceClassGuid;
            public int Flags;
            public UIntPtr Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HIDP_CAPS
        {
            public ushort Usage;
            public ushort UsagePage;
            public ushort InputReportByteLength;
            public ushort OutputReportByteLength;
            public ushort FeatureReportByteLength;

            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 17)]
            public ushort[] Reserved;

            public ushort NumberLinkCollectionNodes;
            public ushort NumberInputButtonCaps;
            public ushort NumberInputValueCaps;
            public ushort NumberInputDataIndices;
            public ushort NumberOutputButtonCaps;
            public ushort NumberOutputValueCaps;
            public ushort NumberOutputDataIndices;
            public ushort NumberFeatureButtonCaps;
            public ushort NumberFeatureValueCaps;
            public ushort NumberFeatureDataIndices;
        }

        [DllImport("hid.dll")]
        private static extern void HidD_GetHidGuid(out Guid hidGuid);

        [DllImport("hid.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool HidD_GetPreparsedData(SafeFileHandle hidDeviceObject, out IntPtr preparsedData);

        [DllImport("hid.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool HidD_FreePreparsedData(IntPtr preparsedData);

        [DllImport("hid.dll")]
        private static extern int HidP_GetCaps(IntPtr preparsedData, out HIDP_CAPS capabilities);

        [DllImport("hid.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool HidD_GetProductString(SafeFileHandle hidDeviceObject, IntPtr buffer, int bufferLength);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr SetupDiGetClassDevs(
            ref Guid classGuid,
            IntPtr enumerator,
            IntPtr hwndParent,
            uint flags);

        [DllImport("setupapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetupDiEnumDeviceInterfaces(
            IntPtr deviceInfoSet,
            IntPtr deviceInfoData,
            ref Guid interfaceClassGuid,
            uint memberIndex,
            ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData);

        [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetupDiGetDeviceInterfaceDetail(
            IntPtr deviceInfoSet,
            ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData,
            IntPtr deviceInterfaceDetailData,
            uint deviceInterfaceDetailDataSize,
            out uint requiredSize,
            IntPtr deviceInfoData);

        [DllImport("setupapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(
            string fileName,
            uint desiredAccess,
            uint shareMode,
            IntPtr securityAttributes,
            uint creationDisposition,
            uint flagsAndAttributes,
            IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool WriteFile(
            SafeFileHandle file,
            byte[] buffer,
            int numberOfBytesToWrite,
            out int numberOfBytesWritten,
            IntPtr overlapped);

        private static readonly IntPtr INVALID_HANDLE_VALUE = new(-1);
        private const int HIDP_STATUS_SUCCESS = 0x00110000;
        private const int ERROR_NO_MORE_ITEMS = 259;
        private const uint DIGCF_PRESENT = 0x00000002;
        private const uint DIGCF_DEVICEINTERFACE = 0x00000010;
        private const uint GENERIC_READ = 0x80000000;
        private const uint GENERIC_WRITE = 0x40000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 3;
    }
}
