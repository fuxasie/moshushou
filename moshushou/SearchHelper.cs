using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using WindowsInput;
using WindowsInput.Native;

namespace moshushou
{
    public class SearchHelper
    {
        #region Win32 API
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", SetLastError = true)] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")][return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindow(IntPtr hWnd);

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;
        #endregion

        private static readonly Dictionary<string, IntPtr> _windowHandleCache = new Dictionary<string, IntPtr>();

        private readonly InputSimulator _inputSimulator;
        private readonly SearchConfig _config;
        private readonly Action<string> _logAction;

        public SearchHelper(SearchConfig config, Action<string> logAction = null)
        {
            _inputSimulator = new InputSimulator();
            _config = config;
            _logAction = logAction;
        }

        private void Log(string message) => _logAction?.Invoke(message);

        /// <summary>
        /// 【V2 - 已修复】核心搜索方法。
        /// 新增 searchText 参数，使其不再依赖外部剪贴板状态，彻底解决剪贴板污染问题。
        /// </summary>
        public bool SearchInApp(string searchText, bool isWework)
        {
            string appName = isWework ? "企业微信" : "微信";
            string className = isWework ? _config.WeworkWindowClassName : _config.WechatWindowClassName;

            IntPtr hwnd = FindAndCacheWindowHandle(appName, className);
            if (hwnd == IntPtr.Zero)
            {
                Log($"❌ 未能找到 {appName} 窗口。");
                return false;
            }

            if (!ForceActivateWindow(hwnd))
            {
                Log($"❌ 激活 {appName} 窗口失败，尝试清除缓存后重试...");
                _windowHandleCache.Remove(appName);
                hwnd = FindAndCacheWindowHandle(appName, className);
                if (hwnd == IntPtr.Zero || !ForceActivateWindow(hwnd))
                {
                    Log($"❌ 重试后依然无法激活 {appName} 窗口。");
                    return false;
                }
            }

            // *** ⭐ 核心修复 ⭐ ***
            // 使用传入的 searchText 参数来设置剪贴板，确保内容正确无误
            if (!SetClipboardWithRetry(searchText))
            {
                Log("❌ 无法设置剪贴板，已达最大重试次数。");
                return false;
            }

            PerformSearchSequence();
            return true;
        }


        private IntPtr FindAndCacheWindowHandle(string appName, string className)
        {
            if (_windowHandleCache.TryGetValue(appName, out IntPtr cachedHwnd) && IsWindow(cachedHwnd))
            {
                return cachedHwnd;
            }

            try
            {
                using (var automation = new UIA3Automation())
                {
                    var window = automation.GetDesktop().FindFirstChild(cf => cf.ByClassName(className))?.AsWindow();
                    if (window != null && window.IsAvailable)
                    {
                        IntPtr foundHwnd = window.Properties.NativeWindowHandle.ValueOrDefault;
                        _windowHandleCache[appName] = foundHwnd;
                        return foundHwnd;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  -> FlaUI 查找时出错: {ex.Message}");
            }
            return IntPtr.Zero;
        }

        private bool ForceActivateWindow(IntPtr hwnd)
        {
            if (hwnd == GetForegroundWindow()) return true;

            if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE);
            else ShowWindow(hwnd, SW_SHOW);

            uint currentThreadId = GetCurrentThreadId();
            uint targetThreadId = GetWindowThreadProcessId(hwnd, out _);

            try
            {
                AttachThreadInput(currentThreadId, targetThreadId, true);
                SetForegroundWindow(hwnd);

                var stopwatch = Stopwatch.StartNew();
                while (GetForegroundWindow() != hwnd && stopwatch.ElapsedMilliseconds < _config.ActivationTimeoutMs)
                {
                    Thread.Sleep(10);
                }
                stopwatch.Stop();

                return GetForegroundWindow() == hwnd;
            }
            finally
            {
                AttachThreadInput(currentThreadId, targetThreadId, false);
            }
        }

        private void PerformSearchSequence()
        {
            _inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_F);
            Thread.Sleep(_config.DelayAfterCtrlF);
            _inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_A);
            Thread.Sleep(_config.DelayKeyboardAction);
            _inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
        }

        private bool SetClipboardWithRetry(string text)
        {
            const int maxRetries = 20;
            const int delayMs = 25;

            for (int i = 0; i < maxRetries; i++)
            {
                bool success = false;
                var thread = new Thread(() =>
                {
                    try
                    {
                        Clipboard.SetDataObject(text, true);
                        success = true;
                    }
                    catch (COMException) { success = false; }
                    catch { success = false; }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

                if (success) return true;

                Thread.Sleep(delayMs);
            }
            return false;
        }
    }
}