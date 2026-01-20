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

        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

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
        public async Task<bool> SearchInAppAsync(string searchText, bool isWework, CancellationToken token = default)
        {
            try
            {
                string appName = isWework ? "企业微信" : "微信";
                IntPtr hwnd = IntPtr.Zero;

                // 初始化变量，防止 retry 逻辑报错
                string className = isWework ? _config.WeworkWindowClassName : _config.WechatWindowClassName;
                string processName = isWework ? "WXWork" : _config.WeChatProcessName;

                if (isWework)
                {
                    // 企业微信
                    hwnd = FindAndCacheWindowHandle(appName, processName, className);
                }
                else
                {
                    // 微信 (4.0): 回退到查找主窗口
                    Log($"ℹ️ 尝试连接微信主窗口...");
                    hwnd = FindAndCacheWindowHandle(appName, processName, className);
                }

                if (hwnd == IntPtr.Zero)
                {
                    Log($"❌ 未能找到 {appName} 或相关搜索窗口。");
                    return false;
                }

                if (token.IsCancellationRequested) return false;

                if (!await ForceActivateWindowAsync(hwnd, token))
                {
                    Log($"❌ 激活 {appName} 窗口失败，尝试清除缓存后重试...");
                    _windowHandleCache.Remove(appName);
                    hwnd = FindAndCacheWindowHandle(appName, processName, className);
                    if (hwnd == IntPtr.Zero || !await ForceActivateWindowAsync(hwnd, token))
                    {
                        Log($"❌ 重试后依然无法激活 {appName} 窗口。");
                        return false;
                    }
                }

                if (token.IsCancellationRequested) return false;

                // *** ⭐ 核心修复 ⭐ ***
                // 使用传入的 searchText 参数来设置剪贴板，确保内容正确无误
                if (!await SetClipboardWithRetryAsync(searchText, token))
                {
                    Log("❌ 无法设置剪贴板，已达最大重试次数。");
                    return false;
                }

                await PerformSearchSequenceAsync(token);
                return true;
            }
            catch (TaskCanceledException)
            {
                Log("🛑 搜索已被用户打断。");
                return false;
            }
            catch (Exception ex)
            {
                Log($"❌ SearchInAppAsync 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// [兼容性包装] 同步版本的 SearchInApp，用于未改造的旧代码
        /// </summary>
        public bool SearchInApp(string searchText, bool isWework)
        {
            // 同步等待异步任务
            return SearchInAppAsync(searchText, isWework, CancellationToken.None).GetAwaiter().GetResult();
        }


        private IntPtr FindAndCacheWindowHandle(string appName, string processName, string className)
        {
            if (_windowHandleCache.TryGetValue(appName, out IntPtr cachedHwnd) && IsWindow(cachedHwnd))
            {
                return cachedHwnd;
            }

            IntPtr foundHwnd = IntPtr.Zero;

            // 1. 优先尝试进程名查找 (适用于 WeChat 4.0 Qt)
            // 策略：先尝试配置的名称，如果失败，尝试常见名称 "Weixin", "WeChat"
            var targetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { processName };
            
            // ⚠️ 修正：仅当目标确实是微信时，才添加微信的别名
            // 之前的逻辑会导致在搜企业微信时(WXWork)如果没找到，会自动fallback到微信，导致逻辑错乱
            if (processName.IndexOf("WeChat", StringComparison.OrdinalIgnoreCase) >= 0 || 
                processName.IndexOf("Weixin", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                targetNames.Add("Weixin");
                targetNames.Add("WeChat");
            }

            foreach (var targetName in targetNames)
            {
                try
                {
                    var processes = Process.GetProcessesByName(targetName);
                    if (processes.Length == 0) continue;

                    Log($"  -> 尝试查找进程: {targetName} (找到 {processes.Length} 个)");

                    foreach (var p in processes)
                    {
                        // ✅ [关键修复] 强制刷新进程信息，确保获取最新的窗口句柄
                        // 如果程序先于目标应用启动，Process 对象可能缓存了旧的（无效的）MainWindowHandle
                        try { p.Refresh(); } catch { }

                        // 调试日志
                        Log($"     - PID: {p.Id}, Title: '{p.MainWindowTitle}', Handle: {p.MainWindowHandle}");

                        // 宽松检查：只要有句柄就行，Qt 窗口有时候 Title 为空或读取不到
                        if (p.MainWindowHandle != IntPtr.Zero)
                        {
                            foundHwnd = p.MainWindowHandle;
                            // 如果找到了有标题的，优先用它（更可能是主窗口）
                            if (!string.IsNullOrEmpty(p.MainWindowTitle))
                            {
                                Log($"     ✅ 命中主窗口 (有标题): {p.MainWindowTitle}");
                                break; 
                            }
                            else
                            {
                                Log($"     ⚠️ 命中窗口 (无标题)，作为备选...");
                            }
                        }
                    }

                    if (foundHwnd != IntPtr.Zero) break; // 找到了就跳出
                }
                catch (Exception ex)
                {
                    Log($"  -> 进程查找出错 ({targetName}): {ex.Message}");
                }
            }

            // 2. 如果进程没找到，回退到 FlaUI / 类名查找 (适用于 企业微信 或 旧版微信)
            if (foundHwnd == IntPtr.Zero && !string.IsNullOrEmpty(className))
            {
                try
                {
                    using (var automation = new UIA3Automation())
                    {
                        var window = automation.GetDesktop().FindFirstChild(cf => cf.ByClassName(className))?.AsWindow();
                        if (window != null && window.IsAvailable)
                        {
                            foundHwnd = window.Properties.NativeWindowHandle.ValueOrDefault;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"  -> FlaUI 查找时出错: {ex.Message}");
                }
            }

            if (foundHwnd != IntPtr.Zero)
            {
                _windowHandleCache[appName] = foundHwnd;
                return foundHwnd;
            }

            return IntPtr.Zero;
        }

        /// <summary>
        /// ✅ [优化版] 简化激活逻辑，不使用 AttachThreadInput
        /// </summary>
        private async Task<bool> ForceActivateWindowAsync(IntPtr hwnd, CancellationToken token)
        {
            if (hwnd == GetForegroundWindow()) return true;

            if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE);
            else ShowWindow(hwnd, SW_SHOW);

            // ✅ [优化] 不再使用 AttachThreadInput，只调用一次 SetForegroundWindow
            SetForegroundWindow(hwnd);
            await Task.Delay(100, token);

            return GetForegroundWindow() == hwnd;
        }



        // SearchHelper.cs

        private async Task PerformSearchSequenceAsync(CancellationToken token)
        {
            if (token.IsCancellationRequested) return;

            // 1. 激活搜索框
            _inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_F);
            await Task.Delay(150, token);

            // 2. 防御性清空
            _inputSimulator.Keyboard.KeyPress(VirtualKeyCode.BACK);
            await Task.Delay(20, token);

            // 3. 全选并删除
            _inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_A);
            await Task.Delay(30, token);
            _inputSimulator.Keyboard.KeyPress(VirtualKeyCode.BACK);
            await Task.Delay(30, token);

            if (token.IsCancellationRequested) return;

            // 4. 粘贴新内容
            _inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);

            // ✅ [优化] 粘贴后等待列表渲染
            await Task.Delay(200, token);
        }



        private async Task<bool> SetClipboardWithRetryAsync(string text, CancellationToken token)
        {
            const int maxRetries = 20;
            const int delayMs = 25;

            for (int i = 0; i < maxRetries; i++)
            {
                if (token.IsCancellationRequested) return false;

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

                await Task.Delay(delayMs, token);
            }
            return false;
        }
    }
}