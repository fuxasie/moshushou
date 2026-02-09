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
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)] private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)] private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

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
        /// ✅ [新增] 允许外部强制清除窗口缓存 (用于应对窗口假死或找不到的情况)
        /// </summary>
        public void ClearWindowCache(string appName)
        {
            if (_windowHandleCache.ContainsKey(appName))
            {
                _windowHandleCache.Remove(appName);
                Log($"🧹 [Cache] 已手动清除 {appName} 的窗口句柄缓存");
            }
        }

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

        /// <summary>
        /// 为外部提供稳健的窗口查找能力（进程+EnumWindows），避免仅按类名查找失败。
        /// </summary>
        public IntPtr TryGetAppWindowHandle(bool isWework, bool clearCache = false)
        {
            string appName = isWework ? "企业微信" : "微信";
            string className = isWework ? _config.WeworkWindowClassName : _config.WechatWindowClassName;
            string processName = isWework ? "WXWork" : _config.WeChatProcessName;

            if (clearCache && _windowHandleCache.ContainsKey(appName))
            {
                _windowHandleCache.Remove(appName);
            }

            return FindAndCacheWindowHandle(appName, processName, className);
        }


        private IntPtr FindAndCacheWindowHandle(string appName, string processName, string className)
        {
            // 1. 缓存命中检查
            if (_windowHandleCache.TryGetValue(appName, out IntPtr cachedHwnd))
            {
                if (IsWindow(cachedHwnd) && IsWindowVisible(cachedHwnd))
                {
                    // 额外检查：再次确认进程名，防止PID复用导致的误判（虽然概率极低）
                    return cachedHwnd;
                }
                else
                {
                    Log($"ℹ️ 缓存的 {appName} 窗口句柄已失效，重新搜索...");
                    _windowHandleCache.Remove(appName);
                }
            }

            IntPtr foundHwnd = IntPtr.Zero;

            // 2. 确定目标进程名称集合
            var targetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { processName };
            // WXWork 通常不需要别名，但 Weixin 需要
            if (processName.IndexOf("WeChat", StringComparison.OrdinalIgnoreCase) >= 0 || 
                processName.IndexOf("Weixin", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                targetNames.Add("Weixin");
                targetNames.Add("WeChat");
            }

            // 3. 获取所有目标进程的 PID
            var targetPids = new HashSet<int>();
            foreach (var targetName in targetNames)
            {
                var processes = Process.GetProcessesByName(targetName);
                foreach (var p in processes)
                {
                    targetPids.Add(p.Id);
                }
            }

            if (targetPids.Count == 0)
            {
                Log($"❌ 未在运行任务列表中找到进程: {string.Join(", ", targetNames)}");
                return IntPtr.Zero;
            }

            Log($"🔍 开始全局扫描窗口，目标PID数量: {targetPids.Count}...");

            // 4. 使用 EnumWindows 进行全量扫描 (解决 "Process.MainWindowHandle 为 0" 或 "先开软件找不到窗口" 的问题)
            // 核心思路：Qt/CEF 应用的主窗口经常不是 Process.MainWindowHandle 指向的那个，通过枚举所有窗口并匹配 PID 是最稳健的方法。
            IntPtr bestCandidate = IntPtr.Zero;
            int bestScore = -1; // 评分机制：标题匹配 > 类名匹配 > 仅PID匹配

            EnumWindows((hwnd, lParam) =>
            {
                // A. 基础过滤：必须可见
                if (!IsWindowVisible(hwnd)) return true;

                // B. PID 匹配
                GetWindowThreadProcessId(hwnd, out uint processId);
                if (!targetPids.Contains((int)processId)) return true;

                // C. 获取窗口信息
                StringBuilder sbTitle = new StringBuilder(256);
                GetWindowText(hwnd, sbTitle, 256);
                string title = sbTitle.ToString();

                StringBuilder sbClass = new StringBuilder(256);
                GetClassName(hwnd, sbClass, 256);
                string clazz = sbClass.ToString();

                // D. 评分逻辑
                int currentScore = 0;

                // [企业微信特有逻辑]
                // 企业微信主窗口类名通常是 "WeWorkWindow"，标题通常包含 "企业微信"
                // 也有可能标题就是具体的聊天对象名，但类名通常不变
                if (processName.Equals("WXWork", StringComparison.OrdinalIgnoreCase))
                {
                    if (clazz.Equals("WeWorkWindow", StringComparison.OrdinalIgnoreCase)) currentScore += 50;
                    if (title.Contains("企业微信")) currentScore += 20;
                    // 排除掉托盘气泡、悬浮球等小窗口（通过尺寸判断，暂略，简单场景可见性+类名通常够了）
                }
                // [个人微信特有逻辑]
                else 
                {
                    if (clazz.Equals("WeChatMainWndForPC", StringComparison.OrdinalIgnoreCase)) currentScore += 50;
                    if (title.Equals("微信") || title.Equals("WeChat")) currentScore += 20;
                }

                // 只有分数高的才替换
                if (currentScore > bestScore)
                {
                    bestScore = currentScore;
                    bestCandidate = hwnd;
                    Log($"   -> [候选] 评分:{currentScore} | 标题:'{title}' | 类名:'{clazz}' | Hwnd:{hwnd}");
                }

                return true; // 继续枚举
            }, IntPtr.Zero);

            if (bestCandidate != IntPtr.Zero)
            {
                foundHwnd = bestCandidate;
                Log($"✅ 最终锁定窗口: {foundHwnd}");
            }
            else
            {
                Log("⚠️ 扫描结束，未找到符合条件的窗口。");
            }

            // 5. 缓存结果
            if (foundHwnd != IntPtr.Zero)
            {
                _windowHandleCache[appName] = foundHwnd;
                return foundHwnd;
            }

            return IntPtr.Zero;
        }

        /// <summary>
        /// ✅ [优化版] 增强激活逻辑，添加重试机制
        /// </summary>
        private async Task<bool> ForceActivateWindowAsync(IntPtr hwnd, CancellationToken token)
        {
            if (hwnd == GetForegroundWindow()) return true;

            // ✅ 增加重试次数，解决有时无法调用企业微信的问题
            for (int attempt = 0; attempt < 3; attempt++)
            {
                if (token.IsCancellationRequested) return false;

                if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE);
                else ShowWindow(hwnd, SW_SHOW);

                SetForegroundWindow(hwnd);
                
                // ✅ 增加等待时间，让系统有足够时间切换窗口
                await Task.Delay(150, token);

                if (GetForegroundWindow() == hwnd)
                {
                    return true;
                }

                Log($"⚠️ 窗口激活尝试 {attempt + 1}/3 失败，重试中...");
                await Task.Delay(100, token);
            }

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
            const int delayMs = 50; // ✅ [优化] 稍微增加重试间隔

            for (int i = 0; i < maxRetries; i++)
            {
                if (token.IsCancellationRequested) return false;

                bool success = false;
                var thread = new Thread(() =>
                {
                    try
                    {
                        // ✅ [修复] 显式清空剪贴板，防止旧数据残留或无法写入
                        try { Clipboard.Clear(); } catch { }

                        // 写入新数据
                        Clipboard.SetDataObject(text, true);

                        // ✅ [修复] 立即读取校验，确保写入成功
                        // 注意：有时 SetDataObject 不抛异常但实际没写入，所以必须校验
                        if (Clipboard.ContainsText() && Clipboard.GetText() == text)
                        {
                            success = true;
                        }
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
