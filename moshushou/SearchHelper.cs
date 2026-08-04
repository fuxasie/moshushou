using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using moshushou.Input;

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

        private readonly IInputBackend _inputBackend;
        private readonly SearchConfig _config;
        private readonly Action<string> _logAction;

        public SearchHelper(SearchConfig config, IInputBackend inputBackend, Action<string> logAction = null)
        {
            _config = config;
            _inputBackend = inputBackend ?? throw new ArgumentNullException(nameof(inputBackend));
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
            var swTotal = System.Diagnostics.Stopwatch.StartNew();
            System.Diagnostics.Debug.WriteLine($"\n🔎 [SearchInApp] ====== 开始 ====== Text='{searchText}', IsWework={isWework}");
            try
            {
                string appName = isWework ? "企业微信" : "微信";
                IntPtr hwnd = IntPtr.Zero;

                // 初始化变量，防止 retry 逻辑报错
                string className = isWework ? _config.WeworkWindowClassName : _config.WechatWindowClassName;
                string processName = isWework ? "WXWork" : _config.WeChatProcessName;

                System.Diagnostics.Debug.WriteLine($"🔎 [SearchInApp] Step1: 查找窗口句柄... App={appName}, Proc={processName}, Class={className}");
                var swStep = System.Diagnostics.Stopwatch.StartNew();

                if (isWework)
                {
                    hwnd = FindAndCacheWindowHandle(appName, processName, className);
                }
                else
                {
                    Log($"ℹ️ 尝试连接微信主窗口...");
                    hwnd = FindAndCacheWindowHandle(appName, processName, className);
                }

                System.Diagnostics.Debug.WriteLine($"🔎 [SearchInApp] Step1完成: Hwnd={hwnd}, 耗时={swStep.ElapsedMilliseconds}ms");

                if (hwnd == IntPtr.Zero)
                {
                    Log($"❌ 未能找到 {appName} 或相关搜索窗口。");
                    System.Diagnostics.Debug.WriteLine($"❌ [SearchInApp] Step1失败: 窗口句柄为零，退出");
                    return false;
                }

                if (token.IsCancellationRequested) return false;

                System.Diagnostics.Debug.WriteLine($"🔎 [SearchInApp] Step2: 激活窗口... Hwnd={hwnd}, 当前前台={GetForegroundWindow()}");
                swStep.Restart();

                if (!await ForceActivateWindowAsync(hwnd, token))
                {
                    Log($"❌ 激活 {appName} 窗口失败，尝试清除缓存后重试...");
                    System.Diagnostics.Debug.WriteLine($"⚠️ [SearchInApp] Step2首次失败，耗时={swStep.ElapsedMilliseconds}ms，清除缓存重试...");
                    _windowHandleCache.Remove(appName);
                    hwnd = FindAndCacheWindowHandle(appName, processName, className);
                    System.Diagnostics.Debug.WriteLine($"🔎 [SearchInApp] Step2重试: 新Hwnd={hwnd}");
                    if (hwnd == IntPtr.Zero || !await ForceActivateWindowAsync(hwnd, token))
                    {
                        Log($"❌ 重试后依然无法激活 {appName} 窗口。");
                        System.Diagnostics.Debug.WriteLine($"❌ [SearchInApp] Step2彻底失败，退出");
                        return false;
                    }
                }

                System.Diagnostics.Debug.WriteLine($"✅ [SearchInApp] Step2完成: 窗口已激活, 耗时={swStep.ElapsedMilliseconds}ms, 当前前台={GetForegroundWindow()}");

                if (token.IsCancellationRequested) return false;

                // *** ⭐ 核心修复 ⭐ ***
                // 使用传入的 searchText 参数来设置剪贴板，确保内容正确无误
                System.Diagnostics.Debug.WriteLine($"🔎 [SearchInApp] Step3: 设置剪贴板... Text='{searchText}'");
                swStep.Restart();

                if (!await SetClipboardWithRetryAsync(searchText, token))
                {
                    Log("❌ 无法设置剪贴板，已达最大重试次数。");
                    System.Diagnostics.Debug.WriteLine($"❌ [SearchInApp] Step3失败: 剪贴板设置失败, 耗时={swStep.ElapsedMilliseconds}ms");
                    return false;
                }

                System.Diagnostics.Debug.WriteLine($"✅ [SearchInApp] Step3完成: 剪贴板已设置, 耗时={swStep.ElapsedMilliseconds}ms");

                System.Diagnostics.Debug.WriteLine($"🔎 [SearchInApp] Step4: 执行搜索序列(Ctrl+F → 清空 → Ctrl+V)...");
                swStep.Restart();

                await PerformSearchSequenceAsync(token);

                System.Diagnostics.Debug.WriteLine($"✅ [SearchInApp] Step4完成: 搜索序列已执行, 耗时={swStep.ElapsedMilliseconds}ms");
                System.Diagnostics.Debug.WriteLine($"✅ [SearchInApp] ====== 全部完成 ====== 总耗时={swTotal.ElapsedMilliseconds}ms");
                return true;
            }
            catch (TaskCanceledException)
            {
                Log("🛑 搜索已被用户打断。");
                System.Diagnostics.Debug.WriteLine($"🛑 [SearchInApp] 被取消, 已耗时={swTotal.ElapsedMilliseconds}ms");
                return false;
            }
            catch (Exception ex)
            {
                Log($"❌ SearchInAppAsync 异常: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"💥 [SearchInApp] 异常: {ex}, 已耗时={swTotal.ElapsedMilliseconds}ms");
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
            // 1. 先构建目标进程名称集合（缓存验证也需要用到）
            var targetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { processName };
            if (processName.IndexOf("WeChat", StringComparison.OrdinalIgnoreCase) >= 0 || 
                processName.IndexOf("Weixin", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                targetNames.Add("Weixin");
                targetNames.Add("WeChat");
            }

            // 2. 缓存命中检查（含进程归属验证）
            if (_windowHandleCache.TryGetValue(appName, out IntPtr cachedHwnd))
            {
                if (IsWindow(cachedHwnd) && IsWindowVisible(cachedHwnd))
                {
                    // 🔧 [修复] 真正验证进程归属，防止窗口关闭重开后句柄失效或被复用
                    bool processMatch = false;
                    try
                    {
                        GetWindowThreadProcessId(cachedHwnd, out uint cachedPid);
                        var proc = Process.GetProcessById((int)cachedPid);
                        processMatch = targetNames.Contains(proc.ProcessName);
                    }
                    catch { /* 进程已退出 */ }

                    if (processMatch)
                    {
                        return cachedHwnd;
                    }
                    else
                    {
                        Log($"⚠️ 缓存的 {appName} 窗口句柄进程不匹配，重新搜索...");
                        _windowHandleCache.Remove(appName);
                    }
                }
                else
                {
                    Log($"ℹ️ 缓存的 {appName} 窗口句柄已失效，重新搜索...");
                    _windowHandleCache.Remove(appName);
                }
            }

            IntPtr foundHwnd = IntPtr.Zero;

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
                    if (title.Contains("企业微信", StringComparison.OrdinalIgnoreCase)) currentScore += 20;
                    // 排除掉托盘气泡、悬浮球等小窗口（通过尺寸判断，暂略，简单场景可见性+类名通常够了）
                }
                // [个人微信特有逻辑]
                else 
                {
                    if (clazz.Equals("WeChatMainWndForPC", StringComparison.OrdinalIgnoreCase)) currentScore += 50;
                    if (title.Equals("微信", StringComparison.OrdinalIgnoreCase) ||
                        title.Equals("WeChat", StringComparison.OrdinalIgnoreCase))
                    {
                        currentScore += 20;
                    }
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
            IntPtr currentFg = GetForegroundWindow();
            System.Diagnostics.Debug.WriteLine($"🪟 [Activate] 目标={hwnd}, 当前前台={currentFg}, 相同={currentFg == hwnd}");
            if (currentFg == hwnd) return true;

            // ✅ 增加重试次数，解决有时无法调用企业微信的问题
            for (int attempt = 0; attempt < 3; attempt++)
            {
                if (token.IsCancellationRequested) return false;

                bool isMinimized = IsIconic(hwnd);
                System.Diagnostics.Debug.WriteLine($"🪟 [Activate] 尝试{attempt + 1}/3: IsIconic={isMinimized}");

                if (isMinimized) ShowWindow(hwnd, SW_RESTORE);
                else ShowWindow(hwnd, SW_SHOW);

                bool setResult = SetForegroundWindow(hwnd);
                System.Diagnostics.Debug.WriteLine($"🪟 [Activate] SetForegroundWindow返回={setResult}");
                
                // ✅ 增加等待时间，让系统有足够时间切换窗口
                await Task.Delay(150, token);

                IntPtr afterFg = GetForegroundWindow();
                System.Diagnostics.Debug.WriteLine($"🪟 [Activate] 等待后前台={afterFg}, 匹配={afterFg == hwnd}");
                if (afterFg == hwnd)
                {
                    return true;
                }

                Log($"⚠️ 窗口激活尝试 {attempt + 1}/3 失败，重试中...");
                await Task.Delay(100, token);
            }

            bool finalResult = GetForegroundWindow() == hwnd;
            System.Diagnostics.Debug.WriteLine($"🪟 [Activate] 最终结果={finalResult}");
            return finalResult;
        }



        // SearchHelper.cs

        private async Task PerformSearchSequenceAsync(CancellationToken token)
        {
            if (token.IsCancellationRequested) return;

            IntPtr fgBefore = GetForegroundWindow();
            System.Diagnostics.Debug.WriteLine($"⌨️ [SearchSeq] 开始，当前前台窗口={fgBefore}");

            // 1. 激活搜索框
            System.Diagnostics.Debug.WriteLine($"⌨️ [SearchSeq] Step1: 发送 Ctrl+F...");
            _inputBackend.KeyChord(InputKey.LeftControl, InputKey.F);
            await Task.Delay(150, token);

            IntPtr fgAfterCtrlF = GetForegroundWindow();
            System.Diagnostics.Debug.WriteLine($"⌨️ [SearchSeq] Step1完成: Ctrl+F后前台窗口={fgAfterCtrlF}, 变化={fgBefore != fgAfterCtrlF}");

            // 2. 防御性清空
            System.Diagnostics.Debug.WriteLine($"⌨️ [SearchSeq] Step2: 防御性清空(Backspace)...");
            _inputBackend.KeyPress(InputKey.Backspace);
            await Task.Delay(20, token);

            // 3. 全选并删除
            System.Diagnostics.Debug.WriteLine($"⌨️ [SearchSeq] Step3: 全选删除(Ctrl+A → Backspace)...");
            _inputBackend.KeyChord(InputKey.LeftControl, InputKey.A);
            await Task.Delay(30, token);
            _inputBackend.KeyPress(InputKey.Backspace);
            await Task.Delay(30, token);

            if (token.IsCancellationRequested) return;

            // 4. 粘贴新内容
            System.Diagnostics.Debug.WriteLine($"⌨️ [SearchSeq] Step4: 粘贴(Ctrl+V)...");
            _inputBackend.KeyChord(InputKey.LeftControl, InputKey.V);

            // ✅ [优化] 粘贴后等待列表渲染
            await Task.Delay(200, token);

            IntPtr fgAfterPaste = GetForegroundWindow();
            System.Diagnostics.Debug.WriteLine($"⌨️ [SearchSeq] 全部完成: 粘贴后前台窗口={fgAfterPaste}");
        }



        private async Task<bool> SetClipboardWithRetryAsync(string text, CancellationToken token)
        {
            const int maxRetries = 20;
            const int delayMs = 50; // ✅ [优化] 稍微增加重试间隔

            for (int i = 0; i < maxRetries; i++)
            {
                if (token.IsCancellationRequested) return false;

                bool success = false;
                string failReason = null;
                var thread = new Thread(() =>
                {
                    try
                    {
                        // ✅ [修复] 显式清空剪贴板，防止旧数据残留或无法写入
                        try { Clipboard.Clear(); } catch { }

                        // 写入新数据
                        // ⭐ [关键修复] copy=true，将数据 flush 到系统剪贴板
                        // copy=false 时数据仅在当前 STA 线程存活期间有效，
                        // 线程退出后其他应用（如企业微信）的 Ctrl+V 读到的是空数据
                        Clipboard.SetDataObject(text, true);

                        // ✅ [修复] 立即读取校验，确保写入成功
                        if (Clipboard.ContainsText() && Clipboard.GetText() == text)
                        {
                            success = true;
                        }
                        else
                        {
                            bool hasText = Clipboard.ContainsText();
                            string actual = hasText ? Clipboard.GetText() : "<无文本>";
                            failReason = $"校验失败: HasText={hasText}, Actual='{actual}'";
                        }
                    }
                    catch (COMException comEx) { failReason = $"COMException: {comEx.Message}"; }
                    catch (Exception ex) { failReason = $"Exception: {ex.Message}"; }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join(3000); // 最多等3秒，防止线程死锁

                if (success)
                {
                    if (i > 0) System.Diagnostics.Debug.WriteLine($"📋 [Clipboard] 第{i + 1}次重试成功");
                    return true;
                }

                // 每5次或最后一次打印日志，避免刷屏
                if (i % 5 == 0 || i == maxRetries - 1)
                {
                    System.Diagnostics.Debug.WriteLine($"📋 [Clipboard] 第{i + 1}/{maxRetries}次失败: {failReason ?? "未知"}");
                }

                await Task.Delay(delayMs, token);
            }
            System.Diagnostics.Debug.WriteLine($"❌ [Clipboard] 全部{maxRetries}次重试均失败!");
            return false;
        }
    }
}
