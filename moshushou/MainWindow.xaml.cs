using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.ObjectModel; 
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using WindowsInput;
using WindowsInput.Native;

namespace moshushou
{
    public partial class MainWindow : Window, IDisposable
    {
        private Dictionary<string, List<string>> _storeData = new Dictionary<string, List<string>>();
        private Dictionary<string, string> _exportedFilePaths = new Dictionary<string, string>();
        private Point _startPoint;
        private bool _isDragging = false;
        private int _copyingFlag = 0;
        private readonly object _dataLock = new object();
        private string _exportDirectory;
        private int _currentSelectedIndex = -1;
        private List<TreeViewNode> _flatNodeList = new List<TreeViewNode>();
        private TreeViewNode _currentSelectedNode = null;
        private List<string> _currentFilter = new List<string>();
        private readonly ScreenshotHelper _screenshotHelper;



        // 全局快捷键相关
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_UP = 9001;
        private const int HOTKEY_DOWN = 9002;
        private const int HOTKEY_LEFT = 9003;
        private const int HOTKEY_RIGHT = 9004;
        private const int HOTKEY_ENTER = 9005;
        private const int HOTKEY_QUOTE = 9006;
        private IntPtr _windowHandle;
        private HwndSource _source;
        private bool _globalHotkeysRegistered = false;




        // 微信/企业微信切换状态
        private bool _isWeworkTurn = true;



        // ✅ [新增] 记录上一次成功进入的群名 (用于极速模式)
        private string _lastEnteredGroupName = null;


        // 搜索配置
        private SearchConfig _searchConfig;
        private SearchHelper _searchHelper;

        private int _searchInProgress = 0; // ✅ 改用 int 配合 Interlocked
        private readonly InputSimulator _inputSimulator;

        private bool _currentItemPasted = false;
        private string _lastPastedStoreName = null;

        // ✅ [新增] 记录上一次成功的聊天窗口句柄 (用于极速模式抢焦点)
        private IntPtr _lastChatWindowHandle = IntPtr.Zero; // <--- 新增行

        // 固定话术
        private const string FIXED_MESSAGE = "现同步未发货预警，超时未交件会考核处罚，请尽快处理转出,已售后的及时发起拦截。（注：未处理售后请勿虚假拦截，核实虚假正常考核处罚。字节超时未发出总部将发起拦截）";

        // 商家信息
        private List<BusinessInfo> _businessInfoList = new List<BusinessInfo>();


        
        // ✅ 新增：支持动态更新的数据集合
        private ObservableCollection<TreeViewNode> _treeViewCollection;
        // ✅ 新增：失败归档节点
        private TreeViewNode _failureNode;
        // ✅ 新增：自动化运行标志
        private bool _isAutoRunning = false;

        // ✅ 新增：F1/F2 热键 ID
        private const int HOTKEY_F1 = 9007;
        private const int HOTKEY_F2 = 9008;
        private const uint VK_F1 = 0x70;
        private const uint VK_F2 = 0x71;

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;




        // MainWindow.xaml.cs

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // 核心：线程挂接 API
        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetCursorPos(int x, int y);

        // Windows API 声明
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        // 激活窗口
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }


        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsIconic(IntPtr hWnd); // 判断窗口是否最小化




        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);



        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;

        // 虚拟键码
        private const uint VK_UP = 0x26;
        private const uint VK_DOWN = 0x28;
        private const uint VK_LEFT = 0x25;
        private const uint VK_RIGHT = 0x27;
        private const uint VK_RETURN = 0x0D;
        private const uint VK_OEM_7 = 0xDE;

        // 修饰键
        private const uint MOD_CONTROL = 0x0002;



        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOPMOST = 0x00000008;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_SHOWWINDOW = 0x0040;

        public MainWindow()
        {
            InitializeComponent();
            _inputSimulator = new InputSimulator();

            LoadBusinessInfo();


            ExcelPackage.License.SetNonCommercialPersonal("fff");

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            if (string.IsNullOrEmpty(baseDir))
            {
                baseDir = Environment.CurrentDirectory;
            }
            _exportDirectory = Path.Combine(baseDir, "ExportedFiles");

            string screenshotBaseDir = Path.Combine(baseDir, "Screenshots");
            _screenshotHelper = new ScreenshotHelper(screenshotBaseDir, (msg) => {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = msg;
                });
            });

            _searchConfig = SearchConfig.Load();
            _searchHelper = new SearchHelper(_searchConfig, (msg) =>
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = msg;
                });
            });

            StoreTreeView.SelectedItemChanged += StoreTreeView_SelectedItemChanged;
            this.Loaded += MainWindow_Loaded;
            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _windowHandle = new WindowInteropHelper(this).Handle;
            _source = HwndSource.FromHwnd(_windowHandle);
            _source.AddHook(HwndHook);
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Dispose();
        }

        public void Dispose()
        {
            UnregisterGlobalHotkeys();
            if (_source != null)
            {
                _source.RemoveHook(HwndHook);
                _source.Dispose();
                _source = null;
            }
        }
 

        private string GetWindowClass(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return string.Empty;
            StringBuilder sb = new StringBuilder(256);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private bool CheckWindowReady(IntPtr targetHwnd, string actionName)
        {
            if (targetHwnd == IntPtr.Zero)
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"❌ [{actionName}] 失败: 窗口句柄丢失");
                return false;
            }

            // 1. 检查是否最小化
            if (IsIconic(targetHwnd))
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"❌ [{actionName}] 失败: 窗口被最小化了！请恢复窗口。");
                // 尝试恢复窗口（可选）
                // ShowWindow(targetHwnd, 9); 
                return false;
            }

            // 2. 检查是否在前台
            IntPtr currentForeground = GetForegroundWindow();
            if (currentForeground != targetHwnd)
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"❌ [{actionName}] 失败: 窗口失去焦点（被遮挡或切换）。");
                // 尝试抢回焦点
                SetForegroundWindow(targetHwnd);
                return false; // 这里建议直接失败，让用户人工干预，或者返回 false 让外层重试
            }

            // 3. 检查坐标是否有效 (防止在屏幕外)
            if (GetWindowRect(targetHwnd, out RECT rect))
            {
                if (rect.Right - rect.Left <= 0 || rect.Bottom - rect.Top <= 0)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"❌ [{actionName}] 失败: 窗口尺寸异常。");
                    return false;
                }
            }

            return true;
        }





        #region 窗口置顶和全局快捷键

        private void AlwaysOnTopCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            this.Topmost = true;
            RegisterGlobalHotkeys();
            StatusTextBlock.Text = "窗口已置顶，全局快捷键已启用";
        }

        private void AlwaysOnTopCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            this.Topmost = false;
            UnregisterGlobalHotkeys();
            StatusTextBlock.Text = "窗口置顶已取消，全局快捷键已禁用";
        }






        /// <summary>
        /// 终极窗口激活：置顶 + 还原 + 线程挂接夺权
        /// </summary>
        private bool RobustActivateWindow(IntPtr targetHwnd)
        {
            if (targetHwnd == IntPtr.Zero) return false;

            // 1. 【视觉层】智能置顶：保证 OCR 不被遮挡
            // 直接复用你刚才写的 EnsureWindowTopMost 方法
            EnsureWindowTopMost(targetHwnd);

            // 2. 【状态层】检查是否最小化，如果是则还原
            if (IsIconic(targetHwnd))
            {
                ShowWindow(targetHwnd, SW_RESTORE);
                System.Threading.Thread.Sleep(200); // 还原动画需要时间
            }
            else
            {
                ShowWindow(targetHwnd, SW_SHOW);
            }

            // 3. 【逻辑层】使用 AttachThreadInput 强行夺取焦点
            uint currentThreadId = GetCurrentThreadId();
            uint targetThreadId = GetWindowThreadProcessId(targetHwnd, out _);

            if (currentThreadId != targetThreadId)
            {
                try
                {
                    // A. 挂接线程：告诉系统我们是一家人
                    AttachThreadInput(currentThreadId, targetThreadId, true);

                    // B. 夺取前景：此时系统不会拦截
                    SetForegroundWindow(targetHwnd);

                    // C. 焦点兜底：有些控件需要显式 Focus
                    // SetFocus(targetHwnd); // 需要引入 API，通常 SetForegroundWindow 够用了

                    // D. 循环确认：确保真的激活了
                    int retries = 0;
                    while (GetForegroundWindow() != targetHwnd && retries < 10)
                    {
                        SetForegroundWindow(targetHwnd);
                        System.Threading.Thread.Sleep(50);
                        retries++;
                    }
                }
                finally
                {
                    // E. 脱钩：操作完必须解绑，否则会卡死
                    AttachThreadInput(currentThreadId, targetThreadId, false);
                }
            }
            else
            {
                // 同线程直接激活
                SetForegroundWindow(targetHwnd);
            }

            return GetForegroundWindow() == targetHwnd;
        }



        /// <summary>
        /// ✅ [智能置顶] 检查窗口状态，仅在未置顶时执行置顶操作
        /// </summary>
        private void EnsureWindowTopMost(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return;

            try
            {
                // 获取窗口当前的扩展样式
                int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);

                // 判断是否已经包含 TOPMOST 属性
                bool isTopMost = (exStyle & WS_EX_TOPMOST) != 0;

                if (!isTopMost)
                {
                    // 只有未置顶时，才执行置顶，避免重复操作
                    SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

                    // 给一点时间让系统反应
                    System.Threading.Thread.Sleep(50);
                    System.Diagnostics.Debug.WriteLine($"[窗口优化] 检测到未置顶，已强制置顶: {hwnd}");
                }
                else
                {
                    // 已经是置顶状态，无需操作，直接返回
                    // System.Diagnostics.Debug.WriteLine($"[窗口优化] 窗口已置顶，跳过设置: {hwnd}");
                }

                // 双重保险：无论是否刚设置过，都请求一次前台激活
                SetForegroundWindow(hwnd);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"置顶操作异常: {ex.Message}");
            }
        }



        private void RegisterGlobalHotkeys()
        {
            if (_globalHotkeysRegistered || _windowHandle == IntPtr.Zero) return;
            try
            {
                // 1. 注册原有的 Ctrl 组合键
                bool upRegistered = RegisterHotKey(_windowHandle, HOTKEY_UP, MOD_CONTROL, VK_UP);
                bool downRegistered = RegisterHotKey(_windowHandle, HOTKEY_DOWN, MOD_CONTROL, VK_DOWN);
                bool leftRegistered = RegisterHotKey(_windowHandle, HOTKEY_LEFT, MOD_CONTROL, VK_LEFT);
                bool rightRegistered = RegisterHotKey(_windowHandle, HOTKEY_RIGHT, MOD_CONTROL, VK_RIGHT);
                bool enterRegistered = RegisterHotKey(_windowHandle, HOTKEY_ENTER, MOD_CONTROL, VK_RETURN);
                bool quoteRegistered = RegisterHotKey(_windowHandle, HOTKEY_QUOTE, MOD_CONTROL, VK_OEM_7);

                // 2. 注册新增的 F1 / F2 (无修饰键)
                bool f1Registered = RegisterHotKey(_windowHandle, HOTKEY_F1, 0, VK_F1);
                bool f2Registered = RegisterHotKey(_windowHandle, HOTKEY_F2, 0, VK_F2);

                if (upRegistered && downRegistered && leftRegistered && rightRegistered && enterRegistered && quoteRegistered && f1Registered && f2Registered)
                {
                    _globalHotkeysRegistered = true;
                    StatusTextBlock.Text = "全局快捷键已启用：F1自动开始，F2停止，Ctrl+方向键操作...";
                }
                else
                {
                    // 如果部分注册失败，清理已注册的，避免残留
                    UnregisterGlobalHotkeys();
                    StatusTextBlock.Text = "全局快捷键注册失败 (可能部分冲突)";
                }
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"快捷键注册错误: {ex.Message}";
            }
        }

        private void UnregisterGlobalHotkeys()
        {
            if (!_globalHotkeysRegistered || _windowHandle == IntPtr.Zero) return;
            try
            {
                // 注销原有 Ctrl 组合键
                UnregisterHotKey(_windowHandle, HOTKEY_UP);
                UnregisterHotKey(_windowHandle, HOTKEY_DOWN);
                UnregisterHotKey(_windowHandle, HOTKEY_LEFT);
                UnregisterHotKey(_windowHandle, HOTKEY_RIGHT);
                UnregisterHotKey(_windowHandle, HOTKEY_ENTER);
                UnregisterHotKey(_windowHandle, HOTKEY_QUOTE);

                // 注销 F1 / F2
                UnregisterHotKey(_windowHandle, HOTKEY_F1);
                UnregisterHotKey(_windowHandle, HOTKEY_F2);

                _globalHotkeysRegistered = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"注销快捷键失败: {ex.Message}");
            }
        }

        // MainWindow.xaml.cs

        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY)
            {
                int id = wParam.ToInt32();
                bool shouldHandle = false;

                if (id == HOTKEY_UP)
                {
                    // 向上导航 (保持原样)
                    Application.Current.Dispatcher.Invoke(() => NavigateTreeView(-1));
                    shouldHandle = true;
                }
                else if (id == HOTKEY_DOWN)
                {
                    // 向下导航 (保持原样)
                    Application.Current.Dispatcher.Invoke(() => NavigateTreeView(1));
                    shouldHandle = true;
                }
                else if (id == HOTKEY_LEFT)
                {
                    // Ctrl+Left: 仅粘贴名称 (保持原样)
                    Application.Current.Dispatcher.Invoke(() => PasteCurrentStoreName());
                    shouldHandle = true;
                }
                else if (id == HOTKEY_RIGHT || id == HOTKEY_QUOTE)
                {
                    // ✅ Ctrl+Right / Ctrl+;: 调用新的独立粘贴流程
                    // (写入剪贴板 -> 盲粘贴 -> 自动发送 -> 后置补全群名)
                    Application.Current.Dispatcher.InvokeAsync(async () =>
                    {
                        await ManualPasteProcessAsync();
                    });
                    shouldHandle = true;
                }
                else if (id == HOTKEY_ENTER)
                {
                    // ✅ Ctrl+Enter: 手动搜索/前进

                    // 关键：强制释放物理按住的 Ctrl 键，防止干扰后续的搜索指令
                    try
                    {
                        _inputSimulator.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                    }
                    catch { }

                    Application.Current.Dispatcher.InvokeAsync(async () =>
                    {
                        // 如果正在自动跑，Enter键作为暂停键
                        if (_isAutoRunning)
                        {
                            StopAutoSending();
                        }
                        else
                        {
                            // 否则执行手动搜索逻辑 (复用自动化核心)
                            await SmartAdvanceOrSearchAsync();
                        }
                    });
                    shouldHandle = true;
                }
                else if (id == HOTKEY_F1)
                {
                    // F1: 启动自动化发送
                    if (!_isAutoRunning)
                    {
                        StartAutoSending();
                    }
                    shouldHandle = true;
                }
                else if (id == HOTKEY_F2)
                {
                    // F2: 停止自动化发送
                    if (_isAutoRunning)
                    {
                        StopAutoSending();
                    }
                    shouldHandle = true;
                }

                handled = shouldHandle;
            }
            return IntPtr.Zero;
        }

        #region 自动化发送控制逻辑 (F1/F2)

        private void StartAutoSending()
        {
            if (_isAutoRunning) return;

            // 强制激活自己，确保 CheckWindowReady 能通过
            this.Activate();
            this.Focus();

            _isAutoRunning = true;
            StatusTextBlock.Text = "🚀 [F1] 自动化发送模式已启动！(按 F2 停止)";

            // 启动后台循环任务
            Task.Run(AutoProcessLoop);
        }

        private void StopAutoSending()
        {
            _isAutoRunning = false;
            StatusTextBlock.Text = "🛑 [F2] 自动化发送已停止。";
        }





        private async Task AutoProcessLoop()
        {
            while (_isAutoRunning)
            {
                bool shouldStop = false;

                // 1. 状态检查与停止条件
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    // A. 焦点自救
                    if (_currentSelectedNode == null && _currentSelectedIndex >= 0 && _currentSelectedIndex < _flatNodeList.Count)
                    {
                        StatusTextBlock.Text = "⚠️ 检测到焦点丢失，正在尝试恢复...";
                        var rescueNode = _flatNodeList[_currentSelectedIndex];
                        FocusAndSelectItem(rescueNode);
                        _currentSelectedNode = rescueNode;
                    }

                    // B. 检查是否到达列表末尾
                    if (_currentSelectedNode == null ||
                        _currentSelectedNode == _failureNode ||
                        _currentSelectedNode.StoreName == "FAIL_SEPARATOR")
                    {
                        StatusTextBlock.Text = "🏁 列表已处理完毕，自动化停止。";
                        shouldStop = true;
                        return;
                    }

                    // 🛑 C. [仅自动模式] 检查是否有群名
                    // 如果没有群名，认为到了未配置区域，自动模式必须停下来，防止乱发
                    if (string.IsNullOrEmpty(_currentSelectedNode.GroupName))
                    {
                        StatusTextBlock.Text = $"🛑 商家 '{_currentSelectedNode.StoreName}' 无群名，自动停止。";
                        FocusAndSelectItem(_currentSelectedNode); // 选中它提示用户
                        shouldStop = true;
                        return;
                    }
                });

                if (shouldStop)
                {
                    _isAutoRunning = false;
                    break;
                }

                // 2. 核心处理
                bool success = await SearchCurrentItemAsync(true);

                if (!_isAutoRunning) break;

                if (success)
                {
                    // ✅ 成功分支
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
                        this.Activate();
                        this.Focus();
                        SetForegroundWindow(_windowHandle);

                        StatusTextBlock.Text += " [成功] 下一条...";
                        NavigateTreeView(1);
                    });
                    
                }
                else
                {
                    // ❌ 失败分支
                    await Application.Current.Dispatcher.InvokeAsync(() => StatusTextBlock.Text += " [失败] 移入重试区...");
                   

                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
                        this.Activate();
                        this.Focus();
                        SetForegroundWindow(_windowHandle);
                        MoveCurrentToFailureNode();
                    });
                    await Task.Delay(500);
                }
            }

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                if (!_isAutoRunning) StatusTextBlock.Text += " (已停止)";
            });
        }


        /// <summary>
        /// ✅ [修复版] 将失败项移动到列表最末尾（同级），并选中下一项
        /// </summary>
        private void MoveCurrentToFailureNode()
        {
            // 此方法必须在 UI 线程调用
            var node = _currentSelectedNode;

            // 如果当前没有选中，或者选中的是失败归档节点本身，则不处理
            if (node == null || node == _failureNode || node.StoreName == "FAIL_SEPARATOR") return;

            // 1. 从主列表 (_treeViewCollection) 中移除当前项
            if (_treeViewCollection.Contains(node))
            {
                _treeViewCollection.Remove(node);

                // 3. ✅ 【修改点】不再添加到 Children，而是添加到主列表的最末尾
                _treeViewCollection.Add(node);
            }

            // 4. 重建扁平列表 (因为顺序变了，必须重新生成索引)
            RebuildFlatNodeList();

            // 5. 修正索引：防止越界
            // 因为刚刚移除了一个元素，当前位置的元素索引可能变了，或者后面没有元素了
            if (_currentSelectedIndex >= _flatNodeList.Count)
            {
                _currentSelectedIndex = _flatNodeList.Count - 1;
            }

            // 如果列表空了（全移完了），归位
            if (_flatNodeList.Count == 0)
            {
                _currentSelectedIndex = -1;
            }

            // =========================================================
            // 🛑 防止死循环核心：抢回焦点 并 指向替补上来的新节点
            // =========================================================

            // A. 抢回窗口焦点
            if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
            this.Activate();
            this.Focus();
            SetForegroundWindow(_windowHandle);

            // B. 更新引用
            // 注意：我们要找的不是刚刚移到末尾的那个 node，而是原本位置替补上来的 nextNode
            // 此时 _currentSelectedIndex 指向的位置就是替补上位的新节点（因为旧的被移走了）
            if (_currentSelectedIndex >= 0 && _currentSelectedIndex < _flatNodeList.Count)
            {
                var nextNode = _flatNodeList[_currentSelectedIndex];

                // 如果当前索引指向的正好是我们刚刚移到末尾的那个节点（说明已经循环一圈了，或者后面没得选了）
                if (nextNode == node)
                {
                    // 尝试找列表里的第一个“非重试”节点，或者干脆停止
                    // 这里简单处理：如果下一项就是刚刚移走的自己，说明列表只有这一个了，或者都处理完了
                    _currentSelectedNode = nextNode;
                }
                else
                {
                    // 选中新的替补节点
                    FocusAndSelectItem(nextNode);
                    _currentSelectedNode = nextNode; // 👈 更新指针，让下一轮循环处理新的人
                }
            }
            else
            {
                _currentSelectedNode = null;
            }
        }

        #endregion




        // MainWindow.xaml.cs

        // 增加一个静态锁对象 (或者使用类成员变量)
        private static readonly SemaphoreSlim _manualLock = new SemaphoreSlim(1, 1);

        private async Task SmartAdvanceOrSearchAsync()
        {
            // 1. 🔒 立即尝试获取锁，如果正在处理中，直接忽略本次按键
            if (!await _manualLock.WaitAsync(0))
            {
                System.Diagnostics.Debug.WriteLine("⚠️ 操作太快，忽略本次 Ctrl+Enter");
                return;
            }

            try
            {
                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    // ... 这里的逻辑保持不变 ...
                    if (_currentSelectedNode == null || string.IsNullOrEmpty(_currentSelectedNode.StoreName))
                    {
                        StatusTextBlock.Text = "⚠️ 请先选择一个商家";
                        return;
                    }

                    // 智能前进判断逻辑 (保持不变)
                    string currentStoreName = _currentSelectedNode.StoreName;
                    if (_currentItemPasted && _lastPastedStoreName == currentStoreName)
                    {
                        StatusTextBlock.Text = "⏭️ [手动] 前进到下一项...";
                        await Task.Delay(50);
                        NavigateTreeView(1);

                        if (_currentSelectedNode == null || string.IsNullOrEmpty(_currentSelectedNode.StoreName))
                        {
                            StatusTextBlock.Text = "✅ 列表到底了！";
                            return;
                        }
                        _currentItemPasted = false;
                        _lastPastedStoreName = null;
                        await Task.Delay(100);
                    }
                    else
                    {
                        StatusTextBlock.Text = "▶️ [手动] 启动处理...";
                        _currentItemPasted = false;
                        _lastPastedStoreName = null;
                    }

                    // 调用核心方法
                    await ManualSmartProcessAsync();
                });
            }
            finally
            {
                // 🔓 释放锁，允许下一次操作
                _manualLock.Release();
            }
        }



        private async Task ManualSmartProcessAsync()
        {
            // 1. 基础检查
            if (_currentSelectedNode == null) return;

            // 2. 🛡️ 净化环境：释放按键
            // 这是手动模式成功的关键，防止物理按键（Ctrl/Enter）干扰自动化的 SearchCurrentItemAsync
            try
            {
                _inputSimulator.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                _inputSimulator.Keyboard.KeyUp(VirtualKeyCode.RETURN);
            }
            catch { }


            // 3. ♻️ 直接复用自动化的核心逻辑！
            // 核心优势：F1 怎么跑，这里就怎么跑。包含完整的：
            // 找窗口 -> 激活 -> 搜索(含防抖) -> OCR验证列表 -> 回车 -> OCR验证标题 -> 粘贴
            bool success = await SearchCurrentItemAsync(false); // isAutoMode = false

            // 4. 轮询逻辑后置处理
            // 如果没有群名（处于轮询模式），且本次搜索结束（无论成败），
            // 都切换一下轮次，以便下次 Ctrl+Enter 搜另一个 APP
            if (string.IsNullOrEmpty(_currentSelectedNode.GroupName))
            {
                _isWeworkTurn = !_isWeworkTurn;

                // 更新 UI 提示下一次搜谁
                string nextApp = _isWeworkTurn ? "企业微信" : "微信";
                StatusTextBlock.Text = $"⏳ 下次轮询: {nextApp}";
            }
        }

        // MainWindow.xaml.cs

        /// <summary>
        /// ✅ [粘贴键专属] 动态识别版
        /// 逻辑：识别窗口身份(读配置) -> 写入剪贴板 -> 盲粘贴 -> (自动发送) -> [无群名则补全]
        /// </summary>
        private async Task ManualPasteProcessAsync()
        {
            // 1. 基础检查
            if (_currentSelectedNode == null)
            {
                StatusTextBlock.Text = "⚠️ 请先选择一个商家";
                return;
            }

            string storeName = _currentSelectedNode.StoreName;
            string originalGroupName = _currentSelectedNode.GroupName;
            bool isFileNode = _currentSelectedNode.IsFileNode;

            // ============================================================
            // 2. 🤖 自动识别窗口身份 (兼容 search_config.json)
            // ============================================================
            IntPtr currentHwnd = GetForegroundWindow();
            string currentClassName = GetWindowClass(currentHwnd);

            // 默认为微信 (false)
            bool isWework = false;

            // 动态对比配置中的类名
            if (currentClassName == _searchConfig.WeworkWindowClassName)
            {
                isWework = true;
                StatusTextBlock.Text = "🤖 检测到：企业微信";
            }
            else if (currentClassName == _searchConfig.WechatWindowClassName)
            {
                isWework = false;
                StatusTextBlock.Text = "🤖 检测到：微信";
            }
            else
            {
                // 兜底：如果类名既不像微信也不像企微，尝试用模糊匹配兜底
                // (防止配置填错导致完全无法识别)
                if (currentClassName.Contains("WeWork") || currentClassName.Contains("WXWork"))
                {
                    isWework = true;
                    StatusTextBlock.Text = "🤖 检测到：企业微信 (模糊匹配)";
                }
                else
                {
                    // 实在认不出来，就沿用当前轮询的状态或者默认为微信
                    // 这里选择不做改变，仅提示
                    // StatusTextBlock.Text = $"⚠️ 未知窗口类名: {currentClassName}";
                }
            }

            // ============================================================
            // 3. 执行核心动作 (传入识别到的 isWework)
            // ============================================================
            bool actionSuccess = false;
            if (isFileNode)
            {
                // 传入 isWework 确保点击坐标正确 (企微宽，微信窄)
                actionSuccess = await PasteExcelFileAsync(storeName, isWework);
            }
            else
            {
                actionSuccess = await PasteFullStoreInfoAsync(storeName, isWework);
            }

            if (!actionSuccess) return;

            // ============================================================
            // 4. 后置智能补全 (仅当原先没有群名时触发)
            // ============================================================
            if (string.IsNullOrEmpty(originalGroupName))
            {
                StatusTextBlock.Text = "👁️ 正在识别群名以补全...";
                try
                {
                    // 等待发送动画
                    await Task.Delay(200);

                    // OCR 获取标题 (传入刚才识别到的身份 isWework)
                    string recognizedTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(currentHwnd, isWework);

                    // 如果第一次没识别到，尝试反向识别一次 (以防万一)
                    if (string.IsNullOrWhiteSpace(recognizedTitle))
                    {
                        recognizedTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(currentHwnd, !isWework);
                    }

                    if (!string.IsNullOrWhiteSpace(recognizedTitle) && recognizedTitle.Length > 1)
                    {
                        // ✅ 识别成功：保存并更新
                        string sourceToSave = isWework ? "企业微信" : "微信";

                        UpdateBusInfo(storeName, recognizedTitle, sourceToSave);
                        StatusTextBlock.Text = $"✅ [补全] 已保存为[{sourceToSave}]: {recognizedTitle}";
                    }
                    else
                    {
                        StatusTextBlock.Text = "⚠️ 未识别到有效标题，跳过补全";
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"补全失败: {ex.Message}");
                }
            }
        }


        private async Task<bool> PasteFullStoreInfoBlindAsync(string storeName)
        {
            // 1. 准备数据
            List<string> trackingNumbers;
            lock (_dataLock)
            {
                if (!_storeData.TryGetValue(storeName, out trackingNumbers))
                {
                    StatusTextBlock.Text = "❌ 未找到商家数据";
                    return false;
                }
                trackingNumbers = trackingNumbers.ToList();
            }

            var sb = new StringBuilder();
            sb.AppendLine(storeName);
            foreach (var num in trackingNumbers) sb.AppendLine(num);
            sb.AppendLine(FIXED_MESSAGE);

            // 2. 核心：主动写入剪贴板 (这是“粘贴商家信息”的关键)
            if (!await SetClipboardWithRetryAsync(sb.ToString()))
            {
                StatusTextBlock.Text = "❌ 剪贴板被占用";
                return false;
            }

            // 稍作等待确保剪贴板生效
            await Task.Delay(50);

            // 3. 盲粘贴 (不移动鼠标，直接 Ctrl+V)
            SimulatePaste();

            // 4. 处理自动发送
            if (AutoSendCheckBox.IsChecked == true)
            {
                // 稍等渲染
                await Task.Delay(200);

                // 发送动作
                SimulateAltS();
                await Task.Delay(50);
                SimulateEnter(); // 补刀

                StatusTextBlock.Text = $"✅ [快捷] 已发送: {storeName}";
            }
            else
            {
                StatusTextBlock.Text = $"📋 [快捷] 已粘贴: {storeName}";
            }

            // 更新状态，方便 Ctrl+Enter 跳转
            _currentItemPasted = true;
            _lastPastedStoreName = storeName;

            return true;
        }
        // MainWindow.xaml.cs

        // MainWindow.xaml.cs

        /// <summary>
        /// ✅ [新增辅助方法] 统一处理：更新内存 -> 保存文件 -> 刷新界面
        /// </summary>
        private void UpdateBusInfo(string storeName, string newGroupName, string source)
        {
            // 1. 更新内存列表 (_businessInfoList)
            var info = _businessInfoList.FirstOrDefault(b => b.StoreName == storeName);
            if (info == null)
            {
                info = new BusinessInfo { StoreName = storeName };
                _businessInfoList.Add(info);
            }

            // 更新属性
            info.GroupName = newGroupName;
            info.Source = source;

            // 2. 保存到本地 JSON 文件
            // (调用你原有的 SaveBusinessInfo 方法)
            SaveBusinessInfo();

            // 3. 刷新 TreeView 界面显示
            // (调用你原有的 UpdateNodeGroupInfo 方法)
            UpdateNodeGroupInfo(storeName, newGroupName, source);
        }

        private async Task<bool> PasteExcelFileBlindAsync(string storeName)
        {
            // 1. 准备文件路径
            string filePath;
            lock (_dataLock) { if (!_exportedFilePaths.TryGetValue(storeName, out filePath)) return false; }

            if (!File.Exists(filePath))
            {
                StatusTextBlock.Text = "❌ 文件不存在";
                return false;
            }

            // 2. 核心：主动写入剪贴板
            bool clipboardSuccess = await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    var data = new DataObject();
                    data.SetData(DataFormats.FileDrop, new string[] { filePath });
                    Clipboard.SetDataObject(data, true);
                    return true;
                }
                catch { return false; }
            });

            if (!clipboardSuccess)
            {
                StatusTextBlock.Text = "❌ 文件剪贴板写入失败";
                return false;
            }
            await Task.Delay(50);

            // 3. 盲粘贴
            SimulatePaste();

            // 4. 处理自动发送
            if (AutoSendCheckBox.IsChecked == true)
            {
                StatusTextBlock.Text = "🚀 正在发送文件...";
                await Task.Delay(500); // 文件加载稍慢
                SimulateAltS();
                await Task.Delay(50);
                SimulateEnter();
                StatusTextBlock.Text = $"✅ [快捷] 文件已发送: {storeName}";
            }
            else
            {
                StatusTextBlock.Text = $"📋 [快捷] 文件已粘贴: {storeName}";
            }

            _currentItemPasted = true;
            _lastPastedStoreName = storeName;

            return true;
        }

        /// <summary>
        /// 🔥 专门为搜索操作设计的窗口激活方法
        /// </summary>
        private async Task<bool> ActivateWindowForSearchAsync(IntPtr targetHwnd, bool isWework)
        {
            const int MAX_ATTEMPTS = 5;

            for (int attempt = 1; attempt <= MAX_ATTEMPTS; attempt++)
            {
                System.Diagnostics.Debug.WriteLine($"[激活] 第 {attempt} 次尝试...");

                // 步骤1：基础激活
                RobustActivateWindow(targetHwnd);
                await Task.Delay(100);

                // 步骤2：检查是否真的激活了
                if (GetForegroundWindow() != targetHwnd)
                {
                    System.Diagnostics.Debug.WriteLine($"[激活] SetForegroundWindow 未生效，尝试点击激活");

                    // 点击窗口标题栏区域强制激活
                    if (GetWindowRect(targetHwnd, out RECT rect))
                    {
                        int clickX = (rect.Left + rect.Right) / 2;
                        int clickY = rect.Top + 30; // 标题栏位置

                        SetCursorPos(clickX, clickY);
                        await Task.Delay(30);
                        mouse_event(MOUSEEVENTF_LEFTDOWN, clickX, clickY, 0, 0);
                        await Task.Delay(30);
                        mouse_event(MOUSEEVENTF_LEFTUP, clickX, clickY, 0, 0);
                        await Task.Delay(150);
                    }
                }

                // 步骤3：企业微信专用 - 额外点击主内容区确保焦点到位
                if (isWework && GetForegroundWindow() == targetHwnd)
                {
                    System.Diagnostics.Debug.WriteLine($"[激活] 企微专用：点击内容区域获取内部焦点");

                    if (GetWindowRect(targetHwnd, out RECT rect))
                    {
                        // 点击窗口中央偏左的位置（通常是聊天列表区域）
                        int contentX = rect.Left + 150;
                        int contentY = (rect.Top + rect.Bottom) / 2;

                        SetCursorPos(contentX, contentY);
                        await Task.Delay(30);
                        mouse_event(MOUSEEVENTF_LEFTDOWN, contentX, contentY, 0, 0);
                        await Task.Delay(30);
                        mouse_event(MOUSEEVENTF_LEFTUP, contentX, contentY, 0, 0);

                        // 🔥 关键：企微需要更长的焦点稳定时间
                        await Task.Delay(300);
                    }
                }

                // 步骤4：最终验证
                if (GetForegroundWindow() == targetHwnd)
                {
                    System.Diagnostics.Debug.WriteLine($"[激活] ✅ 第 {attempt} 次尝试成功");

                    // 企微额外等待，确保内部状态就绪
                    if (isWework)
                    {
                        await Task.Delay(200);
                    }

                    return true;
                }

                await Task.Delay(100);
            }

            System.Diagnostics.Debug.WriteLine($"[激活] ❌ {MAX_ATTEMPTS} 次尝试均失败");
            return false;
        }






        // MainWindow.xaml.cs

        /// <summary>
        /// ✅ [自动模式核心] 乱序匹配增强版 + 详细日志 + 强力激活修复
        /// </summary>
        private async Task<bool> SearchCurrentItemAsync(bool isAutoMode = false)
        {
            // 🔍 [调试日志] 流程开始
            System.Diagnostics.Debug.WriteLine($"\n[{DateTime.Now:HH:mm:ss.fff}] ============== [调试] 开始搜索流程 ==============");

            // 1. 获取数据快照
            string storeName = null;
            string groupName = null;
            string source = null;
            bool isFileNode = false;

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                if (_currentSelectedNode != null)
                {
                    storeName = _currentSelectedNode.StoreName;
                    groupName = _currentSelectedNode.GroupName;
                    source = _currentSelectedNode.Source;
                    isFileNode = _currentSelectedNode.IsFileNode;
                }
            });

            if (string.IsNullOrEmpty(storeName))
            {
                System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [错误] 商家名为空，流程中止。");
                return false;
            }

            var snapshot = new
            {
                StoreName = storeName,
                GroupName = groupName?.Trim(),
                SearchText = !string.IsNullOrEmpty(groupName) ? groupName.Trim() : storeName.Trim(),
                HasGroupName = !string.IsNullOrEmpty(groupName),
                // 判断目标是企微还是微信
                IsWework = !string.IsNullOrEmpty(groupName) ? "企业微信".Equals(source) : _isWeworkTurn
            };

            string appName = snapshot.IsWework ? "企业微信" : "微信";

            System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [快照] 目标应用:{appName}, 搜索词:{snapshot.SearchText}, 有群名:{snapshot.HasGroupName}, 上次群名:{_lastEnteredGroupName ?? "无"}");
            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"🔍 正在 [{appName}] 搜索: {snapshot.SearchText}...");

            // ✅ 定义执行粘贴的动作 (这里修复了参数缺失报错)
            Func<Task<bool>> performPasteAsync = async () =>
            {
                System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [动作] 准备执行粘贴/发送任务...");

                // 关键修复：传入 snapshot.IsWework 参数
                if (isFileNode)
                    return await PasteExcelFileAsync(snapshot.StoreName, snapshot.IsWework);
                else
                    return await PasteFullStoreInfoAsync(snapshot.StoreName, snapshot.IsWework);
            };

            try
            {
                // --------------------------------------------------------
                // 🚀 1. 极速模式
                // --------------------------------------------------------
                if (snapshot.HasGroupName && snapshot.SearchText == _lastEnteredGroupName)
                {
                    System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [极速] 命中同名群条件，开始视觉验证...");
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "👀 [极速] 检测同名群...");

                    bool activateResult = false;
                    if (_lastChatWindowHandle != IntPtr.Zero)
                    {
                        activateResult = RobustActivateWindow(_lastChatWindowHandle);
                    }
                    else
                    {
                        activateResult = RobustActivateWindow(GetForegroundWindow());
                    }

                    await Task.Delay(100);

                    IntPtr checkHwnd = GetForegroundWindow();
                    string titleText = await _screenshotHelper.GetWeChatWindowTitleTextAsync(checkHwnd, snapshot.IsWework);

                    if (_screenshotHelper.IsFuzzyMatch(snapshot.SearchText, titleText))
                    {
                        System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [极速] ✅ 验证通过，跳过搜索步骤。");
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"⚡ [极速] 验证通过，直接发送。");
                        return await performPasteAsync();
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [极速] ❌ 验证失败 (OCR不符)，转入常规搜索。");
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "⚠️ 窗口不符，转常规搜索...");
                        _lastEnteredGroupName = null;
                        _lastChatWindowHandle = IntPtr.Zero;
                    }
                }

                // ============================================================
                // 🔍 2. 常规搜索模式
                // ============================================================

                IntPtr mainHwnd = GetForegroundWindow();
                if (!RobustActivateWindow(mainHwnd))
                {
                    return false;
                }

                // 搜索 (SearchHelper 已包含空格退格修复)
                bool autoSearchSuccess = await Task.Run(() => _searchHelper.SearchInApp(snapshot.SearchText, snapshot.IsWework));
                if (!autoSearchSuccess)
                {
                    return false;
                }

                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "👀 [自动] 验证搜索列表...");
                await Task.Delay(200);

                // --------------------------------------------------------
                // 🔥 步骤 A: 搜索列表 OCR 验证
                // --------------------------------------------------------
                IntPtr searchHwnd = GetForegroundWindow();
                bool isListMatch = false;

                for (int i = 0; i < 3; i++)
                {
                    isListMatch = await _screenshotHelper.CheckSearchResultAsync(searchHwnd, snapshot.SearchText, snapshot.IsWework);
                    if (isListMatch) break;
                    if (i < 2) await Task.Delay(800);
                }

                if (!isListMatch)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ 搜索列表超时或未找到目标，停止。");
                    return false;
                }

                // --------------------------------------------------------
                // 🔥 步骤 B: 直接回车
                // --------------------------------------------------------
                RobustActivateWindow(searchHwnd);
                await Task.Delay(50);
                _inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                await Task.Delay(100);

                // --------------------------------------------------------
                // 🔥 步骤 C: 进群验证
                // --------------------------------------------------------

                IntPtr chatHwnd = GetForegroundWindow();
                // 简单检查窗口
                if (chatHwnd == IntPtr.Zero) return false;

                bool enteredSuccess = false;
                string cleanTarget = snapshot.SearchText.Replace(" ", "").ToLower();

                for (int i = 0; i < 6; i++)
                {
                    string rawTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(chatHwnd, snapshot.IsWework);
                    string cleanTitle = System.Text.RegularExpressions.Regex.Replace(rawTitle, @"\(\d+.*?\)|（\d+.*?）|\(外部\)|（外部）|\s+", "").ToLower();

                    bool isMatch = false;

                    if (_screenshotHelper.IsFuzzyMatch(snapshot.SearchText, rawTitle)) isMatch = true;
                    else if (cleanTitle.Contains(cleanTarget) || (cleanTarget.Contains(cleanTitle) && cleanTitle.Length > 2)) isMatch = true;
                    else
                    {
                        int matchCount = 0;
                        foreach (char c in cleanTarget) if (cleanTitle.Contains(c)) matchCount++;
                        double overlapRate = cleanTarget.Length > 0 ? (double)matchCount / cleanTarget.Length : 0;
                        if (overlapRate > 0.8) isMatch = true;
                    }

                    if (isMatch)
                    {
                        enteredSuccess = true;
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"✅ [自动] 确认进入: {rawTitle}");
                        if (snapshot.HasGroupName)
                        {
                            _lastEnteredGroupName = snapshot.SearchText;
                            _lastChatWindowHandle = chatHwnd;
                        }
                        break;
                    }
                    await Task.Delay(200);
                }

                if (enteredSuccess)
                {
                    return await performPasteAsync();
                }
                else
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"❌ 标题不符，停止。");
                    _lastEnteredGroupName = null;
                    return false;
                }
            }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"💥 流程异常: {ex.Message}");
                return false;
            }
        }


        private async Task<bool> PasteAndVerifySendAsync(string contentToSend, bool isFile)
        {
            IntPtr targetHwnd = GetForegroundWindow();

            if (!RobustActivateWindow(targetHwnd))
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ 窗口无法激活，发送中止");
                return false;
            }


            // 获取窗口坐标
            if (!GetWindowRect(targetHwnd, out RECT rect)) return false;

            bool isWework = false;
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                if (_currentSelectedNode != null) isWework = "企业微信".Equals(_currentSelectedNode.Source);
            });

            // 计算输入框点击位置
            int xOffset = 270 + 30 + (isWework ? 70 : 0);
            int clickX = rect.Left + xOffset;
            int clickY = rect.Bottom - 70;

            // === 动作 A: 激活输入框并粘贴 ===
            SetCursorPos(clickX, clickY);
            await Task.Delay(30);
            mouse_event(MOUSEEVENTF_LEFTDOWN, clickX, clickY, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, clickX, clickY, 0, 0);
            await Task.Delay(50);

            // 2. 粘贴
            SimulatePaste();

            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "⏳ 粘贴中... (等待渲染)");
     

            // 再次检查窗口（防止粘贴期间窗口关了）
            if (!CheckWindowReady(targetHwnd, "验证粘贴")) return false;

            // --- 提取关键词用于验证 ---
            string keyword;
            if (isFile) keyword = Path.GetFileName(contentToSend);
            else
            {
                if (contentToSend.Contains("未发货预警")) keyword = "未发货预警";
                else if (contentToSend.Contains("考核处罚")) keyword = "考核处罚";
                else keyword = contentToSend.Length > 8 ? contentToSend.Substring(0, 8) : contentToSend;
            }

            // 3. 【验证A】粘贴快照
            var resultPaste = await _screenshotHelper.CaptureSplitVerificationAsync(targetHwnd, isWework);

            if (resultPaste.bottomText == null)
            {
                // 这里如果是截图失败，可能是窗口被挡住了，不判死刑，尝试盲发
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "⚠️ 无法截屏，尝试盲发...");
            }
            else
            {
                bool pasteHadText = _screenshotHelper.IsTextMatch(resultPaste.bottomText, keyword);
                if (!pasteHadText)
                {
                    Debug.WriteLine($"[警告] 粘贴检测未通过。OCR结果: {resultPaste.bottomText}");
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "⚠️ 粘贴文字模糊，尝试强行发送...");
                }
            }

            // === 动作 B: 执行发送 (双保险) ===

            // 再次点击输入框
            mouse_event(MOUSEEVENTF_LEFTDOWN, clickX, clickY, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, clickX, clickY, 0, 0);
            await Task.Delay(50);

            // 方案 1: Alt + S
            SimulateAltS();
            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "✉️ 发送指令 (Alt+S)...");
            await Task.Delay(300);

            // 方案 2: Enter (补刀)
            SimulateEnter();
            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text += " + (Enter补刀)...");

            // ============================================================
            // 🌟 轮询验证结果 (宽松版)
            // ============================================================
            string debugInfo = "";

            for (int i = 0; i < 8; i++)
            {
                await Task.Delay(200);

                // ✅ 【关键修复】这里不再调用严格的 CheckWindowReady
                // 而是手动检查焦点。如果焦点丢了，视为用户切换了窗口，默认判定为“发送成功”
                IntPtr currentForeground = GetForegroundWindow();
                if (currentForeground != targetHwnd)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "⚠️ 验证期间窗口失焦(视为发送成功)。");
                    return true; // 👈 只要动作做完了，窗口丢了也算成功，不移除列表项
                }

                // 截图验证
                var resultSend = await _screenshotHelper.CaptureSplitVerificationAsync(targetHwnd, isWework);
                if (resultSend.topText == null || resultSend.bottomText == null) continue;

                // 判据
                bool inputCleared = !_screenshotHelper.IsTextMatch(resultSend.bottomText, keyword);
                bool messageAppeared = _screenshotHelper.IsTextMatch(resultSend.topText, keyword);

                if (i == 7)
                {
                    string top = resultSend.topText?.Replace("\n", "") ?? "";
                    if (top.Length > 10) top = top.Substring(0, 10);
                    debugInfo = $"Top:{top}.. / Clr:{inputCleared}";
                }

                // ✅ 成功情况 1: 上屏
                if (messageAppeared)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "✅ [上屏验证] 发送成功。");
                    return true;
                }

                // ✅ 成功情况 2: 清空
                if (inputCleared)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "✅ [清空验证] 发送成功。");
                    return true;
                }

                // ⚠️ 补刀重试
                if (!inputCleared && i == 3)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text += " (再次重试Enter)...");
                    mouse_event(MOUSEEVENTF_LEFTDOWN, clickX, clickY, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTUP, clickX, clickY, 0, 0);
                    await Task.Delay(50);
                    SimulateEnter();
                }
            }

            // ✅ 兜底: 流程走完了但OCR没验证到。
            // 为了防止“已发送但被移除”的悲剧，这里返回 TRUE (或者你可以选择返回 false 但不移除，目前为了体验建议返回 true)
            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"⚠️ 验证超时(OCR未确认)，视为潜在成功。");
            return true;
        }




        private void SimulateEnter()
        {
            try
            {
                // 模拟按下 Enter 键
                _inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            }
            catch (Exception ex)
            {
                // 即使报错也不要崩溃
                Debug.WriteLine($"模拟 Enter 失败: {ex.Message}");
            }
        }


        private void NavigateTreeView(int direction)
        {
            // 确保列表是最新的
            if (_flatNodeList.Count == 0) RebuildFlatNodeList();
            if (_flatNodeList.Count == 0) return;

            // 计算新索引
            int newIndex = _currentSelectedIndex + direction;

            // 边界检查
            if (newIndex < 0) newIndex = 0;
            else if (newIndex >= _flatNodeList.Count) newIndex = _flatNodeList.Count - 1;

            _currentSelectedIndex = newIndex;
            var selectedNode = _flatNodeList[_currentSelectedIndex];

            // 执行选中
            FocusAndSelectItem(selectedNode);
        }

        /// <summary>
        /// ✅ 修复：搜索当前选中的商家（防止空引用）
        /// </summary>
        private void SearchCurrentStore()
        {
            if (Interlocked.CompareExchange(ref _searchInProgress, 1, 0) == 1)
            {
                StatusTextBlock.Text = "🏃‍ 操作太快了，请稍候...";
                return;
            }

            try
            {
                if (_currentSelectedNode == null || string.IsNullOrEmpty(_currentSelectedNode.StoreName))
                {
                    StatusTextBlock.Text = "列表为空或未选择商家";
                    Interlocked.Exchange(ref _searchInProgress, 0);
                    return;
                }

                // ✅ 关键修复：在异步操作前捕获当前节点的快照
                string searchText;
                bool isWeworkSearch;
                bool hasGroupName;
                string storeName = _currentSelectedNode.StoreName;

                if (!string.IsNullOrEmpty(_currentSelectedNode.GroupName))
                {
                    searchText = _currentSelectedNode.GroupName;
                    hasGroupName = true;
                    isWeworkSearch = "企业微信".Equals(_currentSelectedNode.Source, StringComparison.OrdinalIgnoreCase);
                    string appName = isWeworkSearch ? "企业微信" : "微信";
                    StatusTextBlock.Text = $"⏳ [群聊优先] 正在 [{appName}] 中搜索: {searchText}...";
                }
                else
                {
                    searchText = _currentSelectedNode.StoreName;
                    hasGroupName = false;
                    isWeworkSearch = _isWeworkTurn;
                    string appName = isWeworkSearch ? "企业微信" : "微信";
                    StatusTextBlock.Text = $"⏳ 正在 [{appName}] 中搜索: {searchText}...";
                }

                // ✅ 使用局部变量，避免依赖可能改变的成员变量
                Task.Run(async () =>
                {
                    try
                    {
                        // 1. 先设置剪贴板为搜索关键词
                        if (!await SetClipboardWithRetryAsync(searchText))
                        {
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                StatusTextBlock.Text = $"❌ 无法复制搜索关键词到剪贴板";
                                Interlocked.Exchange(ref _searchInProgress, 0);
                            });
                            return;
                        }

                        // 2. 等待剪贴板稳定
                        await Task.Delay(100);

                        // 3. 执行搜索操作
                        bool success = _searchHelper.SearchInApp(searchText, isWeworkSearch);

                        // 4. 更新UI状态
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            try
                            {
                                if (success)
                                {
                                    // ✅ 使用捕获的局部变量而不是成员变量
                                    if (!hasGroupName)
                                    {
                                        _isWeworkTurn = !isWeworkSearch;
                                    }
                                    StatusTextBlock.Text = $"✅ 已在目标应用中搜索 '{searchText}'。";
                                }
                                else
                                {
                                    StatusTextBlock.Text = $"❌ 搜索 '{searchText}' 失败。";
                                }
                            }
                            finally
                            {
                                Interlocked.Exchange(ref _searchInProgress, 0);
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            StatusTextBlock.Text = $"💥 搜索时发生错误: {ex.Message}";
                            Interlocked.Exchange(ref _searchInProgress, 0);
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"💥 搜索时发生严重错误: {ex.Message}";
                Interlocked.Exchange(ref _searchInProgress, 0);
            }
        }

        /// <summary>
        /// ✅ 修复：自动前进到下一个商家并搜索（防止空引用）
        /// </summary>

        /// <summary>
        /// ✅ 修复：前进到下一项并自动搜索
        /// </summary>

        /// <summary>
        /// ✅ 保留但简化（现在主要使用 SmartAdvanceOrSearchAsync）
        /// </summary>
        private async Task AdvanceToNextAndSearchAsync()
        {
            // 直接调用智能方法
            await SmartAdvanceOrSearchAsync();
        }





        private void ResetSearchState()
        {
            _isWeworkTurn = true;
        }

        private void RebuildFlatNodeList()
        {
            _flatNodeList.Clear();
            // ❌ 原代码: if (StoreTreeView.ItemsSource is List<TreeViewNode> nodes)
            // ✅ 修复: 改用 IEnumerable 或 IList 来兼容 ObservableCollection
            if (StoreTreeView.ItemsSource is IEnumerable<TreeViewNode> nodes)
            {
                foreach (var node in nodes)
                {
                    if (!string.IsNullOrEmpty(node.StoreName))
                    {
                        _flatNodeList.Add(node);
                    }
                }
            }
            if (_currentSelectedIndex < 0 && _flatNodeList.Count > 0)
            {
                _currentSelectedIndex = 0;
            }
        }




        private void TriggerCopyOperation(TreeViewNode node)
        {
            if (string.IsNullOrEmpty(node.StoreName)) return;

            _currentSelectedNode = node;
            ResetSearchState();

            if (Interlocked.CompareExchange(ref _copyingFlag, 1, 0) == 1) return;

            if (node.IsFileNode)
            {
                CopyStoreNameOnly(node.StoreName);
            }
            else
            {
                CopyFullStoreInfoToClipboard(node.StoreName);
            }
        }

        private void PasteCurrentStoreName()
        {
            if (_currentSelectedNode == null || string.IsNullOrEmpty(_currentSelectedNode.StoreName))
            {
                StatusTextBlock.Text = "请先选择一个商家";
                return;
            }
            Task.Run(async () =>
            {
                try
                {
                    string storeName = _currentSelectedNode.StoreName;
                    if (await SetClipboardWithRetryAsync(storeName))
                    {

                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            SimulatePaste();
                            StatusTextBlock.Text = $"已粘贴商家名称: '{storeName}'";
                        });
                    }
                    else
                    {
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "无法复制商家名称到剪贴板");
                    }
                }
                catch (Exception ex)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"粘贴失败: {ex.Message}");
                }
            });
        }

        private void PasteCurrentStoreFullInfo()
        {
            if (_currentSelectedNode == null || string.IsNullOrEmpty(_currentSelectedNode.StoreName))
            {
                StatusTextBlock.Text = "请先选择一个商家";
                return;
            }
            string storeName = _currentSelectedNode.StoreName;
            if (_currentSelectedNode.IsFileNode)
            {
                PasteExcelFile(storeName);
            }
            else
            {
                PasteFullStoreInfo(storeName);
            }
        }




        // MainWindow.xaml.cs

        private async Task<bool> PasteExcelFileAsync(string storeName, bool isWework)
        {
            Action<string> Log = (msg) => System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [粘贴文件] {msg}");
            Log($"开始处理文件: {storeName}, isWework: {isWework}");

            if (string.IsNullOrEmpty(storeName)) return false;

            IntPtr targetHwnd = GetForegroundWindow();
            Log($"当前窗口句柄: {targetHwnd}");

            string filePath;
            lock (_dataLock) { if (!_exportedFilePaths.TryGetValue(storeName, out filePath)) return false; }
            if (!File.Exists(filePath))
            {
                Log($"❌ 文件不存在: {filePath}");
                return false;
            }

            // 剪贴板
            Log("设置文件到剪贴板...");
            bool clipboardSuccess = await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    var data = new DataObject();
                    data.SetData(DataFormats.FileDrop, new string[] { filePath });
                    Clipboard.SetDataObject(data, true);
                    return true;
                }
                catch (Exception ex)
                {
                    Log($"剪贴板异常: {ex.Message}");
                    return false;
                }
            });

            if (!clipboardSuccess)
            {
                Log("❌ 剪贴板设置失败");
                return false;
            }
            Log("剪贴板就绪");

            await Task.Delay(50);

            var innerTask = await Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                if (targetHwnd != GetForegroundWindow())
                {
                    Log("⚠️ 窗口失焦，抢回焦点...");
                    RobustActivateWindow(targetHwnd);
                    await Task.Delay(50);
                }

                int clickX, clickY;
                if (_screenshotHelper.GetInputBoxClickCoordinates(targetHwnd, isWework, out clickX, out clickY))
                {
                    Log($"点击坐标: ({clickX}, {clickY})");
                    SetCursorPos(clickX, clickY);
                    mouse_event(MOUSEEVENTF_LEFTDOWN, clickX, clickY, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTUP, clickX, clickY, 0, 0);
                    await Task.Delay(50);
                }
                else
                {
                    Log("⚠️ 坐标计算失败");
                }

                bool result = false;
                if (AutoSendCheckBox.IsChecked == true)
                {
                    Log("执行自动发送...");
                    result = await PasteAndVerifySendAsync(filePath, true);
                    if (result)
                    {
                        _currentItemPasted = true;
                        _lastPastedStoreName = storeName;
                        StatusTextBlock.Text = $"✅ [自动] 已发送文件: {storeName}";
                        Log("✅ 发送成功");
                    }
                    else Log("❌ 发送失败");
                }
                else
                {
                    Log("执行手动粘贴...");
                    SimulatePaste();
                    Log("Ctrl+V 已发送");
                    StatusTextBlock.Text = $"📋 [自动] 已粘贴文件: {storeName}";
                    _currentItemPasted = true;
                    _lastPastedStoreName = storeName;
                    result = true;
                }
                return result;
            });

            return await innerTask;
        }


        // 兼容旧代码调用 (文件)
        private void PasteExcelFile(string storeName)
        {
            bool isWework = "企业微信".Equals(_currentSelectedNode?.Source);
            _ = PasteExcelFileAsync(storeName, isWework);
        }



        // MainWindow.xaml.cs

        private async Task<bool> PasteFullStoreInfoAsync(string storeName, bool isWework)
        {
            Action<string> Log = (msg) => System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [粘贴文本] {msg}");
            Log($"开始处理商家: {storeName}, isWework: {isWework}");

            if (string.IsNullOrEmpty(storeName)) return false;

            // 1. 窗口检查
            IntPtr targetHwnd = GetForegroundWindow();
            Log($"当前前台窗口句柄: {targetHwnd}");

            if (targetHwnd == IntPtr.Zero)
            {
                Log("❌ 无法获取前台窗口，取消粘贴");
                return false;
            }

            // 2. 数据准备
            List<string> trackingNumbers;
            lock (_dataLock)
            {
                if (!_storeData.TryGetValue(storeName, out trackingNumbers))
                {
                    Log("❌ 未找到商家数据");
                    return false;
                }
                trackingNumbers = trackingNumbers.ToList();
            }

            var sb = new StringBuilder();
            sb.AppendLine(storeName);
            foreach (var num in trackingNumbers) sb.AppendLine(num);
            sb.AppendLine(FIXED_MESSAGE);
            string fullText = sb.ToString();

            // 3. 剪贴板
            Log("正在设置剪贴板...");
            if (!await SetClipboardWithRetryAsync(fullText))
            {
                Log("❌ 剪贴板设置失败");
                return false;
            }
            Log("剪贴板设置成功");

            await Task.Delay(50);

            // 4. UI线程操作
            var innerTask = await Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                // 焦点复查
                if (targetHwnd != GetForegroundWindow())
                {
                    Log($"⚠️ 窗口失焦 (当前: {GetForegroundWindow()})，尝试抢回...");
                    RobustActivateWindow(targetHwnd);
                    await Task.Delay(50);
                }

                // 5. 坐标计算与点击
                int clickX = 0, clickY = 0;
                bool gotCoords = _screenshotHelper.GetInputBoxClickCoordinates(targetHwnd, isWework, out clickX, out clickY);
                Log($"获取点击坐标: {gotCoords}, X={clickX}, Y={clickY}");

                if (gotCoords)
                {
                    Log($"执行鼠标点击: ({clickX}, {clickY})");
                    SetCursorPos(clickX, clickY);
                    mouse_event(MOUSEEVENTF_LEFTDOWN, clickX, clickY, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTUP, clickX, clickY, 0, 0);
                    await Task.Delay(50);
                }
                else
                {
                    Log("⚠️ 无法计算坐标 (可能窗口最小化或句柄无效)");
                }

                // 6. 粘贴与发送
                bool result = false;
                if (AutoSendCheckBox.IsChecked == true)
                {
                    Log("模式: 自动发送");
                    // 注意：PasteAndVerifySendAsync 内部日志未在此处展示，需确保该函数也正常
                    result = await PasteAndVerifySendAsync(fullText, false);
                    if (result)
                    {
                        _currentItemPasted = true;
                        _lastPastedStoreName = storeName;
                        StatusTextBlock.Text = $"✅ [自动] 已发送: {storeName}";
                        Log("✅ 发送流程完成 (PasteAndVerifySendAsync 返回 true)");
                    }
                    else
                    {
                        Log("❌ 发送流程失败 (PasteAndVerifySendAsync 返回 false)");
                    }
                }
                else
                {
                    Log("模式: 手动发送 (仅粘贴)");
                    SimulatePaste();
                    Log("已模拟 Ctrl+V");

                    StatusTextBlock.Text = $"📋 [自动] 已粘贴: {storeName} (等待发送)";
                    _currentItemPasted = true;
                    _lastPastedStoreName = storeName;
                    result = true;
                }

                return result;
            });

            return await innerTask;
        }






        // 兼容旧代码调用 (文本)
        private void PasteFullStoreInfo(string storeName)
        {
            // 尝试推断，如果无法推断默认 false 或根据实际情况
            bool isWework = "企业微信".Equals(_currentSelectedNode?.Source);
            _ = PasteFullStoreInfoAsync(storeName, isWework);
        }







        private void SimulatePaste()
        {
            try
            {
                _inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"模拟粘贴失败: {ex.Message}";
            }
        }

        private void SimulateAltS()
        {
            try
            {
                _inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_S);
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"模拟发送失败: {ex.Message}";
            }
        }

        #endregion

        #region Excel 文件加载与处理

        private void LoadExcelButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "选择一个Excel文件"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                LoadExcelButton.IsEnabled = false;
                StatusTextBlock.Text = "正在读取和处理文件，请稍候...";

                Task.Run(() => LoadAndProcessExcel(openFileDialog.FileName))
                    .ContinueWith(t =>
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            LoadExcelButton.IsEnabled = true;
                            if (t.IsFaulted)
                            {
                                StatusTextBlock.Text = $"处理失败: {t.Exception?.InnerException?.Message ?? "未知错误"}";
                            }
                        });
                    });
            }
        }

        private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!Directory.Exists(_exportDirectory))
                {
                    Directory.CreateDirectory(_exportDirectory);
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = _exportDirectory,
                    UseShellExecute = true,
                    Verb = "open"
                });

                StatusTextBlock.Text = $"已打开导出目录: {_exportDirectory}";
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"无法打开目录: {ex.Message}";
            }
        }



        private void LoadAndProcessExcel(string filePath)
        {
            // 1. 清除内存中的旧数据
            lock (_dataLock)
            {
                _storeData.Clear();
                _exportedFilePaths.Clear();
            }

            // ✅ [新增] 清空 ExportedFiles 目录，防止旧文件残留
            try
            {
                if (Directory.Exists(_exportDirectory))
                {
                    string[] files = Directory.GetFiles(_exportDirectory);
                    foreach (string file in files)
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch (Exception delEx)
                        {
                            // 如果某个文件被占用无法删除，跳过并记录调试信息
                            System.Diagnostics.Debug.WriteLine($"[警告] 无法删除旧文件 '{file}': {delEx.Message}");
                        }
                    }
                }
                else
                {
                    // 如果目录不存在，顺便创建它
                    Directory.CreateDirectory(_exportDirectory);
                }
            }
            catch (Exception ex)
            {
                // 不让清理失败影响主流程，只做记录
                System.Diagnostics.Debug.WriteLine($"[警告] 清理导出目录失败: {ex.Message}");
            }

            // 2. 开始读取新文件
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null || worksheet.Dimension == null)
                    {
                        throw new InvalidOperationException("Excel文件或工作表为空");
                    }

                    // ✅ [保留修复] 本地辅助函数：防止长数字变成科学计数法
                    string GetSafeText(object cellValue)
                    {
                        if (cellValue == null) return string.Empty;

                        // 如果是数字类型，强制格式化为不带指数的字符串
                        if (cellValue is double || cellValue is decimal || cellValue is float || cellValue is long || cellValue is int)
                        {
                            return string.Format("{0:0.#############################}", cellValue);
                        }
                        return cellValue.ToString();
                    }

                    int rowCount = worksheet.Dimension.End.Row;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // 读取并清洗数据
                        string trackingNumber = GetSafeText(worksheet.Cells[row, 1].Value).Trim();
                        string storeName = GetSafeText(worksheet.Cells[row, 2].Value).Trim();

                        if (string.IsNullOrEmpty(trackingNumber) || string.IsNullOrEmpty(storeName)) continue;

                        lock (_dataLock)
                        {
                            if (!_storeData.ContainsKey(storeName))
                            {
                                _storeData[storeName] = new List<string>();
                            }
                            _storeData[storeName].Add(trackingNumber);
                        }
                    }
                }

                // 重置选中状态，避免指向不存在的旧索引
                _currentSelectedIndex = -1;
                _currentSelectedNode = null;

                // 处理并显示数据
                ProcessAndDisplayData();
            }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = $"文件读取错误: {ex.Message}";
                    StoreTreeView.ItemsSource = null;
                });
            }
        }




        /// <summary>
        /// ✅ [修改版] 处理显示数据，并初始化“失败归档区”
        /// </summary>
        private void ProcessAndDisplayData()
        {
            List<KeyValuePair<string, List<string>>> sortedStores;

            // ... (原有的排序逻辑保持不变) ...
            var infoMap = _businessInfoList.GroupBy(b => b.StoreName).ToDictionary(g => g.Key, g => g.FirstOrDefault());

            lock (_dataLock)
            {
                sortedStores = _storeData
                    .Select(kvp => new { Kvp = kvp, Info = infoMap.ContainsKey(kvp.Key) ? infoMap[kvp.Key] : null })
                    .OrderByDescending(x => x.Kvp.Value.Count > 100)
                    .ThenByDescending(x => !string.IsNullOrEmpty(x.Info?.GroupName))
                    .ThenBy(x =>
                    {
                        var src = x.Info?.Source;
                        if ("企业微信".Equals(src)) return 0;
                        if ("微信".Equals(src)) return 1;
                        return 2;
                    })
                    .ThenBy(x => x.Info?.GroupName)
                    .ThenByDescending(x => x.Kvp.Value.Count)
                    .Select(x => x.Kvp)
                    .ToList();
            }

            if (_currentFilter.Count > 0)
            {
                sortedStores = sortedStores.Where(kvp => _currentFilter.Any(filter => kvp.Key.Contains(filter, StringComparison.OrdinalIgnoreCase))).ToList();
            }

            // ✅ 改动：使用 ObservableCollection
            _treeViewCollection = new ObservableCollection<TreeViewNode>();

            try
            {
                Directory.CreateDirectory(_exportDirectory);
            }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"无法创建导出目录: {ex.Message}");
                return;
            }

            foreach (var kvp in sortedStores)
            {
                string storeName = kvp.Key;
                var trackingNumbers = kvp.Value;
                var parentNode = new TreeViewNode
                {
                    Header = $"{storeName} ({trackingNumbers.Count}条)",
                    StoreName = storeName
                };

                var busInfo = infoMap.ContainsKey(storeName) ? infoMap[storeName] : null;
                if (busInfo != null)
                {
                    parentNode.Source = busInfo.Source;
                    parentNode.GroupName = busInfo.GroupName;
                }

                if (trackingNumbers.Count > 100)
                {
                    parentNode.IsFileNode = true;
                    try
                    {
                        string filePath = CreateExcelFile(storeName, trackingNumbers, _exportDirectory);
                        lock (_dataLock) { _exportedFilePaths[storeName] = filePath; }
                        parentNode.Children.Add(new TreeViewNode { Text = "(单击复制名称，拖拽可导出文件)" });
                    }
                    catch (Exception ex)
                    {
                        parentNode.Children.Add(new TreeViewNode { Text = $"(文件创建失败: {ex.Message})" });
                    }
                }
                else
                {
                    parentNode.IsFileNode = false;
                    foreach (var number in trackingNumbers)
                    {
                        parentNode.Children.Add(new TreeViewNode { Text = number });
                    }
                }
                _treeViewCollection.Add(parentNode);
            }

            // ✅ 新增：在尾部添加“发送失败”分隔符节点
            _failureNode = new TreeViewNode
            {
                Header = "========== 🚫 发送失败/待重试 ==========",
                StoreName = "FAIL_SEPARATOR",
                Children = new ObservableCollection<TreeViewNode>() // 初始化子容器
            };
            _treeViewCollection.Add(_failureNode);

            Application.Current.Dispatcher.Invoke(() =>
            {
                // ✅ 绑定新的集合
                StoreTreeView.ItemsSource = _treeViewCollection;

                RebuildFlatNodeList();
                string filterInfo = _currentFilter.Count > 0 ? $"（已筛选 {_currentFilter.Count} 个关键词）" : "";
                StatusTextBlock.Text = $"处理完成，共显示 {sortedStores.Count} 个商家{filterInfo}";
            });
        }




        private string CreateExcelFile(string storeName, List<string> trackingNumbers, string outputDir)
        {
            string safeFileName = string.Join("_", storeName.Split(Path.GetInvalidFileNameChars()));
            string fileName = $"{safeFileName}未发货明细.xlsx";
            string filePath = Path.Combine(outputDir, fileName);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("未发货明细");
                worksheet.Cells[1, 1].Value = "运单号";
                worksheet.Cells[1, 2].Value = "店铺";
                using (var headerRange = worksheet.Cells[1, 1, 1, 2])
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }
                for (int i = 0; i < trackingNumbers.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = trackingNumbers[i];
                    worksheet.Cells[i + 2, 2].Value = storeName;
                }
                worksheet.Column(1).AutoFit(15);
                worksheet.Column(2).AutoFit(20);
                package.SaveAs(new FileInfo(filePath));
            }
            return filePath;
        }

        #endregion

        #region 筛选、删除、TreeView交互


        private void ApplyFilterButton_Click(object sender, RoutedEventArgs e)
        {
            string filterText = FilterTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(filterText))
            {
                StatusTextBlock.Text = "请输入筛选关键词";
                return;
            }
            _currentFilter = filterText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                       .Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
            if (_currentFilter.Count == 0)
            {
                StatusTextBlock.Text = "筛选条件为空";
                return;
            }

            // ✅ 筛选时重置选中状态
            _currentSelectedIndex = -1;
            _currentSelectedNode = null;

            ProcessAndDisplayData();
        }

        private void ClearFilterButton_Click(object sender, RoutedEventArgs e)
        {
            FilterTextBox.Clear();
            _currentFilter.Clear();

            // ✅ 清除筛选时重置选中状态
            _currentSelectedIndex = -1;
            _currentSelectedNode = null;

            ProcessAndDisplayData();
        }


    
        private void DeleteStoreButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string storeName)
            {
                lock (_dataLock)
                {
                    _storeData.Remove(storeName);
                    _exportedFilePaths.Remove(storeName);
                }
                ProcessAndDisplayData();
                StatusTextBlock.Text = $"已删除商家: '{storeName}'";
            }
        }

        private void TreeViewItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is TreeViewNode node && !string.IsNullOrEmpty(node.StoreName))
            {
                if (FindVisualParent<TreeViewItem>(element) is TreeViewItem treeViewItem)
                {
                    treeViewItem.IsSelected = true;
                }
                TriggerCopyOperation(node);
                e.Handled = true;
            }
        }

        private void TreeViewItem_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && !_isDragging)
            {
                Point position = e.GetPosition(null);
                if (Math.Abs(position.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(position.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    if (sender is FrameworkElement element && element.DataContext is TreeViewNode node && node.IsFileNode)
                    {
                        string filePath;
                        lock (_dataLock) { if (!_exportedFilePaths.TryGetValue(node.StoreName, out filePath)) return; }

                        if (File.Exists(filePath))
                        {
                            _isDragging = true;
                            DragDrop.DoDragDrop(element, new DataObject(DataFormats.FileDrop, new string[] { filePath }), DragDropEffects.Copy);
                            _isDragging = false;
                            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"文件 '{Path.GetFileName(filePath)}' 已拖拽导出");
                        }
                    }
                }
            }
        }

        private void TrackingNumber_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBlock textBlock && !string.IsNullOrEmpty(textBlock.Text))
            {
                string trackingNumber = textBlock.Text.Trim();
                if (trackingNumber.StartsWith("(") && trackingNumber.EndsWith(")")) return;

                Task.Run(async () =>
                {
                    if (await SetClipboardWithRetryAsync(trackingNumber))
                    {
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"已复制单号: {trackingNumber}");
                    }
                    else
                    {
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "复制单号失败");
                    }
                });
                e.Handled = true;
            }
        }

        /// <summary>
        /// ✅ 修复：选中项改变时重置粘贴状态
        /// </summary>
        private void StoreTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is TreeViewNode node && !string.IsNullOrEmpty(node.StoreName))
            {
                _currentSelectedNode = node;
                ResetSearchState();

                // ✅ 切换到新项时，检查是否需要重置粘贴状态
                if (_lastPastedStoreName != node.StoreName)
                {
                    _currentItemPasted = false;
                }

                if (_flatNodeList.Contains(node))
                {
                    _currentSelectedIndex = _flatNodeList.IndexOf(node);
                }
                if (Interlocked.CompareExchange(ref _copyingFlag, 1, 0) == 1) return;

                if (node.IsFileNode)
                {
                    CopyStoreNameOnly(node.StoreName);
                }
                else
                {
                    CopyFullStoreInfoToClipboard(node.StoreName);
                }
            }
        }

        #endregion


        /// <summary>
        /// ✅ 优化：仅更新单个节点，避免全量刷新
        /// </summary>

        /// <summary>
        /// ✅ 优化：仅更新单个节点，避免全量刷新
        /// </summary>
        private void UpdateNodeGroupInfo(string storeName, string groupName, string source)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // ❌ 原代码: if (StoreTreeView.ItemsSource is List<TreeViewNode> nodes)
                // ✅ 修复: 改为 IEnumerable<TreeViewNode>
                if (StoreTreeView.ItemsSource is IEnumerable<TreeViewNode> nodes)
                {
                    var targetNode = nodes.FirstOrDefault(n => n.StoreName == storeName);
                    if (targetNode != null)
                    {
                        targetNode.GroupName = groupName;
                        targetNode.Source = source;

                        var trackingCount = 0;
                        lock (_dataLock)
                        {
                            if (_storeData.ContainsKey(storeName))
                            {
                                trackingCount = _storeData[storeName].Count;
                            }
                        }
                        targetNode.Header = $"{storeName} ({trackingCount}条)";

                        StatusTextBlock.Text = $"[OCR] ✅ 已更新商家 '{storeName}' 的群名为: {groupName}";
                    }
                }
            });
        }






        #region 剪贴板操作

        private void CopyStoreNameOnly(string storeName)
        {
            Task.Run(async () =>
            {
                try
                {
                    if (!await SetClipboardWithRetryAsync(storeName)) throw new Exception("剪贴板被占用");
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"已复制商家名称: '{storeName}'");
                }
                catch (Exception ex)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"复制失败: {ex.Message}");
                }
                finally
                {
                    Interlocked.Exchange(ref _copyingFlag, 0);
                }
            });
        }

        private void CopyFullStoreInfoToClipboard(string storeName)
        {
            Task.Run(async () =>
            {
                try
                {
                    List<string> trackingNumbers;
                    lock (_dataLock)
                    {
                        if (!_storeData.TryGetValue(storeName, out trackingNumbers)) throw new Exception("未找到商家数据");
                        trackingNumbers = trackingNumbers.ToList();
                    }
                    var sb = new StringBuilder();
                    sb.AppendLine(storeName);
                    foreach (var num in trackingNumbers) sb.AppendLine(num);
                    sb.AppendLine(FIXED_MESSAGE);

                    if (!await SetClipboardWithRetryAsync(sb.ToString())) throw new Exception("剪贴板被占用");

                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"✅ 已复制 '{storeName}' 的完整信息 ({trackingNumbers.Count} 条单号)");
                }
                catch (Exception ex)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"❌ 复制失败: {ex.Message}");
                }
                finally
                {
                    Interlocked.Exchange(ref _copyingFlag, 0);
                }
            });
        }

        private async Task<bool> SetClipboardWithRetryAsync(string text)
        {
            for (int i = 0; i < 25; i++)
            {
                try
                {
                    bool success = await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        try
                        {
                            Clipboard.SetDataObject(text, true);
                            return true;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"剪贴板设置失败 (尝试 {i + 1}): {ex.Message}");
                            return false;
                        }
                    });
                    if (success) return true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"剪贴板操作异常: {ex.Message}");
                }
                await Task.Delay(20);
            }
            return false;
        }

        #endregion

        #region UI辅助

        private void FilterToggleButton_Checked(object sender, RoutedEventArgs e)
        {
            FilterPanel.Visibility = Visibility.Visible;
        }

        private void FilterToggleButton_Unchecked(object sender, RoutedEventArgs e)
        {
            FilterPanel.Visibility = Visibility.Collapsed;
        }

        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            return parentObject as T ?? FindVisualParent<T>(parentObject);
        }

        #endregion

        #region 业务逻辑辅助方法

        private bool IsTargetChatWindow(IntPtr hwnd, out string processName)
        {
            processName = null;
            if (hwnd == IntPtr.Zero) return false;

            GetWindowThreadProcessId(hwnd, out uint pid);
            if (pid == 0) return false;

            try
            {
                var process = System.Diagnostics.Process.GetProcessById((int)pid);
                if (process.ProcessName.Equals("WeChat", StringComparison.OrdinalIgnoreCase) ||
                    process.ProcessName.Equals("WXWork", StringComparison.OrdinalIgnoreCase))
                {
                    processName = process.ProcessName;
                    return true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"进程检查失败: {ex.Message}");
                return false;
            }

            return false;
        }

        #endregion

        #region 商家信息编辑

        private void EditBusInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentSelectedNode == null)
            {
                StatusTextBlock.Text = "⚠️ 请先在列表中选择一个商家。";
                return;
            }

            string storeName = _currentSelectedNode.StoreName;

            BusinessInfo infoToEdit = _businessInfoList.FirstOrDefault(b => b.StoreName == storeName);
            if (infoToEdit == null)
            {
                infoToEdit = new BusinessInfo { StoreName = storeName };
            }

            var editWindow = new EditBusInfoWindow(infoToEdit);
            bool? result = editWindow.ShowDialog();

            if (result == true)
            {
                BusinessInfo updatedInfo = editWindow.Info;

                // ✅ 更新业务信息列表
                _businessInfoList.RemoveAll(b => b.StoreName == updatedInfo.StoreName);

                if (!string.IsNullOrEmpty(updatedInfo.GroupName))
                {
                    _businessInfoList.Add(updatedInfo);
                }

                SaveBusinessInfo();

                // ✅ 使用局部更新，不重新生成整个 TreeView
                UpdateNodeGroupInfo(storeName, updatedInfo.GroupName, updatedInfo.Source);

                // ✅ 确保当前节点仍然选中
                if (_currentSelectedNode != null && _currentSelectedNode.StoreName == storeName)
                {
                    Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        EnsureNodeSelected(_currentSelectedNode);
                    }, System.Windows.Threading.DispatcherPriority.Loaded);
                }

                StatusTextBlock.Text = $"✅ 已更新商家 '{storeName}' 的信息。";
            }
        }







        private void SaveBusinessInfo()
        {
            string busInfoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "businfo.json");
            try
            {
                var options = new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                    WriteIndented = true
                };

                string json = JsonSerializer.Serialize(_businessInfoList, options);
                File.WriteAllText(busInfoPath, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"❌ 保存 businfo.json 失败: {ex.Message}";
            }
        }

        private void LoadBusinessInfo()
        {
            string busInfoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "businfo.json");
            _businessInfoList.Clear();

            if (File.Exists(busInfoPath))
            {
                try
                {
                    string json = File.ReadAllText(busInfoPath, Encoding.UTF8);
                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        _businessInfoList = JsonSerializer.Deserialize<List<BusinessInfo>>(json) ?? new List<BusinessInfo>();
                    }
                    StatusTextBlock.Text = $"✅ 成功加载 {_businessInfoList.Count} 条商家群聊信息。";
                }
                catch (Exception ex)
                {
                    StatusTextBlock.Text = $"❌ 加载 businfo.json 失败: {ex.Message}";
                    _businessInfoList = new List<BusinessInfo>();
                }
            }
            else
            {
                StatusTextBlock.Text = "ℹ️ 未找到 businfo.json，将仅使用商家名。";
            }
        }

        #endregion

        #region OCR结果处理

        /// <summary>
        /// ✅ OCR识别完成后的回调处理方法
        /// </summary>
        /// <summary>
        /// ✅ 修复：OCR识别完成后保持选中状态并继续流程
        /// </summary>
        /// <summary>
        /// ✅ 修复：OCR识别完成后使用局部更新
        /// </summary>








        /// <summary>
        /// ✅ 修复：OCR识别完成后保持选中状态
        /// </summary>
        private void HandleOcrResult(BusinessInfo ocrResult)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                try
                {
                    if (ocrResult == null || string.IsNullOrWhiteSpace(ocrResult.GroupName) ||
                        ocrResult.GroupName.Contains("失败") || ocrResult.GroupName.Contains("超时") ||
                        ocrResult.GroupName.Contains("未识别"))
                    {
                        StatusTextBlock.Text = $"[OCR] 商家 '{ocrResult?.StoreName}' 的群名识别失败或为空。";
                        return;
                    }

                    var storeName = ocrResult.StoreName;
                    var existingInfo = _businessInfoList.FirstOrDefault(b => b.StoreName == storeName);

                    if (existingInfo != null && !string.IsNullOrWhiteSpace(existingInfo.GroupName))
                    {
                        //StatusTextBlock.Text = $"[OCR] 商家 '{storeName}' 已有群名，本次识别结果被忽略。";
                        return;
                    }

                    if (existingInfo != null)
                    {
                        existingInfo.GroupName = ocrResult.GroupName;
                        existingInfo.Source = ocrResult.Source;
                    }
                    else
                    {
                        _businessInfoList.Add(ocrResult);
                    }

                    SaveBusinessInfo();

                    // ✅ 关键优化：使用局部更新替代全量刷新
                    UpdateNodeGroupInfo(storeName, ocrResult.GroupName, ocrResult.Source);

                    // ✅ 修复：保持当前选中状态
                    if (_currentSelectedNode != null && _currentSelectedNode.StoreName == storeName)
                    {
                        // 延迟一点确保UI更新完成
                        Application.Current.Dispatcher.InvokeAsync(() =>
                        {
                            EnsureNodeSelected(_currentSelectedNode);
                        }, System.Windows.Threading.DispatcherPriority.Loaded);
                    }

                    StatusTextBlock.Text = $"[OCR] ✅ 已保存商家 '{storeName}' 的群名";
                }
                catch (Exception ex)
                {
                    StatusTextBlock.Text = $"💥 处理OCR结果时出错: {ex.Message}";
                }
            });
        }

        /// <summary>
        /// ✅ 新增：确保节点保持选中状态
        /// </summary>
        private void EnsureNodeSelected(TreeViewNode node)
        {
            // ❌ 原代码: if (node == null || StoreTreeView.ItemsSource is not List<TreeViewNode> itemsSource)
            // ✅ 修复: 改为 IList<TreeViewNode>
            if (node == null || StoreTreeView.ItemsSource is not IList<TreeViewNode> itemsSource)
                return;

            int index = itemsSource.IndexOf(node);
            if (index < 0) return;

            if (StoreTreeView.ItemContainerGenerator.ContainerFromIndex(index) is TreeViewItem container)
            {
                if (!container.IsSelected)
                {
                    container.IsSelected = true;
                    container.Focus();
                }
            }
        }





        #endregion
        /// <summary>
        /// ✅ 新增：恢复TreeView的选中状态
        /// </summary>
        /// <param name="storeName">要恢复选中的商家名</param>
        /// <param name="fallbackIndex">如果找不到商家名，使用的备用索引</param>
        private void RestoreSelection(string storeName, int fallbackIndex)
        {
            if (string.IsNullOrEmpty(storeName))
            {
                // 如果没有商家名，尝试使用索引恢复
                if (fallbackIndex >= 0 && fallbackIndex < _flatNodeList.Count)
                {
                    _currentSelectedIndex = fallbackIndex;
                    var node = _flatNodeList[fallbackIndex];
                    _currentSelectedNode = node;
                    FocusAndSelectItem(node);
                }
                return;
            }

            // 重建扁平列表
            RebuildFlatNodeList();

            // 在新列表中查找同名商家
            var targetNode = _flatNodeList.FirstOrDefault(n => n.StoreName == storeName);

            if (targetNode != null)
            {
                _currentSelectedIndex = _flatNodeList.IndexOf(targetNode);
                _currentSelectedNode = targetNode;

                // 使用延迟确保UI已完全更新
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    FocusAndSelectItem(targetNode);
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
            else if (fallbackIndex >= 0 && fallbackIndex < _flatNodeList.Count)
            {
                // 如果找不到原商家，使用备用索引
                _currentSelectedIndex = fallbackIndex;
                var node = _flatNodeList[fallbackIndex];
                _currentSelectedNode = node;
                FocusAndSelectItem(node);
            }
        }



        #region TreeView选中优化

        /// <summary>
        /// ✅ 修复：改进的TreeView选中方法
        /// </summary>

        private void FocusAndSelectItem(TreeViewNode node)
        {
            if (node == null) return;

            // 1. 更新当前选中项记录
            _currentSelectedNode = node;

            // 2. 延迟执行，确保数据源更新后再操作 UI
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    // 尝试获取对应的 UI 容器 (TreeViewItem)
                    var container = StoreTreeView.ItemContainerGenerator.ContainerFromItem(node) as TreeViewItem;

                    // 🛑 如果容器为空（说明项在屏幕外，被虚拟化了，或者还没渲染）
                    if (container == null)
                    {
                        // 强制更新一次布局，让生成器尝试创建容器
                        StoreTreeView.UpdateLayout();
                        container = StoreTreeView.ItemContainerGenerator.ContainerFromItem(node) as TreeViewItem;
                    }

                    // ✅ 如果找到了容器
                    if (container != null)
                    {
                        // 核心：让容器自己把自己“搬”到视野内
                        container.BringIntoView();

                        // 设置选中和焦点
                        container.IsSelected = true;
                        container.Focus();
                    }
                    else
                    {
                        // ⚠️ 兜底：如果实在找不到（极少数情况），尝试手动触发 TreeView 刷新
                        StoreTreeView.Items.Refresh();
                        StatusTextBlock.Text = $"⚠️ 正在定位商家 '{node.StoreName}'..."; // 提示用户
                    }

                    // 触发后续的业务逻辑（如复制）
                    TriggerCopyOperation(node);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"选中项定位失败: {ex.Message}");
                }

            }, System.Windows.Threading.DispatcherPriority.Loaded);
        }

        #endregion
    }

    #region 数据模型

    public class BusinessInfo
    {
        [JsonPropertyName("商家名")]
        public string StoreName { get; set; }

        [JsonPropertyName("来源")]
        public string Source { get; set; }

        [JsonPropertyName("群名")]
        public string GroupName { get; set; }
    }



    public class TreeViewNode : INotifyPropertyChanged
    {
        private string _header;
        private string _groupName;
        private string _source;

        public string Header
        {
            get => _header;
            set
            {
                if (_header != value)
                {
                    _header = value;
                    OnPropertyChanged(nameof(Header));
                }
            }
        }

        public string Text { get; set; }
        public string StoreName { get; set; }
        public bool IsFileNode { get; set; }
        public ObservableCollection<TreeViewNode> Children { get; set; } = new ObservableCollection<TreeViewNode>();

        public string Source
        {
            get => _source;
            set
            {
                if (_source != value)
                {
                    _source = value;
                    OnPropertyChanged(nameof(Source));
                    OnPropertyChanged(nameof(GroupInfo));
                    OnPropertyChanged(nameof(HasGroupInfoVisibility));
                }
            }
        }

        public string GroupName
        {
            get => _groupName;
            set
            {
                if (_groupName != value)
                {
                    _groupName = value;
                    OnPropertyChanged(nameof(GroupName));
                    OnPropertyChanged(nameof(GroupInfo));
                    OnPropertyChanged(nameof(HasGroupInfoVisibility));
                }
            }
        }

        public string GroupInfo => string.IsNullOrEmpty(GroupName) ? "" : $"[{Source}] {GroupName}";
        public Visibility HasGroupInfoVisibility => string.IsNullOrEmpty(GroupName) ? Visibility.Collapsed : Visibility.Visible;

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }







    #endregion
}
