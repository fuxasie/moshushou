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
using moshushou.Input;

namespace moshushou
{
    public enum StorePayloadMode
    {
        Normal,
        Issue,
        CustomMessage
    }

    /// <summary>
    /// 商家节点的发送策略
    /// </summary>
    public enum SendStrategy
    {
        /// <summary>文本一次性发送（数据量较少时）</summary>
        TextDirect,
        /// <summary>文件发送（2列模式 >100 条）</summary>
        FileExcel,
        /// <summary>分段文本发送（4列/5列数据量较多时）</summary>
        TextSegmented
    }

    public partial class MainWindow : Window, IDisposable
    {
        private Dictionary<string, List<string>> _storeData = new Dictionary<string, List<string>>();
        private Dictionary<string, string> _exportedFilePaths = new Dictionary<string, string>();
        private Point _startPoint;
        private bool _isDragging = false;
        private int _copyingFlag = 0;
        private int _childNodeCopyInProgress = 0;
        private int _selectionCopyGuard = 0;
        private int _ctrlSpaceHotkeyInProgress = 0;
        private int _suppressSelectionOsdCount = 0;
        private bool _ctrlSpaceFallbackActive = false;

        // ✅ 低级键盘钩子：追踪物理 Ctrl 键状态（排除 SendInput 和本程序 Virtual HID 事件）
        private volatile bool _physicalCtrlPressed = false;
        private IntPtr _keyboardHookHandle = IntPtr.Zero;
        private LowLevelKeyboardProc _keyboardHookDelegate;
        private readonly object _dataLock = new object();
        private readonly Dictionary<string, int> _ctrlSpaceSegmentCursor = new Dictionary<string, int>(StringComparer.Ordinal);
        private string _exportDirectory;
        private int _currentSelectedIndex = -1;
        private List<TreeViewNode> _flatNodeList = new List<TreeViewNode>();
        // 子节点 -> 父节点的快速查找字典，在 RebuildFlatNodeList 时同步构建，O(1) 取父节点
        private Dictionary<TreeViewNode, TreeViewNode> _childParentMap = new Dictionary<TreeViewNode, TreeViewNode>();

        private TreeViewNode _currentSelectedNode = null;
        private List<string> _currentFilter = new List<string>();
        // ✅ 新增：保存进入筛选前的选中项（用于清空筛选时恢复）
        private string _preFilterSelectedStoreName = null;
        private int _preFilterSelectedIndex = -1;
        // ✅ 新增：搜素模式状态 (false=群名, true=商家)
        private bool _isStoreMode = true;
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



        // 【新增】用于记录已经进入过重试流程的商家，防止无限循环
        private HashSet<string> _failedStores = new HashSet<string>();
        // ✅ [新增] 记录被标记为“需人工”的商家，防止列表刷新后状态丢失
        private HashSet<string> _manualReviewStores = new HashSet<string>();
        // ✅ 记录已发送成功的商家，用于在主列表显示打钩状态
        private readonly object _sentStoreLock = new object();
        private HashSet<string> _sentStores = new HashSet<string>(StringComparer.Ordinal);
        // 发送历史记录去重：避免短时间内同一店铺的重复“选中”记录刷屏
        private string _lastSelectionHistoryStoreName = string.Empty;
        private DateTime _lastSelectionHistoryTime = DateTime.MinValue;

        private sealed class SegmentFailureInfo
        {
            public int FailedSegment { get; init; }
            public int TotalSegments { get; init; }
            public int SentSegments { get; init; }
            public int SentItems { get; init; }
            public int TotalItems { get; init; }
            public string Reason { get; init; } = "发送失败";
        }

        private readonly object _segmentFailureLock = new object();
        private readonly Dictionary<string, SegmentFailureInfo> _segmentFailureInfos = new Dictionary<string, SegmentFailureInfo>();

        // ✅ 新增：定义取消令牌源
        private CancellationTokenSource _searchCts;
        private CancellationTokenSource? _autoRunCts;

        // 微信/企业微信切换状态（false=先微信，true=先企业微信）
        private bool _isWeworkTurn = false;



        // ✅ [新增] 记录上一次成功进入的群名 (用于极速模式)
        private string _lastEnteredGroupName = null;


        // 搜索配置
        private SearchConfig _searchConfig;
        private SearchHelper _searchHelper;

        private int _searchInProgress = 0; // ✅ 改用 int 配合 Interlocked
        private readonly IInputBackend _inputBackend;

        private bool _currentItemPasted = false;
        private string _lastPastedStoreName = null;

        // ✅ [新增] 记录上一次成功的聊天窗口句柄 (用于极速模式抢焦点)
        private IntPtr _lastChatWindowHandle = IntPtr.Zero; // <--- 新增行
        // ✅ [新增] 记录本轮搜索阶段拿到的目标窗口句柄（用于失败后布局验证）
        private IntPtr _lastSearchWindowHandle = IntPtr.Zero;
        private bool _lastSearchWindowIsWework = false;
        // 最近一次布局验证通过的句柄（用于失败后的安全检测兜底）
        private IntPtr _lastLayoutVerifiedHwnd = IntPtr.Zero;
        private bool? _lastLayoutVerifiedIsWework = null;

        // ✅ 防抖写入和防抖复制：防止快速切换商家列表时造成的卡顿
        private System.Windows.Threading.DispatcherTimer _selectionSaveDebounceTimer;
        private string _pendingSaveStoreName;
        private System.Windows.Threading.DispatcherTimer _selectionCopyDebounceTimer;
        private TreeViewNode _pendingCopyNode;

        // 防止自动化搜索阶段被“选中即复制”逻辑污染剪贴板
        private int _clipboardSearchGuard = 0;
        // 标记最近一次自动流程失败是否发生在“已进入群聊并执行发送”阶段
        private int _lastFailureReachedPasteStage = 0;

        // 固定话术 (已迁移至 SearchConfig.FixedMessage 可配置)

        // 商家信息
        private List<BusinessInfo> _businessInfoList = new List<BusinessInfo>();


        
        // ✅ 新增：支持动态更新的数据集合
        private ObservableCollection<TreeViewNode> _treeViewCollection;
        // ✅ 新增：失败归档节点
        private TreeViewNode _failureNode;
        private DebugLogWindow? _debugLogWindow;
        private BusInfoManagerWindow? _busInfoManagerWindow;
        // 杂项设置窗口（非模态单例，保证 OsdWindow 可同时交互）
        private SettingsWindow? _settingsWindow;
        // ✅ 新增：自动化运行标志
        private bool _isAutoRunning = false;

        // ✅ 新增：问题件模式标志 (True=发送完整表格行, False=发送运单号+固定话术)
        private bool _isIssueMode = false;
        // ✅ [新增] 自定义话术模式 (4列: 运单号|话术|店铺|网点)
        private bool _isCustomMessageMode = false;
        // 当前已加载文件快照（用于调试窗口内重新解析）
        private string _lastLoadedFilePath = string.Empty;
        private int _lastLoadedColumnCount = 0;
        private string _activeTailMessage = string.Empty;
        private int _activeIssueSegmentStartCount = 30;



        // ✅ 新增：保存状态防抖计时器
        private System.Windows.Threading.DispatcherTimer _saveDebounceTimer;



        // ✅ 新增：F1/F2 热键 ID
        private const int HOTKEY_F1 = 9007;
        private const int HOTKEY_F2 = 9008;
        private const int HOTKEY_CTRL_SPACE = 9009;
        private const int HOTKEY_CTRL_SHIFT_SPACE = 9010;
        private const int HOTKEY_W = 9011;
        private const int HOTKEY_S = 9012;
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

        // Windows API 声明
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        // 激活窗口
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
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
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        // ✅ 低级键盘钩子 P/Invoke
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;
        private const uint LLKHF_INJECTED = 0x00000010;

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        // 虚拟键码
        private const uint VK_UP = 0x26;
        private const uint VK_DOWN = 0x28;
        private const uint VK_LEFT = 0x25;
        private const uint VK_RIGHT = 0x27;
        private const uint VK_RETURN = 0x0D;
        private const uint VK_OEM_7 = 0xDE;
        private const uint VK_SPACE = 0x20;
        private const uint VK_W = 0x57;
        private const uint VK_S = 0x53;
        private const int VK_CONTROL_KEY = 0x11;
        private const int VK_LCONTROL_KEY = 0xA2;
        private const int VK_RCONTROL_KEY = 0xA3;

        // 修饰键
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;



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
            DebugLogManager.Initialize();

            LoadBusinessInfo();

            // 初始化选中防抖：控制文件状态写入频率
            _selectionSaveDebounceTimer = new System.Windows.Threading.DispatcherTimer();
            _selectionSaveDebounceTimer.Interval = TimeSpan.FromMilliseconds(500);
            _selectionSaveDebounceTimer.Tick += SelectionSaveDebounceTimer_Tick;

            // 初始化选中复制防抖：控制连续按下上下键时抢剪贴板
            _selectionCopyDebounceTimer = new System.Windows.Threading.DispatcherTimer();
            _selectionCopyDebounceTimer.Interval = TimeSpan.FromMilliseconds(200);
            _selectionCopyDebounceTimer.Tick += SelectionCopyDebounceTimer_Tick;


            ExcelPackage.License.SetNonCommercialPersonal("fff");

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            if (string.IsNullOrEmpty(baseDir))
            {
                baseDir = Environment.CurrentDirectory;
            }
            _exportDirectory = Path.Combine(baseDir, "ExportedFiles");

            _searchConfig = SearchConfig.Load();
            DriverInstallationManager.EnsureVirtualHidInstalled(_searchConfig);
            _inputBackend = InputBackendFactory.Create(_searchConfig, message =>
            {
                DebugLogManager.Log("Input", message);
                Debug.WriteLine($"[Input] {message}");
            });
            MouseHelper.Configure(_inputBackend);
            _activeTailMessage = _searchConfig?.FixedMessage ?? string.Empty;
            EmitRecentStoreHistoryToDebugLog();

            string screenshotBaseDir = Path.Combine(baseDir, "Screenshots");
            _screenshotHelper = new ScreenshotHelper(screenshotBaseDir, _searchConfig, (msg) => {
                DebugLogManager.Log("Screenshot", msg);
                Application.Current.Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = msg;
                });
            });

            _searchHelper = new SearchHelper(_searchConfig, _inputBackend, (msg) =>
            {
                DebugLogManager.Log("Search", msg);
                Application.Current.Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = msg;
                });
            });

            // ✅ 初始化防抖计时器 (延迟 1秒 保存)
            _saveDebounceTimer = new System.Windows.Threading.DispatcherTimer();
            _saveDebounceTimer.Interval = TimeSpan.FromSeconds(1);
            _saveDebounceTimer.Tick += (s, e) =>
            {
                _saveDebounceTimer.Stop();
                // 在后台线程执行保存操作，避免阻塞 UI
                Task.Run(() => 
                { 
                    try { _searchConfig?.Save(); } catch { }
                });
            };

            StoreTreeView.SelectedItemChanged += StoreTreeView_SelectedItemChanged;
            this.Loaded += MainWindow_Loaded;
            this.Closing += MainWindow_Closing;
            UpdateListProgressStatus();
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _windowHandle = new WindowInteropHelper(this).Handle;
            _source = HwndSource.FromHwnd(_windowHandle);
            _source.AddHook(HwndHook);
            InstallKeyboardHook();
            UpdatePollingModeButtonState();
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Dispose();
            Application.Current.Shutdown();
        }

        public void Dispose()
        {
            UnregisterGlobalHotkeys();
            UninstallKeyboardHook();
            try
            {
                _inputBackend.ReleaseAll();
                _inputBackend.Dispose();
            }
            catch
            {
            }
            if (_source != null)
            {
                _source.RemoveHook(HwndHook);
                _source.Dispose();
                _source = null;
            }
        }

        internal void PrepareForVirtualHidDriverUninstall()
        {
            if (_inputBackend is VirtualHidBackend virtualHidBackend)
            {
                virtualHidBackend.DisconnectForDriverMaintenance();
                DebugLogManager.Log("Input", "已释放 Virtual HID 控制设备句柄，准备卸载驱动。");
            }
        }

        #region 低级键盘钩子 - 追踪物理 Ctrl 状态

        /// <summary>
        /// 安装低级键盘钩子，用于追踪物理 Ctrl 键的按下/释放状态。
        /// 通过 LLKHF_INJECTED 排除 SendInput，并通过 SyntheticInputTracker
        /// 排除本程序的 Virtual HID 键盘事件。
        /// </summary>
        private void InstallKeyboardHook()
        {
            if (_keyboardHookHandle != IntPtr.Zero) return;

            _keyboardHookDelegate = KeyboardHookCallback;
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                _keyboardHookHandle = SetWindowsHookEx(
                    WH_KEYBOARD_LL,
                    _keyboardHookDelegate,
                    GetModuleHandle(curModule.ModuleName),
                    0);
            }

            if (_keyboardHookHandle == IntPtr.Zero)
            {
                Debug.WriteLine("[KeyboardHook] 安装失败");
            }
            else
            {
                Debug.WriteLine("[KeyboardHook] 安装成功");
            }
        }

        private void UninstallKeyboardHook()
        {
            if (_keyboardHookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_keyboardHookHandle);
                _keyboardHookHandle = IntPtr.Zero;
                Debug.WriteLine("[KeyboardHook] 已卸载");
            }
        }

        /// <summary>
        /// 低级键盘钩子回调：仅追踪物理 Ctrl 键状态，忽略 SendInput 注入的事件。
        /// </summary>
        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var kbs = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                bool isInjected = (kbs.flags & LLKHF_INJECTED) != 0;

                int msg = wParam.ToInt32();
                bool isDown = msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN;
                bool isUp = msg == WM_KEYUP || msg == WM_SYSKEYUP;
                bool isOwnVirtualHidEvent = !isInjected &&
                                            (isDown || isUp) &&
                                            SyntheticInputTracker.TryConsume(kbs.vkCode, isDown);

                // 仅关注物理按键事件（排除 SendInput 和本程序的 Virtual HID）
                if (!isInjected && !isOwnVirtualHidEvent &&
                    (kbs.vkCode == (uint)VK_CONTROL_KEY ||
                     kbs.vkCode == (uint)VK_LCONTROL_KEY ||
                     kbs.vkCode == (uint)VK_RCONTROL_KEY))
                {
                    if (isDown)
                    {
                        _physicalCtrlPressed = true;
                    }
                    else if (isUp)
                    {
                        _physicalCtrlPressed = false;
                    }
                }
            }

            return CallNextHookEx(_keyboardHookHandle, nCode, wParam, lParam);
        }

        #endregion
 
        private void SelectionCopyDebounceTimer_Tick(object sender, EventArgs e)
        {
            _selectionCopyDebounceTimer.Stop();
            var nodeToCopy = _pendingCopyNode;
            if (nodeToCopy == null) return;

            if (_isAutoRunning ||
                Volatile.Read(ref _clipboardSearchGuard) > 0 ||
                Volatile.Read(ref _selectionCopyGuard) > 0)
            {
                return;
            }

            if (Interlocked.CompareExchange(ref _copyingFlag, 1, 0) == 1) return;

            // ✅ [修复] 子节点优先复制 RawData (分段或单行)
            if (!string.IsNullOrEmpty(nodeToCopy.RawData) && nodeToCopy.Strategy != SendStrategy.FileExcel)
            {
                string textToCopy = nodeToCopy.RawData;
                Task.Run(async () =>
                {
                    try
                    {
                        if (!await SetClipboardWithRetryAsync(textToCopy)) 
                            throw new Exception("剪贴板被占用");
                        
                        Application.Current.Dispatcher.Invoke(() => 
                            StatusTextBlock.Text = "✅ 已复制选中的内容");
                    }
                    catch (Exception ex)
                    {
                        Application.Current.Dispatcher.Invoke(() => 
                            StatusTextBlock.Text = $"❌ 复制失败: {ex.Message}");
                    }
                    finally
                    {
                        Interlocked.Exchange(ref _copyingFlag, 0);
                    }
                });
                return;
            }

            // 主列表复制内容跟随搜索模式
            CopyPreferredSearchText(nodeToCopy);
            Interlocked.Exchange(ref _copyingFlag, 0);
        }

        private void SelectionSaveDebounceTimer_Tick(object sender, EventArgs e)
        {
            _selectionSaveDebounceTimer.Stop();
            if (!string.IsNullOrEmpty(_pendingSaveStoreName) && _pendingSaveStoreName != "FAIL_SEPARATOR")
            {
                SaveFileState(_pendingSaveStoreName);
                
                // 也执行一次针对这节点的选择历史记录，需要在数据中再匹配一下 Node
                var node = _flatNodeList?.FirstOrDefault(n => n.StoreName == _pendingSaveStoreName);
                if (node != null) RecordStoreSelectionHistory(node);
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

        private async void TogglePollingModeButton_Click(object sender, RoutedEventArgs e)
        {
            _isStoreMode = !_isStoreMode;
            UpdatePollingModeButtonState();

            if (_isAutoRunning || Volatile.Read(ref _clipboardSearchGuard) > 0)
            {
                return;
            }

            TreeViewNode nodeToCopy = _currentSelectedNode;
            if ((nodeToCopy == null || string.IsNullOrEmpty(nodeToCopy.StoreName) || nodeToCopy.StoreName == "FAIL_SEPARATOR")
                && _currentSelectedIndex >= 0
                && _currentSelectedIndex < _flatNodeList.Count)
            {
                nodeToCopy = _flatNodeList[_currentSelectedIndex];
            }

            if (nodeToCopy == null || string.IsNullOrEmpty(nodeToCopy.StoreName) || nodeToCopy.StoreName == "FAIL_SEPARATOR")
            {
                return;
            }

            if (!TryGetPreferredSearchCopyText(nodeToCopy, out string copyText, out string copyType))
            {
                return;
            }

            bool fallbackToStoreName = !_isStoreMode &&
                                       string.Equals(copyType, "商家名", StringComparison.Ordinal);
            if (!await SetClipboardWithRetryAsync(copyText))
            {
                StatusTextBlock.Text = "模式已切换，但复制失败：剪贴板被占用";
                return;
            }

            string modeLabel = _isStoreMode ? "商家搜索" : "群名搜索";
            string fallbackTip = fallbackToStoreName ? "（无群名已回退商家名）" : string.Empty;
            StatusTextBlock.Text = $"已切换为 [{modeLabel}] 模式，并复制{copyType}: '{copyText}'{fallbackTip}";
        }

        private void UpdatePollingModeButtonState()
        {
            if (TogglePollingModeButton == null) return;

            if (_isStoreMode)
            {
                TogglePollingModeButton.Content = "🏪";
                TogglePollingModeButton.ToolTip = "当前模式：优先搜索商家名\n点击切换为群名模式";
                // 橙色表示商家模式
                TogglePollingModeButton.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F59E0B")); 
                StatusTextBlock.Text = "已切换为 [商家搜索] 模式";
            }
            else
            {
                TogglePollingModeButton.Content = "💬";
                TogglePollingModeButton.ToolTip = "当前模式：优先搜索群聊名\n点击切换为商家模式";
                // 蓝色表示群名模式
                TogglePollingModeButton.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#3B82F6")); 
                StatusTextBlock.Text = "已切换为 [群名搜索] 模式";
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            // 单例保护：若窗口已打开则激活，防止重复创建
            if (_settingsWindow != null && _settingsWindow.IsLoaded)
            {
                if (_settingsWindow.WindowState == WindowState.Minimized)
                    _settingsWindow.WindowState = WindowState.Normal;
                _settingsWindow.Activate();
                return;
            }

            // 使用非模态 Show()，允许用户在设置窗口打开时继续操作 OsdWindow
            _settingsWindow = new SettingsWindow(_searchConfig, this)
            {
                Owner = this
            };
            _settingsWindow.Closed += (_, __) => _settingsWindow = null;
            _settingsWindow.Show();
        }

        private void OpenDebugLogWindowButton_Click(object sender, RoutedEventArgs e)
        {
            if (_debugLogWindow == null || !_debugLogWindow.IsLoaded)
            {
                _debugLogWindow = new DebugLogWindow
                {
                    Owner = this
                };
                _debugLogWindow.Closed += (_, __) => _debugLogWindow = null;
                _debugLogWindow.Show();
                return;
            }

            if (_debugLogWindow.WindowState == WindowState.Minimized)
            {
                _debugLogWindow.WindowState = WindowState.Normal;
            }

            _debugLogWindow.Activate();
        }

        private void OpenBusInfoManagerButton_Click(object sender, RoutedEventArgs e)
        {
            if (_busInfoManagerWindow != null && _busInfoManagerWindow.IsLoaded)
            {
                if (_busInfoManagerWindow.WindowState == WindowState.Minimized)
                {
                    _busInfoManagerWindow.WindowState = WindowState.Normal;
                }

                _busInfoManagerWindow.Activate();
                SyncBusInfoManagerWithCurrentSelection();
                return;
            }

            var sourceItems = _businessInfoList
                .Select(item => new BusinessInfo
                {
                    StoreName = item.StoreName,
                    GroupName = item.GroupName,
                    Source = item.Source
                })
                .ToList();

            _busInfoManagerWindow = new BusInfoManagerWindow(sourceItems)
            {
                Owner = this
            };

            _busInfoManagerWindow.Saved += OnBusInfoManagerSaved;
            _busInfoManagerWindow.Closed += OnBusInfoManagerClosed;
            _busInfoManagerWindow.Show();

            SyncBusInfoManagerWithCurrentSelection();
            StatusTextBlock.Text = "已打开 businfo 映射窗口，可在主列表点选并联动填写。";
        }

        private void OnBusInfoManagerSaved(List<BusinessInfo> updatedBusinessInfos)
        {
            if (updatedBusinessInfos == null)
            {
                return;
            }

            int selectedIndexSnapshot = 0;
            string? selectedStoreSnapshot = GetCurrentSelectedStoreName(out selectedIndexSnapshot);

            _businessInfoList = updatedBusinessInfos
                .Select(item => new BusinessInfo
                {
                    StoreName = item.StoreName,
                    GroupName = item.GroupName,
                    Source = item.Source
                })
                .ToList();

            SaveBusinessInfo();
            ProcessAndDisplayData();

            if (!string.IsNullOrWhiteSpace(selectedStoreSnapshot))
            {
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    RestoreSelection(selectedStoreSnapshot, selectedIndexSnapshot);
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }

            StatusTextBlock.Text = $"✅ 已更新 businfo 映射，共 {_businessInfoList.Count} 条。";
        }

        private void OnBusInfoManagerClosed(object? sender, EventArgs e)
        {
            if (_busInfoManagerWindow != null)
            {
                _busInfoManagerWindow.Saved -= OnBusInfoManagerSaved;
                _busInfoManagerWindow.Closed -= OnBusInfoManagerClosed;
            }

            _busInfoManagerWindow = null;
        }

        private void SyncBusInfoManagerWithCurrentSelection()
        {
            if (_busInfoManagerWindow == null || !_busInfoManagerWindow.IsLoaded)
            {
                return;
            }

            TreeViewNode? selectedNode = StoreTreeView.SelectedItem as TreeViewNode;
            if (selectedNode == null || string.IsNullOrWhiteSpace(selectedNode.StoreName))
            {
                selectedNode = _currentSelectedNode;
            }

            if (selectedNode == null || string.IsNullOrWhiteSpace(selectedNode.StoreName) || selectedNode.StoreName == "FAIL_SEPARATOR")
            {
                return;
            }

            TreeViewNode rootNode = selectedNode;
            if (TryResolveRootNode(selectedNode, out TreeViewNode resolvedRoot) && resolvedRoot != null)
            {
                rootNode = resolvedRoot;
            }

            string storeName = NormalizeStoreNameForBusinessInfo(rootNode.StoreName);
            if (string.IsNullOrWhiteSpace(storeName))
            {
                storeName = rootNode.StoreName?.Trim() ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(storeName))
            {
                return;
            }

            BusinessInfo? mappedInfo = _businessInfoList.FirstOrDefault(item =>
                string.Equals(
                    NormalizeStoreNameForBusinessInfo(item.StoreName),
                    storeName,
                    StringComparison.OrdinalIgnoreCase));

            string groupName = mappedInfo?.GroupName?.Trim() ?? rootNode.GroupName?.Trim() ?? string.Empty;
            string source = mappedInfo?.Source?.Trim() ?? rootNode.Source?.Trim() ?? string.Empty;
            _busInfoManagerWindow.SyncFromMainSelection(storeName, groupName, source);
        }

        private void ScrollToSelectedButton_Click(object sender, RoutedEventArgs e)
        {
            TreeViewNode? targetNode = StoreTreeView.SelectedItem as TreeViewNode;
            if (targetNode == null || string.IsNullOrWhiteSpace(targetNode.StoreName))
            {
                targetNode = _currentSelectedNode;
            }

            if ((targetNode == null || string.IsNullOrWhiteSpace(targetNode.StoreName)) &&
                _currentSelectedIndex >= 0 &&
                _currentSelectedIndex < _flatNodeList.Count)
            {
                targetNode = _flatNodeList[_currentSelectedIndex];
            }

            if (targetNode == null || string.IsNullOrWhiteSpace(targetNode.StoreName))
            {
                StatusTextBlock.Text = "当前没有可定位的选中项";
                return;
            }

            _currentSelectedNode = targetNode;
            if (_flatNodeList.Count == 0)
            {
                RebuildFlatNodeList();
            }
            if (_flatNodeList.Contains(targetNode))
            {
                _currentSelectedIndex = _flatNodeList.IndexOf(targetNode);
            }

            SyncTreeViewSelection(targetNode);
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                if (StoreTreeView.ItemContainerGenerator.ContainerFromItem(targetNode) is TreeViewItem container)
                {
                    container.Focus();
                }
            }, System.Windows.Threading.DispatcherPriority.Background);

            StatusTextBlock.Text = $"已定位到：{targetNode.StoreName}";
            UpdateListProgressStatus();
        }

        public sealed class ParseOverrideDebugContext
        {
            public string FilePath { get; init; } = string.Empty;
            public int DetectedColumnCount { get; init; }
            public string ParseMode { get; init; } = FileParseModes.Auto;
            public int TrackingColumn { get; init; } = 1;
            public int StoreColumn { get; init; } = 2;
            public int IssueSegmentStartCount { get; init; } = 30;
            public string TailMessage { get; init; } = string.Empty;
        }

        public ParseOverrideDebugContext? GetCurrentFileParseContext()
        {
            string filePath = _lastLoadedFilePath;
            if (string.IsNullOrWhiteSpace(filePath))
            {
                filePath = _searchConfig?.LastOpenedFilePath?.Trim() ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                return null;
            }

            FileParseOverride? fileOverride = FindFileParseOverride(filePath);
            int defaultStoreColumn = _isIssueMode ? 4 : (_isCustomMessageMode ? 3 : 2);
            string parseMode = fileOverride != null
                ? FileParseModes.Normalize(fileOverride.ParseMode)
                : (_isIssueMode || _isCustomMessageMode ? FileParseModes.Issue : FileParseModes.Magician);

            return new ParseOverrideDebugContext
            {
                FilePath = filePath,
                DetectedColumnCount = _lastLoadedColumnCount,
                ParseMode = parseMode,
                TrackingColumn = Math.Max(1, fileOverride?.TrackingColumn ?? 1),
                StoreColumn = Math.Max(1, fileOverride?.StoreColumn ?? defaultStoreColumn),
                IssueSegmentStartCount = ResolveIssueSegmentStartCount(fileOverride),
                TailMessage = fileOverride?.TailMessage?.Trim() ?? GetCurrentTailMessage()
            };
        }

        public async Task<bool> ApplyCurrentFileParseOverrideAsync(FileParseOverride overrideRule)
        {
            if (_searchConfig == null)
            {
                StatusTextBlock.Text = "配置未初始化，无法应用解析规则";
                return false;
            }

            string filePath = _lastLoadedFilePath;
            if (string.IsNullOrWhiteSpace(filePath))
            {
                filePath = _searchConfig.LastOpenedFilePath?.Trim() ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                StatusTextBlock.Text = "请先加载一个文件，再修改解析方式";
                return false;
            }

            if (!File.Exists(filePath))
            {
                StatusTextBlock.Text = $"当前文件不存在：{filePath}";
                return false;
            }

            FileParseOverride? previousOverride = FindFileParseOverride(filePath);
            int previousIssueSegmentStartCount = ResolveIssueSegmentStartCount(previousOverride);
            string parseMode = FileParseModes.Normalize(overrideRule?.ParseMode);
            int trackingColumn = Math.Max(1, overrideRule?.TrackingColumn ?? 1);
            int storeColumn = Math.Max(1, overrideRule?.StoreColumn ?? 2);
            int issueSegmentStartCount = Math.Max(2, overrideRule?.IssueSegmentStartCount ?? GetDefaultSegmentSize());
            string tailMessage = overrideRule?.TailMessage?.Trim() ?? string.Empty;
            bool issueSegmentStartChanged =
                string.Equals(parseMode, FileParseModes.Issue, StringComparison.Ordinal) &&
                issueSegmentStartCount != previousIssueSegmentStartCount;

            if (string.Equals(parseMode, FileParseModes.Magician, StringComparison.Ordinal) &&
                string.IsNullOrWhiteSpace(tailMessage))
            {
                StatusTextBlock.Text = "魔术师格式需要填写尾部自定义话术";
                return false;
            }

            _searchConfig.FileParseOverrides ??= new List<FileParseOverride>();
            string normalizedCurrentFilePath = NormalizeFilePathKey(filePath);
            _searchConfig.FileParseOverrides.RemoveAll(item =>
                item != null &&
                !string.IsNullOrWhiteSpace(item.FilePath) &&
                string.Equals(
                    NormalizeFilePathKey(item.FilePath),
                    normalizedCurrentFilePath,
                    StringComparison.OrdinalIgnoreCase));

            if (!string.Equals(parseMode, FileParseModes.Auto, StringComparison.Ordinal))
            {
                _searchConfig.FileParseOverrides.Add(new FileParseOverride
                {
                    FilePath = normalizedCurrentFilePath,
                    ParseMode = parseMode,
                    TrackingColumn = trackingColumn,
                    StoreColumn = storeColumn,
                    IssueSegmentStartCount = issueSegmentStartCount,
                    TailMessage = tailMessage
                });
            }

            bool segmentProgressReset = false;
            if (issueSegmentStartChanged)
            {
                segmentProgressReset = ResetIssueSegmentStateForCurrentFile(filePath);
                DebugLogManager.Log(
                    "ParseMode",
                    $"问题件分段起始条数变更：{previousIssueSegmentStartCount} -> {issueSegmentStartCount}，已清空历史分段勾选/进度");
            }

            _searchConfig.Save();

            string modeTip = parseMode switch
            {
                FileParseModes.Magician => "魔术师格式（两列表格+尾部话术）",
                FileParseModes.Issue => $"问题件格式（运单列={trackingColumn}，店铺列={storeColumn}，分段起始条数={issueSegmentStartCount}）",
                _ => "自动识别"
            };
            string resetTip = segmentProgressReset ? "，已清空历史分段勾选/进度" : string.Empty;

            StatusTextBlock.Text = $"已应用解析规则：{modeTip}{resetTip}，正在重载文件...";
            DebugLogManager.Log("ParseMode", $"应用解析规则：{modeTip} | 文件={Path.GetFileName(filePath)}");

            LoadExcelButton.IsEnabled = false;
            try
            {
                await Task.Run(() => LoadAndProcessExcel(filePath));
                return true;
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"重载失败: {ex.Message}";
                return false;
            }
            finally
            {
                LoadExcelButton.IsEnabled = true;
            }
        }

        private bool ResetIssueSegmentStateForCurrentFile(string filePath)
        {
            if (_searchConfig == null || string.IsNullOrWhiteSpace(filePath))
            {
                return false;
            }

            bool changed = false;
            string normalizedFilePath = NormalizeFilePathKey(filePath);

            _searchConfig.LastIssueFileState ??= new FileState();
            var issueState = _searchConfig.LastIssueFileState;
            if (!string.IsNullOrWhiteSpace(issueState.FilePath) &&
                string.Equals(
                    NormalizeFilePathKey(issueState.FilePath),
                    normalizedFilePath,
                    StringComparison.OrdinalIgnoreCase))
            {
                if (issueState.SegmentFailures != null && issueState.SegmentFailures.Count > 0)
                {
                    issueState.SegmentFailures.Clear();
                    changed = true;
                }
                else if (issueState.SegmentFailures == null)
                {
                    issueState.SegmentFailures = new List<SegmentFailureState>();
                }
            }

            lock (_segmentFailureLock)
            {
                if (_segmentFailureInfos.Count > 0)
                {
                    _segmentFailureInfos.Clear();
                    changed = true;
                }
            }

            lock (_sentStoreLock)
            {
                if (_sentStores.Count > 0)
                {
                    _sentStores.Clear();
                    changed = true;
                }
            }

            if (_ctrlSpaceSegmentCursor.Count > 0)
            {
                _ctrlSpaceSegmentCursor.Clear();
                changed = true;
            }

            return changed;
        }

        private FileParseOverride? FindFileParseOverride(string filePath)
        {
            if (_searchConfig?.FileParseOverrides == null || string.IsNullOrWhiteSpace(filePath))
            {
                return null;
            }

            string normalizedFilePath = NormalizeFilePathKey(filePath);
            return _searchConfig.FileParseOverrides
                .Where(item =>
                    item != null &&
                    !string.IsNullOrWhiteSpace(item.FilePath) &&
                    string.Equals(
                        NormalizeFilePathKey(item.FilePath),
                        normalizedFilePath,
                        StringComparison.OrdinalIgnoreCase))
                .LastOrDefault();
        }

        private string GetCurrentTailMessage()
        {
            string fallback = _searchConfig?.FixedMessage ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(_activeTailMessage))
            {
                return _activeTailMessage;
            }

            _activeTailMessage = fallback;
            return fallback;
        }

        private int GetDefaultSegmentSize()
        {
            int size = _searchConfig?.SegmentSize ?? 30;
            if (size <= 0)
            {
                size = 30;
            }

            return size;
        }

        private static string NormalizeFilePathKey(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFullPath(path.Trim());
            }
            catch
            {
                return path.Trim();
            }
        }

        private int ResolveIssueSegmentStartCount(FileParseOverride? fileOverride)
        {
            int fallback = Math.Max(2, GetDefaultSegmentSize());
            if (fileOverride == null)
            {
                return fallback;
            }

            return fileOverride.IssueSegmentStartCount >= 2
                ? fileOverride.IssueSegmentStartCount
                : fallback;
        }

        private int GetSegmentSizeForPayloadMode(StorePayloadMode mode)
        {
            if (mode == StorePayloadMode.Issue)
            {
                return Math.Max(2, _activeIssueSegmentStartCount);
            }

            return Math.Max(1, GetDefaultSegmentSize());
        }

        private int GetSegmentSizeForStore(string storeName)
        {
            if (string.IsNullOrWhiteSpace(storeName))
            {
                return Math.Max(1, GetDefaultSegmentSize());
            }

            List<string>? rows = null;
            lock (_dataLock)
            {
                if (_storeData.TryGetValue(storeName, out var foundRows) && foundRows != null)
                {
                    rows = foundRows.ToList();
                }
            }

            if (rows == null || rows.Count == 0)
            {
                return Math.Max(1, GetDefaultSegmentSize());
            }

            StorePayloadMode mode = ResolveStorePayloadMode(storeName, rows);
            return GetSegmentSizeForPayloadMode(mode);
        }






        /// <summary>
        /// ✅ [优化版] 窗口激活：移除置顶和线程挂接，降低风控触发风险
        /// </summary>
        private bool RobustActivateWindow(IntPtr targetHwnd)
        {
            if (targetHwnd == IntPtr.Zero) return false;

            // ✅ [优化] 不再置顶，避免触发风控
            // EnsureWindowTopMost(targetHwnd);

            // 【状态层】检查是否最小化，如果是则还原
            if (IsIconic(targetHwnd))
            {
                ShowWindow(targetHwnd, SW_RESTORE);
                System.Threading.Thread.Sleep(200); // 还原动画需要时间
            }
            else
            {
                ShowWindow(targetHwnd, SW_SHOW);
            }

            // ✅ [优化] 简化激活逻辑，不使用 AttachThreadInput
            // 只调用一次 SetForegroundWindow，不循环重试
            SetForegroundWindow(targetHwnd);
            System.Threading.Thread.Sleep(100); // 给系统一点反应时间

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
                bool spaceRegistered = RegisterHotKey(_windowHandle, HOTKEY_CTRL_SPACE, MOD_CONTROL, VK_SPACE);
                bool spaceFallbackRegistered = false;
                _ctrlSpaceFallbackActive = false;

                // Ctrl+Space 常被输入法占用，失败时自动降级到 Ctrl+Shift+Space
                if (!spaceRegistered)
                {
                    spaceFallbackRegistered = RegisterHotKey(_windowHandle, HOTKEY_CTRL_SHIFT_SPACE, MOD_CONTROL | MOD_SHIFT, VK_SPACE);
                    _ctrlSpaceFallbackActive = spaceFallbackRegistered;
                }

                bool wRegistered = RegisterHotKey(_windowHandle, HOTKEY_W, MOD_CONTROL, VK_W);
                bool sRegistered = RegisterHotKey(_windowHandle, HOTKEY_S, MOD_CONTROL, VK_S);

                // 2. 注册新增的 F1 / F2 (无修饰键)
                bool f1Registered = RegisterHotKey(_windowHandle, HOTKEY_F1, 0, VK_F1);
                bool f2Registered = RegisterHotKey(_windowHandle, HOTKEY_F2, 0, VK_F2);

                if (upRegistered &&
                    downRegistered &&
                    leftRegistered &&
                    rightRegistered &&
                    enterRegistered &&
                    quoteRegistered &&
                    (spaceRegistered || spaceFallbackRegistered) &&
                    wRegistered &&
                    sRegistered &&
                    f1Registered &&
                    f2Registered)
                {
                    _globalHotkeysRegistered = true;
                    string ctrlSpaceTip = _ctrlSpaceFallbackActive ? "Ctrl+Shift+Space快捷粘贴(兼容)" : "Ctrl+Space快捷粘贴";
                    StatusTextBlock.Text = $"快捷键：F1开始自动/F2停止，Ctrl+↑↓/WS切换，Ctrl+←复制店铺，Ctrl+→粘贴发送，{ctrlSpaceTip}，Ctrl+Enter手动搜索，Ctrl+'识别群名";
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
                UnregisterHotKey(_windowHandle, HOTKEY_CTRL_SPACE);
                UnregisterHotKey(_windowHandle, HOTKEY_CTRL_SHIFT_SPACE);
                UnregisterHotKey(_windowHandle, HOTKEY_W);
                UnregisterHotKey(_windowHandle, HOTKEY_S);

                // 注销 F1 / F2
                UnregisterHotKey(_windowHandle, HOTKEY_F1);
                UnregisterHotKey(_windowHandle, HOTKEY_F2);

                _globalHotkeysRegistered = false;
                _ctrlSpaceFallbackActive = false;
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

                if (id == HOTKEY_UP || id == HOTKEY_W)
                {
                    // 向上导航 
                    Application.Current.Dispatcher.Invoke(() => NavigateTreeView(-1));
                    shouldHandle = true;
                }
                else if (id == HOTKEY_DOWN || id == HOTKEY_S)
                {
                    // 向下导航 
                    Application.Current.Dispatcher.Invoke(() => NavigateTreeView(1));
                    shouldHandle = true;
                }
                else if (id == HOTKEY_LEFT)
                {
                    // Ctrl+Left: 仅粘贴名称 (保持原样)
                    Application.Current.Dispatcher.Invoke(() => PasteCurrentStoreName());
                    shouldHandle = true;
                }
                else if (id == HOTKEY_RIGHT)
                {
                    // ✅ Ctrl+Right: 调用粘贴流程
                    // (写入剪贴板 -> 盲粘贴 -> 自动发送 -> 后置补全群名)
                    Application.Current.Dispatcher.InvokeAsync(async () =>
                    {
                        await ManualPasteProcessAsync();
                    });
                    shouldHandle = true;
                }
                else if (id == HOTKEY_CTRL_SPACE || id == HOTKEY_CTRL_SHIFT_SPACE)
                {
                    // Ctrl+Space: 按当前选中项执行快捷粘贴/发送（含分段推进）
                    Application.Current.Dispatcher.InvokeAsync(async () =>
                    {
                        if (Interlocked.CompareExchange(ref _ctrlSpaceHotkeyInProgress, 1, 0) == 1)
                        {
                            StatusTextBlock.Text = "⌛ Ctrl+Space 正在执行，请稍后再按";
                            return;
                        }

                        try
                        {
                            await HandleCtrlSpaceHotkeyAsync();
                        }
                        finally
                        {
                            Interlocked.Exchange(ref _ctrlSpaceHotkeyInProgress, 0);
                        }
                    });
                    shouldHandle = true;
                }
                else if (id == HOTKEY_QUOTE)
                {
                    // ✅ Ctrl+': 仅 OCR 识别当前窗口群名，不粘贴不发送
                    Application.Current.Dispatcher.InvokeAsync(async () =>
                    {
                        await OcrRecognizeGroupNameOnlyAsync();
                    });
                    shouldHandle = true;
                }
                else if (id == HOTKEY_ENTER)
                {
                    // ✅ Ctrl+Enter: 手动搜索/前进

                    // 关键：强制释放物理按住的 Ctrl 键，防止干扰后续的搜索指令
                    try
                    {
                        _inputBackend.KeyUp(InputKey.LeftControl);
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
                    // F2: 优先停止自动化；否则停止当前 Ctrl+Enter 手动流程
                    if (_isAutoRunning)
                    {
                        StopAutoSending();
                    }
                    else
                    {
                        StopManualSearchFlow();
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

            bool restartedSegmentedFromBeginning = false;
            bool startedFromSelectedSegment = false;
            string restartedStoreName = string.Empty;
            string selectedSegmentStartMessage = string.Empty;
            TreeViewNode? selectedNodeSnapshot = _currentSelectedNode;

            if (selectedNodeSnapshot != null &&
                TryResolveRootNode(selectedNodeSnapshot, out TreeViewNode selectedRootNode) &&
                selectedRootNode != null &&
                selectedRootNode.Strategy == SendStrategy.TextSegmented)
            {
                if (!ReferenceEquals(selectedNodeSnapshot, selectedRootNode) &&
                    TryApplyAutoSegmentStartFromSelectedChild(selectedRootNode, selectedNodeSnapshot, out string manualStartMessage))
                {
                    startedFromSelectedSegment = true;
                    selectedSegmentStartMessage = manualStartMessage;
                }
                else if (ResetCompletedSegmentedStoreProgress(selectedRootNode, refreshHeader: true))
                {
                    // 按用户意图：F1 选中分段店铺且已全部成功时，重置分段进度并从第1段重新发送
                    restartedSegmentedFromBeginning = true;
                    restartedStoreName = selectedRootNode.StoreName;
                }

                // 自动化搜索依赖主节点的群名信息；若当前选中子段，统一切回主节点
                SelectNodeWithoutCopy(selectedRootNode);
            }

            _autoRunCts?.Cancel();
            _autoRunCts?.Dispose();
            _autoRunCts = new CancellationTokenSource();
            var autoCts = _autoRunCts;

            _isAutoRunning = true;
            StatusTextBlock.Text = startedFromSelectedSegment
                ? selectedSegmentStartMessage
                : restartedSegmentedFromBeginning
                    ? $"🚀 [F1] 自动化发送模式已启动！已重置 '{restartedStoreName}' 的分段成功进度，将从第1段重发。(按 F2 停止)"
                    : "🚀 [F1] 自动化发送模式已启动！将从当前选中项继续。(按 F2 停止)";

            // 启动后台循环任务
            Task.Run(() => AutoProcessLoop(autoCts));
        }

        private bool TryApplyAutoSegmentStartFromSelectedChild(TreeViewNode rootNode, TreeViewNode selectedChildNode, out string statusMessage)
        {
            statusMessage = string.Empty;

            if (rootNode == null ||
                selectedChildNode == null ||
                ReferenceEquals(rootNode, selectedChildNode) ||
                rootNode.Strategy != SendStrategy.TextSegmented ||
                rootNode.Children == null ||
                rootNode.Children.Count <= 0 ||
                string.IsNullOrWhiteSpace(rootNode.StoreName))
            {
                return false;
            }

            int selectedIndex = rootNode.Children.IndexOf(selectedChildNode);
            if (selectedIndex < 0)
            {
                return false;
            }

            int totalSegments = rootNode.Children.Count;
            int sentSegments = Math.Max(0, Math.Min(selectedIndex, totalSegments));
            int totalItems = 0;
            lock (_dataLock)
            {
                if (_storeData.TryGetValue(rootNode.StoreName, out var rows))
                {
                    totalItems = rows.Count;
                }
            }

            if (totalItems <= 0)
            {
                totalItems = rootNode.Children.Sum(child => CountContentLines(child.RawData));
            }

            int sentItems = 0;
            for (int i = 0; i < selectedIndex && i < rootNode.Children.Count; i++)
            {
                sentItems += CountContentLines(rootNode.Children[i].RawData);
            }

            if (totalItems > 0)
            {
                sentItems = Math.Min(sentItems, totalItems);
            }

            ClearStoreSentMark(rootNode.StoreName, refreshHeader: false);
            _ctrlSpaceSegmentCursor[rootNode.StoreName] = selectedIndex;
            UpdateSegmentProgressVisual(
                rootNode.StoreName,
                sentSegments: sentSegments,
                totalSegments: totalSegments,
                sentItems: sentItems,
                totalItems: totalItems,
                reason: sentSegments > 0 ? "发送中(手动定位)" : "发送中");
            SaveFileState(rootNode.StoreName);

            statusMessage =
                $"🚀 [F1] 自动化发送模式已启动！将从 '{rootNode.StoreName}' 第 {selectedIndex + 1}/{totalSegments} 段开始，前 {sentSegments} 段已标记完成。(按 F2 停止)";
            return true;
        }

        private void StopAutoSending()
        {
            _isAutoRunning = false;
            _autoRunCts?.Cancel();
            StatusTextBlock.Text = "🛑 [F2] 自动化发送已停止。";
        }

        private void StopManualSearchFlow()
        {
            var cts = Interlocked.Exchange(ref _searchCts, null);
            if (cts == null)
            {
                StatusTextBlock.Text = "🛑 [F2] 当前无可停止的手动流程。";
                return;
            }

            try
            {
                cts.Cancel();
            }
            catch
            {
                // 忽略取消过程中的状态竞争异常
            }

            StatusTextBlock.Text = "🛑 [F2] 当前 Ctrl+Enter 流程已停止。";
        }





        private async Task AutoProcessLoop(CancellationTokenSource autoCts)
        {
            var token = autoCts.Token;
            int consecutiveFailures = 0;
            // 自动化失败后原地重试次数（每个商家、每个阶段独立计数）
            const int maxInlineRetriesPerStore = 1;
            var inlineRetryUsedMap = new Dictionary<string, int>(StringComparer.Ordinal);

            string BuildInlineRetryKey(string storeName, bool inRetryArea)
            {
                return $"{(inRetryArea ? "retry" : "main")}::{storeName}";
            }

            void ClearInlineRetryState(string? storeName)
            {
                if (string.IsNullOrWhiteSpace(storeName))
                {
                    return;
                }

                inlineRetryUsedMap.Remove(BuildInlineRetryKey(storeName, false));
                inlineRetryUsedMap.Remove(BuildInlineRetryKey(storeName, true));
            }

            while (_isAutoRunning && !token.IsCancellationRequested)
            {
                bool shouldStop = false;
                bool? attemptAppIsWework = null;
                bool shouldRunSecurityCheckAfterFailure = true;

                // 1. 状态检查与停止条件
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    // A. 焦点自救
                    if (_currentSelectedNode == null && _currentSelectedIndex >= 0 && _currentSelectedIndex < _flatNodeList.Count)
                    {
                        var rescueNode = _flatNodeList[_currentSelectedIndex];
                        FocusAndSelectItem(rescueNode);
                        _currentSelectedNode = rescueNode;
                    }

                    // B. 遇到分隔符自动跳过
                    if (_currentSelectedNode != null && _currentSelectedNode.StoreName == "FAIL_SEPARATOR")
                    {
                        StatusTextBlock.Text = "⬇️ 进入自动重试区...";
                        consecutiveFailures = 0;
                        if (!TryNavigateToNextNodeForAuto())
                        {
                            StatusTextBlock.Text = "🏁 自动重试区已处理完毕，自动化停止。";
                            shouldStop = true;
                        }
                        return;
                    }

                    // C. 检查是否到达列表末尾
                    if (_currentSelectedNode == null)
                    {
                        StatusTextBlock.Text = "🏁 列表已处理完毕，自动化停止。";
                        shouldStop = true;
                        return;
                    }

                    // D. F1 自动化不按“整店铺✅”跳过。
                    //    分段店铺是否跳过由分段断点状态决定（仅跳过已发送的段）。

                    // E. 检查是否有群名
                    if (string.IsNullOrWhiteSpace(_currentSelectedNode.GroupName))
                    {
                        StatusTextBlock.Text = $"🛑 商家 '{_currentSelectedNode.StoreName}' 无群名，自动停止。";
                        shouldStop = true;
                        FocusAndSelectItem(_currentSelectedNode);
                        return;
                    }

                    string src = _currentSelectedNode.Source?.Trim() ?? string.Empty;
                    if (string.Equals(src, "企业微信", StringComparison.OrdinalIgnoreCase))
                    {
                        attemptAppIsWework = true;
                    }
                    else if (string.Equals(src, "微信", StringComparison.OrdinalIgnoreCase))
                    {
                        attemptAppIsWework = false;
                    }
                });

                if (shouldStop)
                {
                    _isAutoRunning = false;
                    break;
                }

                if (_currentSelectedNode == null || _currentSelectedNode.StoreName == "FAIL_SEPARATOR")
                {
                    try { await Task.Delay(100, token); } catch (OperationCanceledException) { break; }
                    continue;
                }

                // 2. 核心处理
                // ✅ [修复] 调用带点击逻辑的版本（与 Ctrl+Enter 手动模式一致）
                bool success = await SearchCurrentItemAsync(true, token);

                if (!_isAutoRunning || token.IsCancellationRequested) break;

                var autoLoopUiTask = await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
                    this.Activate();
                    this.Focus();
                    SetForegroundWindow(_windowHandle);

                    if (success)
                    {
                        string? successStoreName = _currentSelectedNode?.StoreName;
                        ClearInlineRetryState(successStoreName);
                        consecutiveFailures = 0;
                        if (TryNavigateToNextNodeForAuto())
                        {
                            StatusTextBlock.Text += " [成功] 下一条...";
                        }
                        else
                        {
                            _isAutoRunning = false;
                            StatusTextBlock.Text = "🏁 列表已处理完毕，自动化停止。";
                        }
                    }
                    else
                    {
                        bool reachedPasteStage = Volatile.Read(ref _lastFailureReachedPasteStage) == 1;
                        if (_currentSelectedNode != null && !string.IsNullOrWhiteSpace(_currentSelectedNode.StoreName))
                        {
                            string failureStageLabel = reachedPasteStage ? "发送阶段失败" : "搜索/进群阶段失败";
                            RecordStoreSendHistory(_currentSelectedNode.StoreName, "自动化发送", false, failureStageLabel);
                        }
                        bool shouldVerifyLayout = true;
                        bool isLayoutValid = true;
                        bool? matchedAppIsWework = null;
                        IntPtr matchedHwnd = IntPtr.Zero;
                        IntPtr lastCheckedHwnd = IntPtr.Zero;

                        if (shouldVerifyLayout)
                        {
                            // 任何失败在进入重试处理前都必须做布局验证。
                            // 规则：布局正常 -> 允许移动到重试区；布局异常 -> 立即停止自动化。
                            TreeViewNode activeNode = _currentSelectedNode;
                            bool? preferredAppIsWework = null;
                            if (activeNode != null)
                            {
                                string src = activeNode.Source?.Trim() ?? string.Empty;
                                if (string.Equals(src, "企业微信", StringComparison.OrdinalIgnoreCase))
                                {
                                    preferredAppIsWework = true;
                                }
                                else if (string.Equals(src, "微信", StringComparison.OrdinalIgnoreCase))
                                {
                                    preferredAppIsWework = false;
                                }
                            }

                            var appCandidates = new List<bool>();
                            if (preferredAppIsWework.HasValue)
                            {
                                // Source 明确时，仅验证该应用；失败即停，不再跨应用兜底。
                                appCandidates.Add(preferredAppIsWework.Value);
                            }
                            else if (_lastSearchWindowHandle != IntPtr.Zero)
                            {
                                // Source 缺失时，优先使用本轮搜索窗口所属应用。
                                appCandidates.Add(_lastSearchWindowIsWework);
                            }
                            else
                            {
                                // 最后兜底：Source 缺失且无搜索句柄时，才两端都试。
                                appCandidates.Add(false);
                                appCandidates.Add(true);
                            }
                            string nodeSource = activeNode?.Source?.Trim() ?? "<null>";
                            string candidateLog = string.Join(" -> ", appCandidates.Select(x => x ? "企业微信" : "微信"));
                            string failureStage = reachedPasteStage ? "发送阶段失败" : "搜索/进群阶段失败";
                            System.Diagnostics.Debug.WriteLine($"[AutoLoop] 失败后布局验证候选: Stage={failureStage}, Source='{nodeSource}', Candidates={candidateLog}");
                            System.Diagnostics.Debug.WriteLine(
                                $"[AutoLoop] 布局验证策略: Source明确={(preferredAppIsWework.HasValue ? "是" : "否")}, " +
                                $"仅当前应用={(preferredAppIsWework.HasValue ? "是" : "否")}");

                            // VerifyChatWindowLayoutAsync 内部已包含 3 次重试，这里不再重复外层轮询。
                            const int layoutRoundsPerApp = 1;
                            isLayoutValid = false;

                            foreach (bool appIsWework in appCandidates.Distinct())
                            {
                                string targetClass = appIsWework ? _searchConfig.WeworkWindowClassName : _searchConfig.WechatWindowClassName;
                                string appName = appIsWework ? "企业微信" : "微信";
                                var handleCandidates = new List<IntPtr>();

                                // 1) 优先使用本轮搜索阶段记录的窗口句柄（最可信）
                                if (_lastSearchWindowHandle != IntPtr.Zero && _lastSearchWindowIsWework == appIsWework)
                                {
                                    if (IsTargetChatWindow(_lastSearchWindowHandle, out string cachedProcessName) &&
                                        IsProcessNameForApp(cachedProcessName, appIsWework))
                                    {
                                        handleCandidates.Add(_lastSearchWindowHandle);
                                        System.Diagnostics.Debug.WriteLine(
                                            $"[AutoLoop] 布局验证命中搜索阶段句柄: App={appName}, Hwnd={_lastSearchWindowHandle}, Proc={cachedProcessName}");
                                    }
                                    else
                                    {
                                        System.Diagnostics.Debug.WriteLine(
                                            $"[AutoLoop] 搜索阶段句柄不可用: App={appName}, Hwnd={_lastSearchWindowHandle}");
                                    }
                                }

                                // 2) 使用 SearchHelper 的稳健查找（按进程+枚举窗口）
                                IntPtr robustHwnd = _searchHelper?.TryGetAppWindowHandle(appIsWework) ?? IntPtr.Zero;
                                if (robustHwnd != IntPtr.Zero)
                                {
                                    handleCandidates.Add(robustHwnd);
                                    System.Diagnostics.Debug.WriteLine(
                                        $"[AutoLoop] 布局验证稳健查找命中: App={appName}, Hwnd={robustHwnd}");
                                }

                                // 3) 最后回退类名直查（兼容旧窗口类）
                                IntPtr classHwnd = FindWindow(targetClass, null);
                                if (classHwnd != IntPtr.Zero)
                                {
                                    handleCandidates.Add(classHwnd);
                                    System.Diagnostics.Debug.WriteLine(
                                        $"[AutoLoop] 布局验证类名查找命中: App={appName}, Class={targetClass}, Hwnd={classHwnd}");
                                }

                                var distinctHandles = handleCandidates
                                    .Where(h => h != IntPtr.Zero)
                                    .Distinct()
                                    .ToList();

                                if (distinctHandles.Count == 0)
                                {
                                    System.Diagnostics.Debug.WriteLine(
                                        $"[AutoLoop] 布局验证前未找到窗口: App={appName}, Class={targetClass} (无可用句柄)");
                                    continue;
                                }

                                foreach (IntPtr targetHwnd in distinctHandles)
                                {
                                    lastCheckedHwnd = targetHwnd;

                                    for (int round = 0; round < layoutRoundsPerApp; round++)
                                    {
                                        RobustActivateWindow(targetHwnd);
                                        await Task.Delay(300);

                                        bool oneRoundValid = await Task.Run(async () => await _screenshotHelper.VerifyChatWindowLayoutAsync(targetHwnd, appIsWework));
                                        if (oneRoundValid)
                                        {
                                            isLayoutValid = true;
                                            matchedAppIsWework = appIsWework;
                                            matchedHwnd = targetHwnd;
                                            break;
                                        }

                                        if (round < layoutRoundsPerApp - 1)
                                        {
                                            await Task.Delay(350);
                                        }
                                    }

                                    if (isLayoutValid) break;
                                }

                                if (isLayoutValid) break;
                            }
                        }

                        if (!isLayoutValid)
                        {
                            // 🛑 布局验证失败 -> 认为是严重异常 (窗口关闭/退出登录/被挡住)
                            // 此时直接停止自动化，不进入重试区
                            this.Activate();
                            _lastLayoutVerifiedHwnd = IntPtr.Zero;
                            _lastLayoutVerifiedIsWework = null;
                             
                            _isAutoRunning = false;
                            StatusTextBlock.Text = "🛑 [严重] 窗口布局异常(检测不到群名/消息/输入框)，自动化已停止！";
                            System.Diagnostics.Debug.WriteLine($"[AutoLoop] 布局验证失败 (LastHwnd={lastCheckedHwnd})，停止自动化。");
                        }
                        else
                        {
                            string matchedAppName = matchedAppIsWework == true ? "企业微信" : "微信";
                            System.Diagnostics.Debug.WriteLine($"[AutoLoop] 布局验证通过: App={matchedAppName}, Hwnd={matchedHwnd}");
                            _lastLayoutVerifiedHwnd = matchedHwnd;
                            _lastLayoutVerifiedIsWework = matchedAppIsWework;

                            var node = _currentSelectedNode;
                            if (node != null)
                            {
                                bool isRetryAreaStore = _failedStores.Contains(node.StoreName);
                                string retryKey = BuildInlineRetryKey(node.StoreName, isRetryAreaStore);
                                int usedRetries = inlineRetryUsedMap.TryGetValue(retryKey, out int tmpUsed) ? tmpUsed : 0;

                                if (usedRetries < maxInlineRetriesPerStore)
                                {
                                    inlineRetryUsedMap[retryKey] = usedRetries + 1;
                                    shouldRunSecurityCheckAfterFailure = false;
                                    string phaseTag = isRetryAreaStore ? "重试区" : "主列表";
                                    StatusTextBlock.Text += $" [失败] 原位重试({usedRetries + 1}/{maxInlineRetriesPerStore})...";
                                    DebugLogManager.Log("自动重试", $"商家={node.StoreName}, 阶段={phaseTag}, 第{usedRetries + 1}次原位重试");
                                    System.Diagnostics.Debug.WriteLine($"[AutoLoop] 失败后原位重试: Store={node.StoreName}, Phase={phaseTag}, Retry={usedRetries + 1}/{maxInlineRetriesPerStore}");

                                    if (node.Strategy == SendStrategy.TextSegmented)
                                    {
                                        ApplySegmentFailureProgressToNode(node);
                                    }
                                }
                                else
                                {
                                    inlineRetryUsedMap.Remove(retryKey);
                                    consecutiveFailures++;

                                    if (isRetryAreaStore)
                                    {
                                        // --- 重试区再次失败（原位重试已耗尽） ---
                                        StatusTextBlock.Text += $" [重试失败 {consecutiveFailures}] 标记需人工...";
                                        DebugLogManager.Log("自动重试", $"商家={node.StoreName}, 重试区失败已耗尽，标记需人工");
                                        RecordStoreSendHistory(node.StoreName, "发送", false, "重试区失败已耗尽，标记需人工");

                                        // ✅ 记录到人工名单
                                        _manualReviewStores.Add(node.StoreName);
                                        ClearStoreSentMark(node.StoreName);
                                        SaveFileState();

                                        RefreshStoreNodeHeader(node);

                                        ClearInlineRetryState(node.StoreName);
                                        if (!TryNavigateToNextNodeForAuto())
                                        {
                                            _isAutoRunning = false;
                                            StatusTextBlock.Text += " 🏁 自动重试区已处理完毕，自动化停止。";
                                        }
                                    }
                                    else
                                    {
                                        // --- 主列表失败（原位重试已耗尽） ---
                                        StatusTextBlock.Text += $" [初次失败 {consecutiveFailures}] 移入重试区...";
                                        DebugLogManager.Log("自动重试", $"商家={node.StoreName}, 主列表重试已耗尽，移入重试区");
                                        RecordStoreSendHistory(node.StoreName, "发送", false, "主列表重试已耗尽，移入重试区");

                                        _failedStores.Add(node.StoreName);
                                        ClearStoreSentMark(node.StoreName);
                                        ApplySegmentFailureProgressToNode(node);
                                        MoveCurrentToFailureNode();
                                        SaveFileState();

                                        // 进入重试区后是新阶段，清掉主列表阶段计数
                                        ClearInlineRetryState(node.StoreName);
                                    }
                                }
                            }

                            if (consecutiveFailures >= 3)
                            {
                                _isAutoRunning = false;
                                StatusTextBlock.Text += " 🛑 [熔断] 连续失败 3 次，已自动停止！";
                            }
                        }
                    }
                });
                await autoLoopUiTask;

                if (!_isAutoRunning) break;

                // ✅ 关键修复：成功/失败后都需要等待，确保微信有足够时间处理粘贴内容
                if (success)
                {
                    try
                    {
                        await Task.Delay(500, token); // 成功后等待 500ms，防止下一个商家的粘贴覆盖当前商家
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                }
                else
                {
                    if (shouldRunSecurityCheckAfterFailure)
                    {
                        // ✅ [新增] 失败时检测是否出现了“设备异常/需扫码”的安全验证
                        if (!attemptAppIsWework.HasValue && _lastSearchWindowHandle != IntPtr.Zero)
                        {
                            attemptAppIsWework = _lastSearchWindowIsWework;
                        }

                        bool isSecurityBlock = CheckForSecurityVerification(attemptAppIsWework);
                        if (isSecurityBlock)
                        {
                            _isAutoRunning = false;
                            await Application.Current.Dispatcher.InvokeAsync(() =>
                            {
                                StatusTextBlock.Text = "🛑[停止] 无法检测到窗口，自动化紧急停止！";
                            });
                            break;
                        }
                    }

                    try
                    {
                        await Task.Delay(500, token); // 失败后等待 500ms
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                }
            }

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                if (!_isAutoRunning && consecutiveFailures < 3)
                    StatusTextBlock.Text += " (已停止)";
            });

            autoCts.Dispose();
            if (ReferenceEquals(_autoRunCts, autoCts))
            {
                _autoRunCts = null;
            }
        }

        private void MoveCurrentToFailureNode()
        {
            // 此方法必须在 UI 线程调用
            var node = _currentSelectedNode;

            // 校验
            if (node == null || node == _failureNode || node.StoreName == "FAIL_SEPARATOR") return;

            // 🔒 锁定索引：因为我们即将把当前项移走，当前位置会被“下一项”填补
            // 所以我们不需要改变 _currentSelectedIndex，它自然就会指向“下一项”
            int targetSlotIndex = _currentSelectedIndex;

            // 从集合中移除并添加到末尾
            if (_treeViewCollection.Contains(node))
            {
                _treeViewCollection.Remove(node); // 移除当前
                _treeViewCollection.Add(node);    // 加到最后
            }

            // 重建扁平索引
            RebuildFlatNodeList();

            // 保持选中在“下一个合适项”，避免置空
            if (!SelectBestNode(targetSlotIndex))
            {
                _currentSelectedNode = null;
            }
        }
        #endregion



        private async Task SmartAdvanceOrSearchAsync()
        {
            // 1. ✅ 取消上一次正在进行的任务 (如果存在)
            if (_searchCts != null)
            {
                _searchCts.Cancel();
                _searchCts.Dispose();
                _searchCts = null;

                // 📝 记录中断日志（不再手动翻转轮询状态，由 PerformLightweightSearchAsync 统一管理）
                StatusTextBlock.Text = "⏭️ [中断] 正在重新启动...";
                System.Diagnostics.Debug.WriteLine($"[SmartAdvanceOrSearchAsync] 任务被打断，当前 _isWeworkTurn={_isWeworkTurn}");
            }

            // 2. 创建新的令牌
            _searchCts = new CancellationTokenSource();
            var token = _searchCts.Token;

            try
            {
                // 使用 Dispatcher 确保 UI 访问安全
                var smartFlowUiTask = await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    if (_currentSelectedNode == null || string.IsNullOrEmpty(_currentSelectedNode.StoreName))
                    {
                        StatusTextBlock.Text = "⚠️ 请先选择一个商家";
                        return;
                    }

                    // 智能前进判断逻辑 (保持原样)
                    string currentStoreName = _currentSelectedNode.StoreName;
                    if (_currentItemPasted && _lastPastedStoreName == currentStoreName)
                    {
                        StatusTextBlock.Text = "⏭️ [手动] 前进到下一项...";
                        // 注意：NavigateTreeView 是同步的，如果需要支持取消，这里其实很快，通常不用改
                        NavigateTreeView(1);

                        if (_currentSelectedNode == null || string.IsNullOrEmpty(_currentSelectedNode.StoreName))
                        {
                            StatusTextBlock.Text = "✅ 列表到底了！";
                            return;
                        }
                        _currentItemPasted = false;
                        _lastPastedStoreName = null;

                        // 给一点时间让UI刷新，支持取消
                        try { await Task.Delay(100, token); } catch (TaskCanceledException) { return; }
                    }
                    else
                    {
                        StatusTextBlock.Text = "▶️ [手动] 启动处理...";
                        _currentItemPasted = false;
                        _lastPastedStoreName = null;
                    }

                    // 3. ✅ 调用处理函数，传入 token
                    await ManualSmartProcessAsync(token);
                });
                await smartFlowUiTask;
            }
            catch (OperationCanceledException)
            {
                // 任务被新的按键取消了，这是正常现象
                System.Diagnostics.Debug.WriteLine("上一次搜索已被用户手动打断。");
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"❌ 错误: {ex.Message}";
            }
        }

        /// <summary>
        /// 安全检测：仅检测当前流程对应的应用，避免“微信流程误查企微”导致误停。
        /// </summary>
        private bool CheckForSecurityVerification(bool? expectedAppIsWework)
        {
            try
            {
                bool? appIsWework = expectedAppIsWework;

                if (!appIsWework.HasValue && _currentSelectedNode != null)
                {
                    string src = _currentSelectedNode.Source?.Trim() ?? string.Empty;
                    if (string.Equals(src, "企业微信", StringComparison.OrdinalIgnoreCase))
                    {
                        appIsWework = true;
                    }
                    else if (string.Equals(src, "微信", StringComparison.OrdinalIgnoreCase))
                    {
                        appIsWework = false;
                    }
                }

                if (!appIsWework.HasValue && _lastSearchWindowHandle != IntPtr.Zero)
                {
                    appIsWework = _lastSearchWindowIsWework;
                }

                if (!appIsWework.HasValue)
                {
                    System.Diagnostics.Debug.WriteLine("[安全检测] ℹ️ 无法确定当前应用，跳过安全检测");
                    return false;
                }

                string appName = appIsWework.Value ? "企业微信" : "微信";
                System.Diagnostics.Debug.WriteLine(
                    $"[安全检测] 开始: App={appName}, LastSearchHwnd={_lastSearchWindowHandle}, LastSearchIsWework={_lastSearchWindowIsWework}, " +
                    $"LastLayoutHwnd={_lastLayoutVerifiedHwnd}, LastLayoutIsWework={_lastLayoutVerifiedIsWework}");

                // 优先信任“刚通过布局验证”的窗口句柄，避免仅靠进程名误判
                if (TryTrustWindowHandleForSecurity(_lastLayoutVerifiedHwnd, _lastLayoutVerifiedIsWework, appIsWework.Value, "布局验证"))
                {
                    return false;
                }

                // 次级兜底：信任本轮搜索阶段命中的窗口句柄
                if (TryTrustWindowHandleForSecurity(_lastSearchWindowHandle, _lastSearchWindowIsWework, appIsWework.Value, "搜索阶段"))
                {
                    return false;
                }

                if (appIsWework.Value)
                {
                    return CheckAppSecurityVerification(new[] { "WXWork" }, "企业微信", "WeChatLogin");
                }

                string configuredWeChatName = _searchConfig?.WeChatProcessName?.Trim() ?? "Weixin";
                var weChatProcessCandidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    configuredWeChatName,
                    "Weixin",
                    "WeChat"
                };
                return CheckAppSecurityVerification(weChatProcessCandidates, "微信", null);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[安全检测] 异常: {ex.Message}");
                return false;
            }
        }

        private bool TryTrustWindowHandleForSecurity(IntPtr hwnd, bool? handleIsWework, bool expectedIsWework, string stageTag)
        {
            if (hwnd == IntPtr.Zero || !handleIsWework.HasValue || handleIsWework.Value != expectedIsWework)
            {
                return false;
            }

            bool valid = IsTargetChatWindow(hwnd, out string processName) &&
                         IsProcessNameForApp(processName, expectedIsWework);
            if (valid)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[安全检测] ✅ 复用{stageTag}句柄通过: Hwnd={hwnd}, Proc={processName}");
                return true;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[安全检测] ⚠️ {stageTag}句柄不可用: Hwnd={hwnd}, Proc={(string.IsNullOrWhiteSpace(processName) ? "<unknown>" : processName)}");
            return false;
        }

        private bool CheckAppSecurityVerification(IEnumerable<string> processNames, string appName, string? loginWindowClass)
        {
            var candidateNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (processNames != null)
            {
                foreach (var name in processNames)
                {
                    string normalized = NormalizeProcessNameCandidate(name);
                    if (!string.IsNullOrWhiteSpace(normalized))
                    {
                        candidateNames.Add(normalized);
                    }
                }
            }

            if (candidateNames.Count == 0)
            {
                candidateNames.Add(appName == "企业微信" ? "WXWork" : "Weixin");
            }

            var processMap = new Dictionary<int, Process>();
            System.Diagnostics.Debug.WriteLine($"[安全检测] {appName}进程候选: {string.Join("/", candidateNames)}");
            foreach (var name in candidateNames)
            {
                Process[] matched;
                try
                {
                    matched = Process.GetProcessesByName(name);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[安全检测] 获取进程列表失败: Name={name}, Error={ex.Message}");
                    continue;
                }

                System.Diagnostics.Debug.WriteLine($"[安全检测] 候选命中: Name={name}, Count={matched.Length}");
                foreach (var p in matched)
                {
                    if (!processMap.ContainsKey(p.Id))
                    {
                        processMap[p.Id] = p;
                    }
                }
            }
            var processes = processMap.Values.ToList();

            if (processes.Count == 0)
            {
                string msg = $"⛔ {appName}进程不存在，需要停止 (Candidates={string.Join("/", candidateNames)})";
                System.Diagnostics.Debug.WriteLine($"[安全检测] {msg}");
                DebugLogManager.Log("安全检测", msg);
                return true;
            }

            if (!string.IsNullOrWhiteSpace(loginWindowClass))
            {
                IntPtr loginWindowHandle = FindWindow(loginWindowClass, null);
                if (loginWindowHandle != IntPtr.Zero)
                {
                    string msg = $"⛔ 检测到登录窗口 ({loginWindowClass})，{appName}可能已退出登录";
                    System.Diagnostics.Debug.WriteLine($"[安全检测] {msg}");
                    DebugLogManager.Log("安全检测", msg);
                    return true;
                }
            }

            bool hasValidWindow = false;
            foreach (var process in processes)
            {
                try
                {
                    process.Refresh();
                    System.Diagnostics.Debug.WriteLine(
                        $"[安全检测] 进程明细: PID={process.Id}, Name={process.ProcessName}, MainHwnd={process.MainWindowHandle}, Title='{process.MainWindowTitle}'");
                    if (process.MainWindowHandle != IntPtr.Zero)
                    {
                        hasValidWindow = true;
                        System.Diagnostics.Debug.WriteLine($"[安全检测] ✅ {appName}窗口存在: PID={process.Id}, Title='{process.MainWindowTitle}'");
                        break;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[安全检测] 读取进程窗口失败: PID={process.Id}, Error={ex.Message}");
                }
            }

            if (!hasValidWindow)
            {
                string msg = $"⛔ {appName}进程存在但无有效窗口，可能已退出登录";
                System.Diagnostics.Debug.WriteLine($"[安全检测] {msg}");
                DebugLogManager.Log("安全检测", msg);
                return true;
            }

            return false;
        }

        private static string NormalizeProcessNameCandidate(string? rawName)
        {
            if (string.IsNullOrWhiteSpace(rawName))
            {
                return string.Empty;
            }

            string name = rawName.Trim();
            int slashIndex = Math.Max(name.LastIndexOf('\\'), name.LastIndexOf('/'));
            if (slashIndex >= 0 && slashIndex < name.Length - 1)
            {
                name = name[(slashIndex + 1)..];
            }

            if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                name = name[..^4];
            }

            return name.Trim();
        }

        private async Task ManualSmartProcessAsync(CancellationToken token)
        {
            // 1. 基础检查
            if (_currentSelectedNode == null) return;

            // ✅ [优化] 无群聊商家 或 强制轮询模式 -> 轻量级轮询搜索模式
            // ✅ [修改] 使用新的模式变量代替 CheckBox
            bool forcePolling = _isStoreMode;
            if (string.IsNullOrEmpty(_currentSelectedNode.GroupName) || forcePolling)
            {
                await PerformLightweightSearchAsync(NormalizeStoreNameForSearch(_currentSelectedNode.StoreName));
                return;
            }

            // 2. 🛡️ 净化环境：释放按键
            try
            {
                _inputBackend.KeyUp(InputKey.LeftControl);
                _inputBackend.KeyUp(InputKey.Enter);
            }
            catch { }

            // 3. ♻️ 调用核心逻辑 (传入 token)
            bool success = await SearchCurrentItemAsync(false, token);

            // 原来的轮询切换逻辑已移除，由 SmartAdvanceOrSearchAsync 在按键打断时接管
            if (!success && token.IsCancellationRequested)
            {
                return;
            }
        }
        /// <summary>
        /// ✅ [自动模式核心] 支持取消令牌 (CancellationToken)
        /// </summary>
        private async Task<bool> SearchCurrentItemAsync(bool isAutoMode = false, CancellationToken token = default)
        {
            // 🔍 [调试日志]
            System.Diagnostics.Debug.WriteLine($"\n[{DateTime.Now:HH:mm:ss.fff}] ============== [调试] 开始搜索流程 ==============");

            if (isAutoMode)
            {
                Interlocked.Exchange(ref _lastFailureReachedPasteStage, 0);
                _lastSearchWindowHandle = IntPtr.Zero;
                _lastSearchWindowIsWework = false;
            }

            Interlocked.Increment(ref _clipboardSearchGuard);

            // 1. 获取数据快照
            string storeName = null;
            string groupName = null;
            string source = null;
            SendStrategy strategy = SendStrategy.TextDirect;

            try
            {
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    if (_currentSelectedNode != null)
                    {
                        storeName = _currentSelectedNode.StoreName;
                        groupName = _currentSelectedNode.GroupName;
                        source = _currentSelectedNode.Source;
                        strategy = _currentSelectedNode.Strategy;
                    }
                });

                if (string.IsNullOrEmpty(storeName)) return false;
                string normalizedStoreName = NormalizeStoreNameForSearch(storeName);
                if (string.IsNullOrWhiteSpace(normalizedStoreName))
                {
                    normalizedStoreName = storeName.Trim();
                }

                string normalizedGroupName = groupName?.Trim();
                bool hasValidGroupName = !string.IsNullOrWhiteSpace(normalizedGroupName);

                // F1 自动化：强制只按群名搜索，不允许退化为商家名
                if (isAutoMode && !hasValidGroupName)
                {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        StatusTextBlock.Text = $"⏭️ 商家 '{normalizedStoreName}' 无群聊，跳过搜索。";
                    });
                    return false;
                }

                var snapshot = new
                {
                    StoreName = storeName,
                    GroupName = normalizedGroupName,
                    SearchText = hasValidGroupName ? normalizedGroupName : normalizedStoreName,
                    HasGroupName = hasValidGroupName,
                    IsWework = hasValidGroupName ? "企业微信".Equals(source, StringComparison.OrdinalIgnoreCase) : _isWeworkTurn
                };

                string appName = snapshot.IsWework ? "企业微信" : "微信";
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"🔍 正在 [{appName}] 搜索: {snapshot.SearchText}...");

                // ✅ [重构] 定义粘贴动作：根据 SendStrategy 三路分发
                Func<Task<bool>> performPasteAsync = async () =>
                {
                    switch (strategy)
                    {
                        case SendStrategy.FileExcel:
                            return await PasteExcelFileAsync(snapshot.StoreName, snapshot.IsWework, token);
                        case SendStrategy.TextSegmented:
                            return await PasteStoreInfoInSegmentsAsync(snapshot.StoreName, snapshot.IsWework, token, isAutoMode);
                        default:
                            return await PasteFullStoreInfoAsync(snapshot.StoreName, snapshot.IsWework, token);
                    }
                };

                // ✅ 检查取消
                if (token.IsCancellationRequested) return false;

                // --------------------------------------------------------
                // 🚀 1. 极速模式 (略微简化展示，关键是加入 ActivateWindow 的等待)
                // --------------------------------------------------------
                if (snapshot.HasGroupName && snapshot.SearchText == _lastEnteredGroupName)
                {
                    // ... (省略部分日志) ...

                    bool activateResult = false;
                    if (_lastChatWindowHandle != IntPtr.Zero)
                        activateResult = RobustActivateWindow(_lastChatWindowHandle);
                    else
                        activateResult = RobustActivateWindow(GetForegroundWindow());

                    // ✅ 延时支持取消
                    try { await Task.Delay(100, token); } catch (TaskCanceledException) { return false; }

                    IntPtr checkHwnd = GetForegroundWindow();
                    string titleText = await _screenshotHelper.GetWeChatWindowTitleTextAsync(checkHwnd, snapshot.IsWework);

                    if (_screenshotHelper.IsFuzzyMatch(snapshot.SearchText, titleText))
                    {
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"⚡ [极速] 验证通过，直接发送。");
                        return await performPasteAsync();
                    }
                    else
                    {
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "⚠️ 窗口不符，转常规搜索...");
                        _lastEnteredGroupName = null;
                        _lastChatWindowHandle = IntPtr.Zero;
                    }
                }

                // ============================================================
                // 🔍 2. 常规搜索模式
                // ============================================================

                // ✅ 检查取消
                if (token.IsCancellationRequested) return false;

                IntPtr mainHwnd = GetForegroundWindow();
                if (!RobustActivateWindow(mainHwnd)) return false;

                // 执行搜索
                bool autoSearchSuccess = await _searchHelper.SearchInAppAsync(snapshot.SearchText, snapshot.IsWework, token);
                if (!autoSearchSuccess) return false;

                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "👀 [自动] 验证搜索列表...");

                // ✅ 延时支持取消
                try { await Task.Delay(200, token); } catch (TaskCanceledException) { return false; }

                // --------------------------------------------------------
                // 🔥 步骤 A: 搜索列表 OCR 验证
                // --------------------------------------------------------
                // 🔧 [修复] 优先使用 SearchHelper 缓存的已验证窗口句柄，
                // 而非 GetForegroundWindow()，避免获取到非目标窗口（如搜索弹窗/子窗口）
                IntPtr searchHwnd = _searchHelper.TryGetAppWindowHandle(snapshot.IsWework);
                if (searchHwnd == IntPtr.Zero)
                    searchHwnd = GetForegroundWindow();
                System.Diagnostics.Debug.WriteLine($"🔍 [DEBUG_TRACE] searchHwnd 来源: {(searchHwnd == GetForegroundWindow() ? "Foreground" : "Cache")}, Hwnd={searchHwnd}");
                if (isAutoMode && searchHwnd != IntPtr.Zero)
                {
                    _lastSearchWindowHandle = searchHwnd;
                    _lastSearchWindowIsWework = snapshot.IsWework;
                    System.Diagnostics.Debug.WriteLine(
                        $"[AutoLoop] 已记录搜索阶段窗口: App={(snapshot.IsWework ? "企业微信" : "微信")}, Hwnd={searchHwnd}");
                }
                bool isListMatch = false;

      
                System.Drawing.Point? clickPos = null;
                System.Drawing.Rectangle? matchedSearchRect = null;

                for (int i = 0; i < 3; i++)
                {
                    // ✅ 循环中检查取消
                    if (token.IsCancellationRequested) return false;

                    System.Diagnostics.Debug.WriteLine($"🔍 [DEBUG_TRACE] 步骤AB-验证与定位(第{i+1}次), AutoMode={isAutoMode}");
                    try 
                    {
                        // 🛡️ [防御] 每次验证前确保窗口是前台
                        if (searchHwnd != GetForegroundWindow())
                        {
                            System.Diagnostics.Debug.WriteLine("⚠️ [DEBUG_TRACE] 验证前发现窗口失焦，尝试重新激活...");
                            RobustActivateWindow(searchHwnd);
                            await Task.Delay(200, token);
                        }

                        // 调用合并后的方法：验证 + 获取坐标
                        var result = await _screenshotHelper.FindAndVerifySearchResultAsync(searchHwnd, snapshot.SearchText, snapshot.IsWework);
                        
                        if (result.success && result.clickPoint.HasValue)
                        {
                            isListMatch = true;
                            clickPos = result.clickPoint;
                            matchedSearchRect = result.matchedScreenBBox;
                            System.Diagnostics.Debug.WriteLine($"✅ [DEBUG_TRACE] 验证并定位成功: {clickPos}");
                            break;
                        }
                        
                        System.Diagnostics.Debug.WriteLine($"🔍 [DEBUG_TRACE] 步骤AB尝试失败");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"💥 [DEBUG_TRACE] 步骤AB 发生异常: {ex.ToString()}");
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"💥 验证过程异常: {ex.Message}");
                        return false;
                    }

                    if (i < 2)
                    {
                         // 🔧 [修复] 验证失败时，清除窗口缓存并重新获取正确的窗口句柄
                         _searchHelper.ClearWindowCache(snapshot.IsWework ? "企业微信" : "微信");
                         IntPtr refreshedHwnd = _searchHelper.TryGetAppWindowHandle(snapshot.IsWework);
                         if (refreshedHwnd != IntPtr.Zero && refreshedHwnd != searchHwnd)
                         {
                             System.Diagnostics.Debug.WriteLine($"🔄 [DEBUG_TRACE] 刷新 searchHwnd: {searchHwnd} -> {refreshedHwnd}");
                             searchHwnd = refreshedHwnd;
                         }
                         try { await Task.Delay(800, token); } catch (TaskCanceledException) { return false; }
                    }
                }

                if (!isListMatch || !clickPos.HasValue || !matchedSearchRect.HasValue)
                {
                    System.Diagnostics.Debug.WriteLine($"❌ [DEBUG_TRACE] 验证或定位最终失败，准备退出。AutoMode={isAutoMode}");
                    
                    // 🚨 [关键修复] 如果验证彻底失败，说明可能截错图或者窗口状态异常
                    _searchHelper.ClearWindowCache(snapshot.IsWework ? "企业微信" : "微信");
                    System.Diagnostics.Debug.WriteLine("🧹 [DEBUG_TRACE] 已清除窗口缓存，防止死循环。");

                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ 搜索列表超时或未找到目标，停止。");
                    return false;
                }

                System.Diagnostics.Debug.WriteLine($"✅ [DEBUG_TRACE] 流程继续... AutoMode={isAutoMode}");

                // --------------------------------------------------------
                // 🔥 步骤 B (后半): 鼠标点击群聊
                // --------------------------------------------------------
                
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "🔍 [OCR] 正在点击群聊...");
                
                IntPtr hwndForClick = searchHwnd;
                // 复用 clickPos, 无需再次查找
                
                bool enteredSuccess = false;
                string cleanTarget = snapshot.SearchText.Replace(" ", "").ToLowerInvariant();

                    if (clickPos.HasValue)
                    {
                        int clickX = clickPos.Value.X;
                        int clickY = clickPos.Value.Y;

                        // 🔍 [Debug] 获取窗口位置用于安全检查
                        RECT debugRect;
                        GetWindowRect(hwndForClick, out debugRect);
                        System.Diagnostics.Debug.WriteLine($"🔍 [DEBUG_TRACE] 准备点击: ({clickX}, {clickY}), 窗口范围: [{debugRect.Left},{debugRect.Top} - {debugRect.Right},{debugRect.Bottom}]");

                        // ⚠️ [Debug] 警告：点击位置过于靠近顶部 (可能是标题栏/关闭按钮)
                        if (clickY - debugRect.Top < 50)
                        {
                             System.Diagnostics.Debug.WriteLine($"⚠️⚠️ [DEBUG_TRACE] 严重警告！点击位置 ({clickY}) 距离窗口顶部 ({debugRect.Top}) 过近 (<50px)，可能误触标题栏！");
                        }
                    
                    // 🎯 检查是否为特殊坐标：(-1, -1) 表示"最常使用"，直接回车
                    if (clickX == -1 && clickY == -1)
                    {
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "🎯 [OCR] 检测到'最常使用'，直接回车进入...");
                        System.Diagnostics.Debug.WriteLine("🎯 [MainWindow] 检测到特殊坐标(-1,-1)，表示'最常使用'，直接按回车");
                        
                        // 直接按回车进入
                        _inputBackend.KeyPress(InputKey.Enter);
                        try { await Task.Delay(300, token); } catch (TaskCanceledException) { return false; }
                    }

                    else
                    {
                        // 通用：正常坐标，若二次识别失败则重定位新框并重试，不在首轮直接中断流程
                        bool clickCompleted = false;
                        const int maxClickAttempts = 3;
                        var currentMatchedRect = matchedSearchRect.Value;
                        // 🚨 [修改] 移除 6 像素误差，严格校验 (tolerance = 0)
                        const int pointInRectTolerance = 0;

                        for (int clickAttempt = 0; clickAttempt < maxClickAttempts && !clickCompleted; clickAttempt++)
                        {
                            try
                            {
                                if (clickAttempt == 0)
                                {
                                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"🖱️ 拟人化移动: ({clickX}, {clickY})");
                                }
                                else
                                {
                                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"🔄 [OCR] 二次未通过，重识别群聊 ({clickAttempt + 1}/{maxClickAttempts})...");
                                    var refreshResult = await _screenshotHelper.FindAndVerifySearchResultAsync(hwndForClick, snapshot.SearchText, snapshot.IsWework);
                                    if (!refreshResult.success || !refreshResult.clickPoint.HasValue)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"⚠️ [DEBUG_TRACE] 第{clickAttempt + 1}次重识别未找到可点击目标");
                                        try { await Task.Delay(120, token); } catch (TaskCanceledException) { return false; }
                                        continue;
                                    }

                                    clickX = refreshResult.clickPoint.Value.X;
                                    clickY = refreshResult.clickPoint.Value.Y;
                                    if (!refreshResult.matchedScreenBBox.HasValue)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"⚠️ [DEBUG_TRACE] 第{clickAttempt + 1}次重识别缺少目标框信息");
                                        continue;
                                    }
                                    currentMatchedRect = refreshResult.matchedScreenBBox.Value;
                                    System.Diagnostics.Debug.WriteLine($"🔄 [DEBUG_TRACE] 重识别新坐标: ({clickX}, {clickY})");
                                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"🖱️ 重识别后移动: ({clickX}, {clickY})");
                                }

                                await MouseHelper.MoveMouseSmoothlyAsync(clickX, clickY, 95);
                                try { await Task.Delay(25, token); } catch (TaskCanceledException) { return false; }

                                var movedPoint = MouseHelper.GetCursorPosition();
                                System.Diagnostics.Debug.WriteLine($"🔍 [DEBUG_TRACE] 鼠标已移动到: ({movedPoint.X}, {movedPoint.Y})，开始二次框内校验...");

                                // 📸 [新增] 保存校验时的调试图片 (点击前) - 全窗口版
                                try
                                {
                                    // 获取目标窗口的 Rect
                                    RECT windowRect;
                                    if (GetWindowRect(hwndForClick, out windowRect))
                                    {
                                        int width = windowRect.Right - windowRect.Left;
                                        int height = windowRect.Bottom - windowRect.Top;

                                        if (width > 0 && height > 0)
                                        {
                                            using (var bmp = new System.Drawing.Bitmap(width, height))
                                            using (var g = System.Drawing.Graphics.FromImage(bmp))
                                            {
                                                // 截取整个窗口
                                                g.CopyFromScreen(windowRect.Left, windowRect.Top, 0, 0, new System.Drawing.Size(width, height));
                                                
                                                // 画出目标框 (相对于窗口，蓝色)
                                                // currentMatchedRect 是屏幕坐标
                                                var relRect = new System.Drawing.Rectangle(
                                                    currentMatchedRect.X - windowRect.Left, 
                                                    currentMatchedRect.Y - windowRect.Top, 
                                                    currentMatchedRect.Width, 
                                                    currentMatchedRect.Height);
                                                
                                                g.DrawRectangle(new System.Drawing.Pen(System.Drawing.Color.Blue, 3), relRect);
                                                
                                                // 画出鼠标位置 (相对于窗口，红色)
                                                var relPoint = new System.Drawing.Point(
                                                    movedPoint.X - windowRect.Left, 
                                                    movedPoint.Y - windowRect.Top);
                                                
                                                g.FillEllipse(System.Drawing.Brushes.Red, relPoint.X - 4, relPoint.Y - 4, 8, 8);
                                                
                                                // 保存
                                                // [已禁用] Debug_Yolo 调试目录相关代码
                                                // string debugDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Debug_Yolo", DateTime.Now.ToString("yyyyMMdd"));
                                                // System.IO.Directory.CreateDirectory(debugDir);
                                                // string debugPath = System.IO.Path.Combine(debugDir, $"Verify_Full_{DateTime.Now:HHmmss_fff}.png");
                                                //bmp.Save(debugPath, System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DebugImage] 全窗口保存失败: {ex.Message}");
                                }

                                bool secondPass =
                                    movedPoint.X >= currentMatchedRect.Left - pointInRectTolerance &&
                                    movedPoint.X <= currentMatchedRect.Right + pointInRectTolerance &&
                                    movedPoint.Y >= currentMatchedRect.Top - pointInRectTolerance &&
                                    movedPoint.Y <= currentMatchedRect.Bottom + pointInRectTolerance;

                                if (!secondPass)
                                {
                                    System.Diagnostics.Debug.WriteLine(
                                        $"⚠️ [DEBUG_TRACE] 第{clickAttempt + 1}次二次校验失败: 点({movedPoint.X},{movedPoint.Y}) 不在框[{currentMatchedRect.Left},{currentMatchedRect.Top},{currentMatchedRect.Right},{currentMatchedRect.Bottom}]，准备重识别+OCR");
                                    continue;
                                }

                                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"🖱️ 框内校验通过，执行点击: ({movedPoint.X}, {movedPoint.Y})");
                                await MouseHelper.LeftClickCurrentAsync();
                                System.Diagnostics.Debug.WriteLine("🔍 [DEBUG_TRACE] LeftClickCurrentAsync 完成");
                                clickCompleted = true;
                            }
                            catch (Exception clickEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"💥 [DEBUG_TRACE] 鼠标点击过程异常(第{clickAttempt + 1}次): {clickEx.Message}");
                            }
                        }

                        if (!clickCompleted)
                        {
                            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ [OCR] 多次重识别仍未通过点击校验");
                            _lastEnteredGroupName = null;
                            return false;
                        }

                        // 等待窗口切换 (优化 400 → 300ms)
                        try { await Task.Delay(300, token); } catch (TaskCanceledException) { return false; }
                    }
                    
                    System.Diagnostics.Debug.WriteLine("🔍 [DEBUG_TRACE] 延迟结束，准备调用 GetWeChatWindowTitleTextAsync...");

                    // 验证是否进入了正确的群聊
                    IntPtr chatHwnd = GetForegroundWindow();
                    string rawTitle = "";
                    
                    // 🔄 [增强] 标题验证重试机制 (最多 3 次，每次间隔 500ms)
                    // 原因：点击后 UI 刷新需要时间，严苛的场景验证(GetWeChatWindowTitleTextAsync)可能会因为
                    // "新页面还未渲染出的 Label_ChatInfo" 而直接返回 null。给它一点时间加载。
                    for (int v = 0; v < 3; v++)
                    {
                        try
                        {
                             rawTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(chatHwnd, snapshot.IsWework);
                             System.Diagnostics.Debug.WriteLine($"🔍 [DEBUG_TRACE] GetWeChatWindowTitleTextAsync (尝试 {v+1}/3) 返回: '{rawTitle}'");
                             
                             // 如果获取到了标题，且不为空，直接跳出
                             if (!string.IsNullOrEmpty(rawTitle)) break;
                        }
                        catch (Exception titleEx)
                        {
                             System.Diagnostics.Debug.WriteLine($"💥 [DEBUG_TRACE] 获取窗口标题异常: {titleEx.Message}");
                        }
                        
                        // 如果还没成功，等待 UI 渲染
                        if (v < 2) 
                        {
                            System.Diagnostics.Debug.WriteLine("⏳ [DEBUG_TRACE] 标题验证未通过/为空，等待 UI 渲染...");
                            try { await Task.Delay(500, token); } catch (TaskCanceledException) { return false; }
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(rawTitle) && _screenshotHelper.IsFuzzyMatch(snapshot.SearchText, rawTitle))
                    {
                        enteredSuccess = true;
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"✅ [OCR] 成功进入: {rawTitle}");
                        
                        if (snapshot.HasGroupName)
                        {
                            _lastEnteredGroupName = snapshot.SearchText;
                            _lastChatWindowHandle = chatHwnd;
                        }
                    }
                    else
                    {
                        // ❌ OCR 点击后标题不匹配，直接返回失败（不再使用键盘导航，减少侵入操作）
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"❌ [OCR] 标题不匹配: '{rawTitle}'，搜索失败");
                        
                        // ⚠️ [修复] 移除 ESC 键，因为在企业微信中 ESC 可能会直接关闭窗口！
                        
                        _lastEnteredGroupName = null;
                        return false;
                    }
                }
                else
                {
                    // ❌ OCR 未找到任何可点击的群聊位置
                    // 这通常意味着搜索结果只有"搜索网络结果"分类，没有本地匹配的群聊
                    // 不应该走键盘导航，因为那会点击"搜索网络结果"下的内容
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ [OCR] 未找到群聊，搜索结果可能只有网络搜索");
                    System.Diagnostics.Debug.WriteLine("❌ [MainWindow] OCR 返回 null，可能只有搜索网络结果，跳过键盘导航直接返回失败");
                    _lastEnteredGroupName = null;
                    return false;
                }

                // --------------------------------------------------------
                // 🔥 步骤 C: 键盘导航兜底方案 (如果 OCR 点击成功但验证失败)
                // --------------------------------------------------------
                if (!enteredSuccess)
                {
                    // 键盘导航仅用于 OCR 点击了但验证群聊名称失败的情况
                    // 默认从 1 (按1次Down) 开始，以避开默认的 "搜索网络"
                    int loopStart = 1;

                    // 循环次数: 往后试 3 个 (减少侵入性操作，避免微信卡死)
                    int maxAttempts = loopStart + 3; 

                    for (int attemptIndex = loopStart; attemptIndex <= maxAttempts; attemptIndex++)
                    {
                        if (token.IsCancellationRequested) return false;

                        // 1. 如果不是第一次 (即上次尝试失败)，必须重新触发搜索
                        // (因为微信 enter 进群后，搜索面板通常会消失)
                        if (attemptIndex > loopStart)
                        {
                            StatusTextBlock.Text = $"🔄 [第{attemptIndex}项] 重新搜索...";
                            // 重新搜
                            if (!await _searchHelper.SearchInAppAsync(snapshot.SearchText, snapshot.IsWework, token))
                            {
                                StatusTextBlock.Text = "❌ 重试搜索失败";
                                return false;
                            }
                            // 等待列表渲染
                            try { await Task.Delay(300, token); } catch { return false; }
                        }

                        // 2. 键盘导航：按 Down 键 attemptIndex 次
                        StatusTextBlock.Text = $"⌨️ [第{attemptIndex}项] 键盘选中...";
                        
                        for (int k = 0; k < attemptIndex; k++)
                        {
                            _inputBackend.KeyPress(InputKey.DownArrow);
                            try { await Task.Delay(50, token); } catch { return false; }
                        }

                        // 3. 回车进入
                        _inputBackend.KeyPress(InputKey.Enter);
                        
                        // 等待响应
                        try { await Task.Delay(300, token); } catch { return false; }

                        // 4. 验证标题
                        IntPtr chatHwnd = GetForegroundWindow();
                        
                        // 🚨 [陷阱检测] 如果现在的窗口不是微信主界面，说明可能不幸踩中了 "Search Web"
                        // 检测方法：获取当前 Process 名字，如果是浏览器 (chrome, edge) 或者 new search window
                        // 这里简化处理：直接 verify title，如果 title 不对，就会 continue
                        
                        if (chatHwnd == IntPtr.Zero) continue;

                        string rawTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(chatHwnd, snapshot.IsWework);
                        if (string.IsNullOrEmpty(rawTitle))
                        {
                            try { await Task.Delay(300, token); } catch (TaskCanceledException) { return false; }
                            rawTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(chatHwnd, snapshot.IsWework);
                        }

                        string cleanTitle = System.Text.RegularExpressions.Regex.Replace(rawTitle ?? "", @"\(\d+.*?\)|（\d+.*?）|\(外部\)|（外部）|\s+", "").ToLowerInvariant();
                        
                        bool isMatch = false;
                        if (_screenshotHelper.IsFuzzyMatch(snapshot.SearchText, rawTitle)) isMatch = true;
                        else if (cleanTitle.Contains(cleanTarget) || (cleanTarget.Contains(cleanTitle) && cleanTitle.Length > 2)) isMatch = true;

                        if (isMatch)
                        {
                            enteredSuccess = true;
                            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"✅ [第{attemptIndex}项] 找到目标: {rawTitle}");
                            
                            if (snapshot.HasGroupName)
                            {
                                _lastEnteredGroupName = snapshot.SearchText;
                                _lastChatWindowHandle = chatHwnd;
                            }
                            break; 
                        }
                        else
                        {
                            StatusTextBlock.Text = $"⚠️ [第{attemptIndex}项] 标题不符({rawTitle})，重试...";
                            
                            // 🛠️ [容错] 如果不幸打开了弹窗，尝试按 ESC 关闭，以便下次搜索能正常进行
                            // (特别是搜索页面)
                            _inputBackend.KeyPress(InputKey.Escape);
                            try { await Task.Delay(100, token); } catch (TaskCanceledException) { return false; }
                        }
                    } // End Loop
                } // End if (!enteredSuccess)

                if (enteredSuccess)
                {
                    // 粘贴前最后检查一次
                    if (token.IsCancellationRequested) return false;
                    bool pasteSuccess = await performPasteAsync();
                    if (isAutoMode && !pasteSuccess)
                    {
                        Interlocked.Exchange(ref _lastFailureReachedPasteStage, 1);
                    }
                    return pasteSuccess;
                }
                else
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ OCR 和键盘方案均未能找到目标群聊。");
                    _lastEnteredGroupName = null;
                    return false;
                }
            }
            catch (OperationCanceledException)
            {
                // 任务被取消，静默返回 false
                return false;
            }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"💥 流程异常: {ex.Message}");
                return false;
            }
            finally
            {
                int after = Interlocked.Decrement(ref _clipboardSearchGuard);
                if (after < 0)
                {
                    Interlocked.Exchange(ref _clipboardSearchGuard, 0);
                }
            }
        }

        private async Task ManualSmartProcessAsync()
        {
            // 1. 基础检查
            if (_currentSelectedNode == null) return;

            // ✅ [优化] 无群聊商家 -> 轻量级轮询搜索模式
            // 仅在微信/企业微信之间循环执行 Ctrl+F 搜索粘贴，让用户自己观察结果并手动进入
            if (string.IsNullOrEmpty(_currentSelectedNode.GroupName))
            {
                await PerformLightweightSearchAsync(NormalizeStoreNameForSearch(_currentSelectedNode.StoreName));
                return;
            }

            // 2. 🛡️ 净化环境：释放按键
            // 这是手动模式成功的关键，防止物理按键（Ctrl/Enter）干扰自动化的 SearchCurrentItemAsync
            try
            {
                _inputBackend.KeyUp(InputKey.LeftControl);
                _inputBackend.KeyUp(InputKey.Enter);
            }
            catch { }


            // 3. ♻️ 直接复用自动化的核心逻辑！
            // 核心优势：F1 怎么跑，这里就怎么跑。包含完整的：
            // 找窗口 -> 激活 -> 搜索(含防抖) -> OCR验证列表 -> 回车 -> OCR验证标题 -> 粘贴
            bool success = await SearchCurrentItemAsync(false, CancellationToken.None); // isAutoMode = false
        }


        // MainWindow.xaml.cs

        /// <summary>
        /// ✅ [轻量级轮询搜索] 无群聊商家专用
        /// 仅在微信/企业微信之间循环执行 Ctrl+F 搜索粘贴，不做 OCR 验证和自动进入群聊
        /// 用户自己观察搜索结果并手动选择进入
        /// </summary>
        private async Task PerformLightweightSearchAsync(string searchText)
        {
            searchText = NormalizeStoreNameForSearch(searchText);
            if (string.IsNullOrWhiteSpace(searchText))
            {
                StatusTextBlock.Text = "⚠️ 搜索关键词为空，已跳过。";
                return;
            }

            // 1. 🛡️ 释放物理按键，防止干扰
            try
            {
                _inputBackend.KeyUp(InputKey.LeftControl);
                _inputBackend.KeyUp(InputKey.Enter);
            }
            catch { }

            await Task.Delay(50);

            // 2. 确定本次搜索的目标 APP
            bool isWework = _isWeworkTurn;
            string appName = isWework ? "企业微信" : "微信";
            
            // 🔍 调试日志
            System.Diagnostics.Debug.WriteLine($"[轮询搜索] 当前轮次: {appName}, _isWeworkTurn={_isWeworkTurn}");
            
            StatusTextBlock.Text = $"🔍 [轮询] 正在 {appName} 搜索: {searchText}...";

            // 3. 🔧 [修复] 每次轮询搜索前强制清除窗口缓存，防止使用已过期的句柄
            // （企业微信关闭窗口后进程仍在托盘运行，旧句柄能通过所有验证但实际已失效）
            _searchHelper.ClearWindowCache(appName);

            // 4. 调用 SearchHelper 执行搜索操作（Ctrl+F + 粘贴），不做后续验证
            bool searchSuccess = await _searchHelper.SearchInAppAsync(searchText, isWework);

            if (searchSuccess)
            {
                // 4a. 搜索成功，正常切换轮次
                _isWeworkTurn = !_isWeworkTurn;
                string nextApp = _isWeworkTurn ? "企业微信" : "微信";
                System.Diagnostics.Debug.WriteLine($"[轮询搜索] 搜索完成, 下次轮询: {nextApp}, _isWeworkTurn={_isWeworkTurn}");
                StatusTextBlock.Text = $"✅ [轮询] 已在 {appName} 搜索完成，请手动选择群聊。下次轮询: {nextApp}";
            }
            else
            {
                // 4b. 搜索失败（窗口找不到等），不切换轮次，下次仍尝试同一个 APP
                System.Diagnostics.Debug.WriteLine($"[轮询搜索] {appName} 搜索失败，不切换轮次, _isWeworkTurn={_isWeworkTurn}");
                StatusTextBlock.Text = $"⚠️ [轮询] {appName} 搜索失败（窗口未找到），下次仍尝试 {appName}";
            }
        }

        /// <summary>
        /// ✅ [Ctrl+' 专用] 仅 OCR 识别当前窗口群名，不粘贴不发送
        /// 沿用三倍放大等处理逻辑
        /// </summary>
        private async Task OcrRecognizeGroupNameOnlyAsync()
        {
            // 1. 基础检查
            if (_currentSelectedNode == null)
            {
                StatusTextBlock.Text = "⚠️ 请先选择一个商家";
                return;
            }

            string storeName = _currentSelectedNode.StoreName;
            StatusTextBlock.Text = "👁️ 正在 OCR 识别群名...";

            // 2. 识别当前窗口身份
            IntPtr currentHwnd = GetForegroundWindow();
            string currentClassName = GetWindowClass(currentHwnd);
            bool isWework = false;

            if (currentClassName == _searchConfig.WeworkWindowClassName)
            {
                isWework = true;
            }
            else if (currentClassName != _searchConfig.WechatWindowClassName)
            {
                // 兜底：通过进程名判断
                try
                {
                    GetWindowThreadProcessId(currentHwnd, out uint pid);
                    if (pid > 0)
                    {
                        using (var p = Process.GetProcessById((int)pid))
                        {
                            if (p.ProcessName.Equals("WXWork", StringComparison.OrdinalIgnoreCase))
                            {
                                isWework = true;
                            }
                        }
                    }
                }
                catch { }
            }

            string appName = isWework ? "企业微信" : "微信";
            StatusTextBlock.Text = $"👁️ 正在 OCR [{appName}] 群名...";

            // 3. 执行 OCR（沿用 ScreenshotHelper 的三倍放大等处理逻辑）
            try
            {
                string recognizedTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(currentHwnd, isWework);

                // 如果第一次没识别到，尝试反向识别一次
                if (string.IsNullOrWhiteSpace(recognizedTitle))
                {
                    recognizedTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(currentHwnd, !isWework);
                    if (!string.IsNullOrWhiteSpace(recognizedTitle))
                    {
                        isWework = !isWework;
                        appName = isWework ? "企业微信" : "微信";
                    }
                }

                if (!string.IsNullOrWhiteSpace(recognizedTitle) && recognizedTitle.Length > 1)
                {
                    // ✅ 识别成功：保存并更新
                    string sourceToSave = isWework ? "企业微信" : "微信";
                    UpdateBusInfo(storeName, recognizedTitle, sourceToSave);
                    StatusTextBlock.Text = $"✅ [OCR] 已识别并保存 [{sourceToSave}]: {recognizedTitle}";
                }
                else
                {
                    StatusTextBlock.Text = "⚠️ [OCR] 未识别到有效群名";
                }
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"❌ [OCR] 识别失败: {ex.Message}";
            }
        }

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
            SendStrategy strategy = _currentSelectedNode.Strategy;

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
                // ✅ 兜底修复：针对 WeChat 4.0 (Qt) 类名变动，增加从进程名判断
                bool processIdentified = false;
                try
                {
                    GetWindowThreadProcessId(currentHwnd, out uint pid);
                    if (pid > 0)
                    {
                        using (var p = Process.GetProcessById((int)pid))
                        {
                            if (p.ProcessName.Equals(_searchConfig.WeChatProcessName, StringComparison.OrdinalIgnoreCase))
                            {
                                isWework = false;
                                processIdentified = true;
                                StatusTextBlock.Text = $"🤖 检测到：微信 (Qt进程: {p.ProcessName})";
                            }
                            else if (p.ProcessName.Equals("WXWork", StringComparison.OrdinalIgnoreCase))
                            {
                                isWework = true;
                                processIdentified = true;
                                StatusTextBlock.Text = "🤖 检测到：企业微信 (进程名)";
                            }
                        }
                    }
                }
                catch { /* 忽略进程获取权限错误 */ }

                if (!processIdentified)
                {
                    // 兜底：模糊匹配
                    if (currentClassName.Contains("WeWork") || currentClassName.Contains("WXWork"))
                    {
                        isWework = true;
                        StatusTextBlock.Text = "🤖 检测到：企业微信 (模糊匹配)";
                    }
                    else
                    {
                         // 默认为微信，不做处理
                    }
                }
            }

            // ============================================================
            // 3. 执行核心动作 (传入识别到的 isWework)
            // ============================================================
            bool actionSuccess = false;

            // ✅ [修复] 如果选中了子节点（分段或单行），直接发送 RawData
            // 排除 FileExcel 模式（通常期望发送文件或全量）
            if (!string.IsNullOrEmpty(_currentSelectedNode.RawData) && 
                _currentSelectedNode.Strategy != SendStrategy.FileExcel)
            {
                 Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "📤 正在发送选中片段...");
                 
                 // 先写入剪贴板
                 if (await SetClipboardWithRetryAsync(_currentSelectedNode.RawData))
                 {
                     // 调用通用发送逻辑 (false 表示不是文件)
                     actionSuccess = await PasteAndVerifySendAsync(_currentSelectedNode.RawData, false);
                 }
            }
            else
            {
                switch (strategy)
                {
                case SendStrategy.FileExcel:
                    actionSuccess = await PasteExcelFileAsync(storeName, isWework);
                    break;
                case SendStrategy.TextSegmented:
                    actionSuccess = await PasteStoreInfoInSegmentsAsync(storeName, isWework, CancellationToken.None);
                    break;
                default:
                    actionSuccess = await PasteFullStoreInfoAsync(storeName, isWework);
                    break;
            }
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

        private static bool IsControlKeyPressed()
        {
            return (GetAsyncKeyState(VK_CONTROL_KEY) & 0x8000) != 0 ||
                   (GetAsyncKeyState(VK_LCONTROL_KEY) & 0x8000) != 0 ||
                   (GetAsyncKeyState(VK_RCONTROL_KEY) & 0x8000) != 0;
        }

        private async Task EnsureCtrlReleasedBeforeHotkeySendAsync()
        {
            // 释放注入层 Ctrl，避免发送 Enter 时被识别为 Ctrl+Enter
            try
            {
                _inputBackend.KeyUp(InputKey.LeftControl);
                _inputBackend.KeyUp(InputKey.RightControl);
            }
            catch
            {
                // 忽略输入模拟器异常
            }

            // 等待用户物理按键松开，避免“有时发不出去”
            for (int i = 0; i < 16; i++)
            {
                if (!IsControlKeyPressed())
                {
                    break;
                }

                await Task.Delay(5);
            }
        }

        /// <summary>
        /// 在快捷键处理完成后，如果用户仍然物理按住 Ctrl 键，
        /// 则重新注入 Ctrl 键按下状态，确保后续 Ctrl+W/S 等快捷键能连贯响应。
        /// 使用低级键盘钩子追踪的物理状态（不受 SendInput 注入影响）。
        /// </summary>
        private void RestoreCtrlKeyIfPhysicallyHeld()
        {
            bool physicalState = _physicalCtrlPressed;
            bool asyncState = IsControlKeyPressed();
            Debug.WriteLine($"[RestoreCtrl] 物理状态={physicalState}, GetAsyncKeyState={asyncState}");

            if (physicalState)
            {
                try
                {
                    _inputBackend.KeyDown(InputKey.LeftControl);
                    Debug.WriteLine("[RestoreCtrl] ✅ 已恢复 Ctrl 注入");

                    // 启动异步清理：用户松开物理 Ctrl 后自动发送配套 KeyUp，防止 Ctrl 卡住
                    _ = Task.Run(async () =>
                    {
                        for (int i = 0; i < 200; i++) // 最多等 10 秒
                        {
                            await Task.Delay(50);
                            if (!_physicalCtrlPressed)
                            {
                                try
                                {
                                    _inputBackend.KeyUp(InputKey.LeftControl);
                                    Debug.WriteLine("[RestoreCtrl] 🧹 用户已松开，自动清理注入的 Ctrl");
                                }
                                catch { }
                                return;
                            }
                        }
                        // 超时兜底：强制释放
                        try
                        {
                            _inputBackend.KeyUp(InputKey.LeftControl);
                            Debug.WriteLine("[RestoreCtrl] ⏰ 超时兜底释放 Ctrl");
                        }
                        catch { }
                    });
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[RestoreCtrl] ❌ 恢复失败: {ex.Message}");
                }
            }
            else
            {
                Debug.WriteLine("[RestoreCtrl] 用户已松开 Ctrl，不恢复");
            }
        }

        private async Task HandleCtrlSpaceHotkeyAsync()
        {
            await EnsureCtrlReleasedBeforeHotkeySendAsync();

            try
            {
                if (_currentSelectedNode == null)
                {
                    StatusTextBlock.Text = "⚠️ 请先选择一个商家或子项";
                    return;
                }

                if (!TryResolveRootNode(_currentSelectedNode, out TreeViewNode rootNode))
                {
                    StatusTextBlock.Text = "⚠️ 未能定位当前所属主列表项";
                    return;
                }

                if (string.IsNullOrWhiteSpace(rootNode.StoreName) || rootNode.StoreName == "FAIL_SEPARATOR")
                {
                    StatusTextBlock.Text = "⚠️ 当前项不可执行 Ctrl+Space";
                    return;
                }

                bool autoSend = AutoSendCheckBox.IsChecked == true;
                bool isSegmentedRoot = rootNode.Strategy == SendStrategy.TextSegmented &&
                                       rootNode.Children != null &&
                                       rootNode.Children.Count > 0;

                string seqPrefix = "";
                int index = _flatNodeList.IndexOf(rootNode);
                if (index >= 0)
                {
                    seqPrefix = $"[{index + 1}] ";
                }

                if (isSegmentedRoot)
                {
                    OsdWindow.ShowMessage(rootNode.StoreName, seqPrefix, rootNode.GroupName);
                    await HandleCtrlSpaceSegmentedAsync(rootNode, _currentSelectedNode, autoSend);
                    return;
                }

                if (!ReferenceEquals(_currentSelectedNode, rootNode))
                {
                    SuppressNextSelectionOsd();
                }
                SelectNodeWithoutCopy(rootNode);
                OsdWindow.ShowMessage(rootNode.StoreName, seqPrefix, rootNode.GroupName);

                bool success = rootNode.Strategy == SendStrategy.FileExcel
                    ? await PasteExcelForCtrlSpaceAsync(rootNode.StoreName, autoSend)
                    : await PasteStoreFullTextForCtrlSpaceAsync(rootNode.StoreName, autoSend);

                if (!success)
                {
                    string action = autoSend ? "Ctrl+Space发送" : "Ctrl+Space粘贴";
                    RecordStoreSendHistory(rootNode.StoreName, action, false, "快捷键执行失败");
                    return;
                }

                if (autoSend)
                {
                    MarkStoreAsSent(rootNode.StoreName, "Ctrl+Space发送", "快捷键发送成功");
                }
                else
                {
                    ClearStoreSentMark(rootNode.StoreName, refreshHeader: true);
                    RecordStoreSendHistory(rootNode.StoreName, "Ctrl+Space粘贴", true, "快捷键粘贴成功");
                }

                _ctrlSpaceSegmentCursor[rootNode.StoreName] = 0;
                
                if (_searchConfig.SkipNextOnCtrlSpace)
                {
                    StatusTextBlock.Text += " (勾选了不跳转，停留在当前项)";
                }
                else
                {
                    // 等待目标窗口（企业微信等）处理完 Ctrl+V 粘贴操作并从剪贴板读取内容后，
                    // 再用下一项的商家名覆盖剪贴板，避免粘贴内容被提前覆盖
                    await Task.Delay(200);
                    await AdvanceToNextMainNodeAndCopyStoreNameAsync(rootNode);
                }
            }
            finally
            {
                // 粘贴流程结束后，恢复 Ctrl 键状态，确保用户可以连贯操作 Ctrl+W/S
                RestoreCtrlKeyIfPhysicallyHeld();
            }
        }

        private async Task HandleCtrlSpaceSegmentedAsync(TreeViewNode rootNode, TreeViewNode selectedNode, bool autoSend)
        {
            int totalSegments = rootNode.Children?.Count ?? 0;
            if (totalSegments <= 0)
            {
                StatusTextBlock.Text = "⚠️ 分段列表为空";
                return;
            }

            bool restartedFromBeginning = ResetCompletedSegmentedStoreProgress(rootNode, refreshHeader: true);
            int selectedIndex = -1;
            if (!ReferenceEquals(selectedNode, rootNode))
            {
                selectedIndex = rootNode.Children.IndexOf(selectedNode);
            }

            bool hasCursor = _ctrlSpaceSegmentCursor.TryGetValue(rootNode.StoreName, out int cursorIndex);
            int currentIndex;
            if (restartedFromBeginning || ReferenceEquals(selectedNode, rootNode))
            {
                // 主项起发时固定从第一段开始
                currentIndex = 0;
            }
            else if (hasCursor && selectedIndex == cursorIndex - 1)
            {
                // 选中项是“刚发完的一段”时，按游标推进下一段
                currentIndex = cursorIndex;
            }
            else if (selectedIndex >= 0)
            {
                // 用户手动选中某个子段时，优先从该段发起
                currentIndex = selectedIndex;
            }
            else if (hasCursor)
            {
                currentIndex = cursorIndex;
            }
            else
            {
                currentIndex = 0;
            }

            if (currentIndex < 0) currentIndex = 0;
            if (currentIndex >= totalSegments) currentIndex = totalSegments - 1;

            TreeViewNode currentSegmentNode = rootNode.Children[currentIndex];
            SelectNodeWithoutCopy(currentSegmentNode, rootNode);

            string payload = currentSegmentNode.RawData;
            if (string.IsNullOrWhiteSpace(payload))
            {
                payload = currentSegmentNode.Text?.Trim() ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(payload) ||
                (payload.StartsWith("(") && payload.EndsWith(")")))
            {
                StatusTextBlock.Text = $"⚠️ 第 {currentIndex + 1}/{totalSegments} 段无可粘贴内容";
                return;
            }

            // ===== 更新悬浮窗：显示当前段号和该段最后一条运单号 =====
            {
                int listIdx = _flatNodeList.IndexOf(rootNode);
                string listSeq = listIdx >= 0 ? $"[{listIdx + 1}] " : string.Empty;

                // 提取当前段最后一条有效运单号（兼容 Tab 分割多列格式）
                string lastTrackingNo = string.Empty;
                if (!string.IsNullOrWhiteSpace(payload))
                {
                    var lines = payload.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    if (lines.Length > 0)
                    {
                        string lastLine = lines[lines.Length - 1];
                        // 多列格式取第一列
                        int tabIdx = lastLine.IndexOf('\t');
                        lastTrackingNo = tabIdx > 0 ? lastLine.Substring(0, tabIdx).Trim() : lastLine.Trim();
                    }
                }

                string segInfo = $"{listSeq}第 {currentIndex + 1}/{totalSegments} 段";
                OsdWindow.ShowMessage(rootNode.StoreName, segInfo, rootNode.GroupName, lastTrackingNo);
            }

            bool success = await PasteTextPayloadForCtrlSpaceAsync(
                rootNode.StoreName,
                payload,
                autoSend,
                $"第 {currentIndex + 1}/{totalSegments} 段");

            if (!success)
            {
                string action = autoSend ? "Ctrl+Space发送" : "Ctrl+Space粘贴";
                RecordStoreSendHistory(rootNode.StoreName, action, false, $"第 {currentIndex + 1}/{totalSegments} 段执行失败");
                return;
            }

            int sentSegments = currentIndex + 1;
            int totalItems = rootNode.Children.Sum(child => CountContentLines(child.RawData));
            if (totalItems <= 0)
            {
                lock (_dataLock)
                {
                    if (_storeData.TryGetValue(rootNode.StoreName, out var rows))
                    {
                        totalItems = rows.Count;
                    }
                }
            }

            int sentItems = 0;
            for (int i = 0; i <= currentIndex && i < rootNode.Children.Count; i++)
            {
                sentItems += CountContentLines(rootNode.Children[i].RawData);
            }
            if (sentItems <= 0)
            {
                int segmentSize = GetSegmentSizeForStore(rootNode.StoreName);
                sentItems = Math.Min(totalItems, sentSegments * segmentSize);
            }

            bool isSegmentCompleted = sentSegments >= totalSegments;
            UpdateSegmentProgressVisual(
                rootNode.StoreName,
                sentSegments: sentSegments,
                totalSegments: totalSegments,
                sentItems: sentItems,
                totalItems: totalItems,
                reason: isSegmentCompleted ? "发送完成" : "发送中(Ctrl+Space)");

            int nextSegmentIndex = currentIndex + 1;
            if (nextSegmentIndex < totalSegments)
            {
                _ctrlSpaceSegmentCursor[rootNode.StoreName] = nextSegmentIndex;
                // 保持当前选中项为“已发送段”，下次 Ctrl+Space 自动推进到下一段
                SelectNodeWithoutCopy(currentSegmentNode, rootNode);
                StatusTextBlock.Text += $"，已完成({currentIndex + 1}/{totalSegments})，下次发送第{nextSegmentIndex + 1}段";
                if (autoSend)
                {
                    RecordStoreSendHistory(rootNode.StoreName, "Ctrl+Space发送", true, $"第 {currentIndex + 1}/{totalSegments} 段发送成功");
                }
                return;
            }

            _ctrlSpaceSegmentCursor[rootNode.StoreName] = 0;
            SelectNodeWithoutCopy(currentSegmentNode, rootNode);
            StatusTextBlock.Text += "，分段已完成";
            if (autoSend)
            {
                MarkStoreAsSent(rootNode.StoreName, "Ctrl+Space发送", "快捷键分段发送完成");
            }
            else
            {
                ClearStoreSentMark(rootNode.StoreName, refreshHeader: true);
                RecordStoreSendHistory(rootNode.StoreName, "Ctrl+Space粘贴", true, "快捷键分段粘贴完成");
            }
            if (_searchConfig.SkipNextOnCtrlSpace)
            {
                StatusTextBlock.Text += " (勾选了不跳转，停留在当前项)";
            }
            else
            {
                await AdvanceToNextMainNodeAndCopyStoreNameAsync(rootNode);
            }
        }

        private bool ResetCompletedSegmentedStoreProgress(TreeViewNode rootNode, bool refreshHeader = true)
        {
            if (rootNode == null ||
                rootNode.Strategy != SendStrategy.TextSegmented ||
                string.IsNullOrWhiteSpace(rootNode.StoreName))
            {
                return false;
            }

            bool isMarkedSent = IsStoreMarkedSent(rootNode.StoreName);
            bool isSegmentCompleted = false;
            if (TryGetSegmentFailureInfo(rootNode.StoreName, out var segmentInfo) && segmentInfo != null)
            {
                int totalSegments = Math.Max(1, rootNode.Children?.Count ?? segmentInfo.TotalSegments);
                isSegmentCompleted = IsSegmentCompleted(segmentInfo, totalSegments);
            }

            if (!isMarkedSent && !isSegmentCompleted)
            {
                return false;
            }

            ClearStoreSentMark(rootNode.StoreName, refreshHeader);
            _ctrlSpaceSegmentCursor[rootNode.StoreName] = 0;
            ClearSegmentFailureInfo(rootNode.StoreName);
            return true;
        }

        private async Task<bool> PasteStoreFullTextForCtrlSpaceAsync(string storeName, bool autoSend)
        {
            if (!TryBuildStoreFullPayload(storeName, out string payload, out int trackingCount))
            {
                return false;
            }

            bool success = await PasteTextPayloadForCtrlSpaceAsync(storeName, payload, autoSend);
            if (success)
            {
                string modeText = autoSend ? "发送" : "粘贴";
                DebugLogManager.Log("CtrlSpace", $"商家={storeName}, 模式=整单{modeText}, 条数={trackingCount}");
            }

            return success;
        }

        private async Task<bool> PasteExcelForCtrlSpaceAsync(string storeName, bool autoSend)
        {
            string filePath;
            lock (_dataLock)
            {
                if (!_exportedFilePaths.TryGetValue(storeName, out filePath))
                {
                    StatusTextBlock.Text = "❌ 未找到文件路径";
                    return false;
                }
            }

            if (!File.Exists(filePath))
            {
                StatusTextBlock.Text = "❌ 文件不存在";
                return false;
            }

            var data = new DataObject();
            data.SetData(DataFormats.FileDrop, new[] { filePath });

            // 快捷键发送走极速参数，减少剪贴板重试等待带来的卡顿
            if (!await SetClipboardWithRetryAsync(data, maxAttempts: 8, retryDelayMs: 8))
            {
                StatusTextBlock.Text = "❌ 文件剪贴板写入失败";
                return false;
            }

            await Task.Delay(8);
            SimulatePaste();

            if (autoSend)
            {
                // 按用户要求：快捷键发送仅使用 Enter 单键，不再 Alt+S 双保险
                const int enterDelayMs = 8;
                await Task.Delay(enterDelayMs);
                SimulateEnter();
                DebugLogManager.Log("CtrlSpace", $"文件发送延迟: {enterDelayMs}ms, 文件={Path.GetFileName(filePath)}");
                StatusTextBlock.Text = $"✅ [Ctrl+Space] 已发送文件: {storeName}";
            }
            else
            {
                StatusTextBlock.Text = $"📋 [Ctrl+Space] 已粘贴文件: {storeName}";
            }

            _currentItemPasted = true;
            _lastPastedStoreName = storeName;
            return true;
        }

        private async Task<bool> PasteTextPayloadForCtrlSpaceAsync(string storeName, string payload, bool autoSend, string? segmentTag = null)
        {
            // 快捷键发送走极速参数，减少剪贴板重试等待带来的卡顿
            if (!await SetClipboardWithRetryAsync(payload, maxAttempts: 8, retryDelayMs: 8))
            {
                StatusTextBlock.Text = "❌ 剪贴板被占用";
                return false;
            }

            await Task.Delay(8);
            SimulatePaste();

            string suffix = string.IsNullOrWhiteSpace(segmentTag) ? string.Empty : $" ({segmentTag})";
            if (autoSend)
            {
                // 按用户要求：快捷键发送仅使用 Enter 单键，不再 Alt+S 双保险
                const int enterDelayMs = 8;
                await Task.Delay(enterDelayMs);
                SimulateEnter();
                DebugLogManager.Log("CtrlSpace", $"文本发送延迟: {enterDelayMs}ms, 长度={payload.Length}");
                StatusTextBlock.Text = $"✅ [Ctrl+Space] 已发送: {storeName}{suffix}";
            }
            else
            {
                StatusTextBlock.Text = $"📋 [Ctrl+Space] 已粘贴: {storeName}{suffix}";
            }

            _currentItemPasted = true;
            _lastPastedStoreName = storeName;
            return true;
        }

        private bool IsForegroundWeWorkWindow()
        {
            IntPtr hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero)
            {
                return false;
            }

            try
            {
                GetWindowThreadProcessId(hwnd, out uint pid);
                if (pid == 0)
                {
                    return false;
                }

                using var process = Process.GetProcessById((int)pid);
                return process.ProcessName.Equals("WXWork", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private bool TryBuildStoreFullPayload(string storeName, out string payload, out int trackingCount)
        {
            payload = string.Empty;
            trackingCount = 0;

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

            trackingCount = trackingNumbers.Count;
            var payloadMode = ResolveStorePayloadMode(storeName, trackingNumbers);
            bool isNormalMode = payloadMode == StorePayloadMode.Normal;

            var sb = new StringBuilder();
            // 统一规则：仅普通(2列)模式在开头追加店铺名；4列/5列均不追加
            if (isNormalMode)
            {
                sb.AppendLine(storeName);
            }

            foreach (var num in trackingNumbers)
            {
                sb.AppendLine(num);
            }

            if (isNormalMode)
            {
                sb.AppendLine(GetCurrentTailMessage());
            }

            payload = sb.ToString();
            return true;
        }

        private bool TryResolveRootNode(TreeViewNode selectedNode, out TreeViewNode rootNode)
        {
            rootNode = null;
            if (selectedNode == null || _treeViewCollection == null)
            {
                return false;
            }

            foreach (var top in _treeViewCollection)
            {
                if (top == null || string.IsNullOrWhiteSpace(top.StoreName))
                {
                    continue;
                }

                if (ReferenceEquals(top, selectedNode))
                {
                    rootNode = top;
                    return true;
                }

                if (top.Children != null && top.Children.Contains(selectedNode))
                {
                    rootNode = top;
                    return true;
                }
            }

            return false;
        }

        private TreeViewNode? GetNextMainNode(TreeViewNode rootNode)
        {
            if (rootNode == null)
            {
                return null;
            }

            if (_flatNodeList.Count == 0)
            {
                RebuildFlatNodeList();
            }

            int currentIndex = _flatNodeList.IndexOf(rootNode);
            if (currentIndex < 0)
            {
                currentIndex = _currentSelectedIndex;
            }

            for (int i = currentIndex + 1; i < _flatNodeList.Count; i++)
            {
                var node = _flatNodeList[i];
                if (IsSelectableNode(node))
                {
                    return node;
                }
            }

            return null;
        }

        private async Task AdvanceToNextMainNodeAndCopyStoreNameAsync(TreeViewNode rootNode)
        {
            var nextNode = GetNextMainNode(rootNode);
            if (nextNode == null)
            {
                StatusTextBlock.Text += "，已到列表末尾";
                return;
            }

            if (!ReferenceEquals(_currentSelectedNode, nextNode))
            {
                SuppressNextSelectionOsd();
            }
            SelectNodeWithoutCopy(nextNode);
            if (!TryGetPreferredSearchCopyText(nextNode, out string copyText, out string copyType))
            {
                StatusTextBlock.Text += "，已跳到下一项(无可复制内容)";
                return;
            }

            bool copied = await SetClipboardWithRetryAsync(copyText, maxAttempts: 8, retryDelayMs: 8);
            StatusTextBlock.Text += copied
                ? $"，已跳到下一项并复制{copyType}"
                : $"，已跳到下一项(复制{copyType}失败)";
        }

        private void SelectNodeWithoutCopy(TreeViewNode targetNode, TreeViewNode? rootToExpand = null)
        {
            if (targetNode == null)
            {
                return;
            }

            Interlocked.Increment(ref _selectionCopyGuard);
            try
            {
                if (rootToExpand != null)
                {
                    ExpandRootNode(rootToExpand);
                }

                if (_currentSelectedNode != null && !ReferenceEquals(_currentSelectedNode, targetNode))
                {
                    _currentSelectedNode.IsSelected = false;
                }

                targetNode.IsSelected = true;
                _currentSelectedNode = targetNode;

                TreeViewNode indexAnchor = rootToExpand ?? targetNode;
                if (_flatNodeList.Contains(indexAnchor))
                {
                    _currentSelectedIndex = _flatNodeList.IndexOf(indexAnchor);
                    try
                    {
                        ScrollToNode(indexAnchor);
                    }
                    catch (InvalidOperationException)
                    {
                        // 容器正在生成中，忽略滚动操作
                        Debug.WriteLine("[SelectNodeWithoutCopy] ScrollToNode 被跳过：容器正在生成中");
                    }
                }

                try
                {
                    StoreTreeView.UpdateLayout();

                    TreeViewItem? container = null;
                    if (rootToExpand != null)
                    {
                        container = GetTreeViewItemForNode(rootToExpand, targetNode);
                    }
                    else if (_flatNodeList.Contains(targetNode))
                    {
                        container = StoreTreeView.ItemContainerGenerator.ContainerFromItem(targetNode) as TreeViewItem;
                    }

                    container?.Focus();
                }
                catch (InvalidOperationException)
                {
                    // 容器正在生成中，忽略焦点操作
                    Debug.WriteLine("[SelectNodeWithoutCopy] UpdateLayout/Focus 被跳过：容器正在生成中");
                }
            }
            finally
            {
                Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    Interlocked.Exchange(ref _selectionCopyGuard, 0);
                }), System.Windows.Threading.DispatcherPriority.ContextIdle);
            }
        }

        private void ExpandRootNode(TreeViewNode rootNode)
        {
            if (rootNode == null)
            {
                return;
            }

            if (StoreTreeView.ItemContainerGenerator.ContainerFromItem(rootNode) is TreeViewItem rootContainer)
            {
                if (!rootContainer.IsExpanded)
                {
                    rootContainer.IsExpanded = true;
                    StoreTreeView.UpdateLayout();
                }
            }
        }

        private TreeViewItem? GetTreeViewItemForNode(TreeViewNode rootNode, TreeViewNode targetNode)
        {
            if (rootNode == null || targetNode == null)
            {
                return null;
            }

            if (StoreTreeView.ItemContainerGenerator.ContainerFromItem(rootNode) is not TreeViewItem rootContainer)
            {
                return null;
            }

            if (!rootContainer.IsExpanded)
            {
                rootContainer.IsExpanded = true;
                StoreTreeView.UpdateLayout();
            }

            if (ReferenceEquals(rootNode, targetNode))
            {
                return rootContainer;
            }

            return rootContainer.ItemContainerGenerator.ContainerFromItem(targetNode) as TreeViewItem;
        }


        private async Task<bool> PasteFullStoreInfoBlindAsync(string storeName)
        {
            // ✅ [Support Custom Message Mode]
            // If in custom mode, the key is "RealStoreName##Message", we need to extract the real name for display/logging if needed.
            // But for looking up data, we use the full key.
            
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

            var payloadMode = ResolveStorePayloadMode(storeName, trackingNumbers);
            bool isNormalMode = payloadMode == StorePayloadMode.Normal;
            string displayStoreName = storeName;

            var sb = new StringBuilder();
            
            // 🔍 [DEBUG] 输出当前模式，帮助调试粘贴内容
            System.Diagnostics.Debug.WriteLine($"[PASTE-DEBUG] payloadMode={payloadMode}, _isCustomMessageMode={_isCustomMessageMode}, _isIssueMode={_isIssueMode}, storeName={storeName}");
            
            // 统一规则：仅普通(2列)模式在开头追加店铺名；4列/5列均不追加
            if (isNormalMode)
            {
                sb.AppendLine(displayStoreName);
            }
            
            foreach (var num in trackingNumbers) sb.AppendLine(num);
            
            if (isNormalMode)
            {
                sb.AppendLine(GetCurrentTailMessage());
            }

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

                MarkStoreAsSent(storeName);
                StatusTextBlock.Text = $"✅ [快捷] 已发送: {displayStoreName}";
            }
            else
            {
                StatusTextBlock.Text = $"📋 [快捷] 已粘贴: {displayStoreName}";
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
            string businessStoreName = NormalizeStoreNameForBusinessInfo(storeName);
            if (string.IsNullOrWhiteSpace(businessStoreName))
            {
                businessStoreName = storeName?.Trim() ?? string.Empty;
            }

            // 1. 更新内存列表 (_businessInfoList)
            var info = _businessInfoList.FirstOrDefault(b => b.StoreName == businessStoreName);
            if (info == null && !string.Equals(storeName, businessStoreName, StringComparison.Ordinal))
            {
                // 兼容旧数据：历史上可能把 "商家名##话术" 写进了 StoreName
                info = _businessInfoList.FirstOrDefault(b => b.StoreName == storeName);
            }
            if (info == null)
            {
                info = new BusinessInfo { StoreName = businessStoreName };
                _businessInfoList.Add(info);
            }
            else
            {
                // 迁移为标准键：仅保留真实商家名
                info.StoreName = businessStoreName;
            }

            _businessInfoList.RemoveAll(b =>
                b != info &&
                (string.Equals(b.StoreName, businessStoreName, StringComparison.Ordinal) ||
                 string.Equals(b.StoreName, storeName, StringComparison.Ordinal)));

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
                    Clipboard.SetDataObject(data, false);
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
                MarkStoreAsSent(storeName);
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

                        await MouseHelper.HumanLikeClickAsync(clickX, clickY, 90);
                        await Task.Delay(110);
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

                        await MouseHelper.HumanLikeClickAsync(contentX, contentY, 95);

                        // 🔥 关键：企微需要更长的焦点稳定时间
                        await Task.Delay(220);
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






        // ✅ [已删除简化版 SearchCurrentItemAsync]
        // 原因：存在两个同名方法导致调用混乱
        // 现在所有调用统一使用 SearchCurrentItemAsync(bool isAutoMode, CancellationToken token) 版本
        // 该版本位于第 922 行，包含完整的 OCR 定位 + 鼠标点击逻辑
        private async Task<bool> PasteAndVerifySendAsync(string contentToSend, bool isFile, CancellationToken token = default)
        {
            Action<string> Log = (msg) => System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [SendVerify] {msg}");

            if (token.IsCancellationRequested)
            {
                return false;
            }

            // ✅ Fix: 如果是发送文本，必须在此处更新剪贴板，否则会重复粘贴上一次的文件
            if (!isFile && !string.IsNullOrEmpty(contentToSend))
            {
                bool copied = await SetClipboardWithRetryAsync(contentToSend);
                if (!copied)
                {
                    Log("❌ [SendVerify] 设置剪贴板失败");
                    return false;
                }

                try { await Task.Delay(100, token); } catch (OperationCanceledException) { return false; }
            }

            IntPtr targetHwnd = GetForegroundWindow();

            if (!RobustActivateWindow(targetHwnd))
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ 窗口无法激活，发送中止");
                return false;
            }


            // 获取窗口坐标
            if (!GetWindowRect(targetHwnd, out RECT rect)) return false;

            bool isWework = false;
            string expectedGroupName = string.Empty;
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                if (_currentSelectedNode != null)
                {
                    isWework = "企业微信".Equals(_currentSelectedNode.Source);
                    expectedGroupName = _currentSelectedNode.GroupName?.Trim() ?? string.Empty;
                }
            });

            // ------------------------------------------------------------
            // 🚀 [安全加固] 在极速模式发送前，强制校验当前窗口标题是否匹配群名
            // ------------------------------------------------------------
            if (!string.IsNullOrWhiteSpace(expectedGroupName))
            {
                // [强制] 每次都校验，移除之前的跳过优化
                Log($"🔍 [SendVerify] 正在校验群名: 预期='{expectedGroupName}'");
                 
                string currentTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(targetHwnd, isWework);
                 
                bool isMatch = _screenshotHelper.IsFuzzyMatch(expectedGroupName, currentTitle);
                if (!isMatch && !string.IsNullOrEmpty(currentTitle) && 
                    (currentTitle.Contains(expectedGroupName, StringComparison.OrdinalIgnoreCase) ||
                     expectedGroupName.Contains(currentTitle, StringComparison.OrdinalIgnoreCase)))
                {
                    isMatch = true;
                }

                if (!isMatch)
                {
                    string msg = $"⚠️ [SendVerify] 校验失败: 预期'{expectedGroupName}' vs 实际'{currentTitle}'，本次发送终止";
                    Log(msg);
                    DebugLogManager.Log("SendVerify", msg);
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = msg);
                    // 关键修复：禁止在发送函数内递归触发 SearchCurrentItemAsync，避免分段发送重入嵌套。
                    return false;
                }
                else
                {
                    Log($"✅ [SendVerify] 窗口校验通过: {currentTitle}");
                }
            }

            int clickX = 0, clickY = 0;
            // ✅ Fix: 使用 YOLO 获取精准点击坐标，而不是硬编码
            var inputRes = await _screenshotHelper.GetInputBoxClickCoordinatesAsync(targetHwnd, isWework);
            if (inputRes.success)
            {
                clickX = inputRes.x;
                clickY = inputRes.y;
                Log($"[SendVerify] YOLO 锁定输入框坐标: ({clickX}, {clickY})");
            }
            else
            {
                // 兜底策略：如果检测失败，回退到原有估算坐标
                clickX = rect.Left + (isWework ? 380 : 310);
                clickY = rect.Bottom - 70;
                Log($"⚠️ [SendVerify] YOLO 失败，使用硬编码坐标: ({clickX}, {clickY})");
            }

            // === 动作 A: 激活输入框并粘贴 ===
            await MouseHelper.HumanLikeClickAsync(clickX, clickY, 90);
            try { await Task.Delay(35, token); } catch (OperationCanceledException) { return false; }

            // 自动发送文本前先清空输入框，避免残留内容叠加
            if (_isAutoRunning && !isFile)
            {
                bool cleared = await ClearInputBoxBeforeAutoSendAsync(token);
                if (!cleared)
                {
                    Log("⚠️ [SendVerify] 清空输入框未完全确认，继续发送流程");
                }
            }

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

            // 3. 【优化】移除发送前的 OCR 粘贴检测
            // 用户反馈该步骤拖慢速度且识别不准，直接跳过，依靠后续的“发送后验证”来确认
            Log("⚡ [极速] 已跳过粘贴内容检查(窗口已验证)，准备发送...");
            try { await Task.Delay(100, token); } catch (OperationCanceledException) { return false; } // 留给 UI 极短渲染时间

            // === 动作 B: 执行发送 (双保险) ===

            // 再次点击输入框
            await MouseHelper.HumanLikeClickAsync(clickX, clickY, 80);
            try { await Task.Delay(35, token); } catch (OperationCanceledException) { return false; }

            bool skipEnter = false;
            // ✅ 新增：如果是企业微信发送文件，尝试处理“转为在线表格”弹窗
            // ✅ 新增：如果是企业微信发送文件，尝试处理“转为在线表格”弹窗
            if (isFile && isWework)
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "🔎 检测企业微信弹窗(OCR)...");
                
                // OCR 查找 "使用原文件"
                // OCR 查找 "使用原文件"
                System.Drawing.Point? clickPoint = null;
                for (int k = 0; k < 5; k++)
                {
                    if (token.IsCancellationRequested) return false;

                    clickPoint = await _screenshotHelper.FindPopupTextPositionAsync(targetHwnd, "使用原文件");
                    if (clickPoint != null) break;
                    try { await Task.Delay(300, token); } catch (OperationCanceledException) { return false; }
                }

                if (clickPoint != null)
                {
                    int cx = (int)clickPoint.Value.X;
                    int cy = (int)clickPoint.Value.Y;
                    Log($"⚡ OCR 定位到 '使用原文件' ({cx}, {cy})，执行点击...");

                    await MouseHelper.HumanLikeClickAsync(cx, cy, 90);

                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "⚡ 已点击'使用原文件' (自动发送)");
                    skipEnter = true; 
                }
                else
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "⚠️ 未检测到弹窗，使用常规发送");
                }
            }

            if (!skipEnter)
            {
                // 方案 1: Alt + S
                SimulateAltS();
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "✉️ 发送指令 (Alt+S)...");
                try { await Task.Delay(300, token); } catch (OperationCanceledException) { return false; }

                // 方案 2: Enter (补刀)
                SimulateEnter();
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text += " + (Enter补刀)...");
            }
            else
            {
                try { await Task.Delay(400, token); } catch (OperationCanceledException) { return false; } // 弹窗点击后缓冲
            }

            // ============================================================
            // 🌟 YOLO 精准验证 (3屏快速截取 + 模糊匹配)
            // ============================================================
            const int MAX_VERIFY_ROUNDS = 3; // 最多验证3轮（每轮3屏 = 最多9次截取）

            for (int round = 0; round < MAX_VERIFY_ROUNDS; round++)
            {
                try { await Task.Delay(200, token); } catch (OperationCanceledException) { return false; } // 等待 UI 渲染
                if (token.IsCancellationRequested) return false;

                // 检查窗口焦点：焦点丢失视为发送失败，要求严格验证
                IntPtr currentForeground = GetForegroundWindow();
                if (currentForeground != targetHwnd)
                {
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ 验证期间窗口失焦，发送失败。");
                    return false;
                }

                // 调用 YOLO 精准验证（内部快速截取3屏）
                var verifyResult = await _screenshotHelper.VerifySendWithYoloAsync(targetHwnd, isWework, keyword, token);
                if (token.IsCancellationRequested || string.Equals(verifyResult.verifyMethod, "已取消", StringComparison.Ordinal))
                {
                    return false;
                }

                if (verifyResult.success)
                {
                    if (!string.IsNullOrWhiteSpace(expectedGroupName))
                    {
                        string finalTitle = await _screenshotHelper.GetWeChatWindowTitleTextAsync(targetHwnd, isWework);
                        bool titleMatched = _screenshotHelper.IsFuzzyMatch(expectedGroupName, finalTitle);
                        if (!titleMatched && !string.IsNullOrWhiteSpace(finalTitle))
                        {
                            titleMatched = finalTitle.Contains(expectedGroupName, StringComparison.OrdinalIgnoreCase)
                                           || expectedGroupName.Contains(finalTitle, StringComparison.OrdinalIgnoreCase);
                        }

                        if (!titleMatched)
                        {
                            string mismatchMsg = $"❌ [SendVerify] 发送后群标题校验失败: 预期='{expectedGroupName}', 实际='{finalTitle}'，标记失败并重试搜索";
                            Log(mismatchMsg);
                            DebugLogManager.Log("SendVerify", mismatchMsg);
                            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = mismatchMsg);
                            return false;
                        }

                        string postOkMsg = $"✅ [SendVerify] 发送后群标题二次校验通过: {finalTitle}";
                        Log(postOkMsg);
                        DebugLogManager.Log("SendVerify", postOkMsg);
                    }

                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"✅ [{verifyResult.verifyMethod}] 发送成功。");
                    return true;
                }

                // 补刀重试：发送未成功，再次点击输入框 + Enter
                if (round < MAX_VERIFY_ROUNDS - 1)
                {
                    Log($"⚠️ 第{round + 1}轮验证未通过，补刀重试...");
                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"⚠️ 第{round + 1}轮未通过，补刀Enter重试...");
                    await MouseHelper.HumanLikeClickAsync(clickX, clickY, 80);
                    try { await Task.Delay(35, token); } catch (OperationCanceledException) { return false; }
                    SimulateEnter();
                }
            }

            // 所有轮次验证均未通过 → 发送失败
            Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"❌ [YOLO验证] 发送失败: 未检测到发送内容。");
            return false;
        }
        private async Task<bool> ClearInputBoxBeforeAutoSendAsync(CancellationToken token = default)
        {
            const int clearDelayMs = 8;

            try
            {
                _inputBackend.KeyChord(InputKey.LeftControl, InputKey.A);
                await Task.Delay(clearDelayMs, token);
                _inputBackend.KeyPress(InputKey.Backspace);
                await Task.Delay(clearDelayMs, token);

                // 二次清空：兼容输入法/富文本偶发未全清的情况
                _inputBackend.KeyChord(InputKey.LeftControl, InputKey.A);
                await Task.Delay(clearDelayMs, token);
                _inputBackend.KeyPress(InputKey.Delete);
                await Task.Delay(clearDelayMs, token);
                return true;
            }
            catch (OperationCanceledException)
            {
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SendVerify] 清空输入框异常: {ex.Message}");
                return false;
            }
        }

        private void SimulateEnter()
        {
            try
            {
                // 模拟按下 Enter 键
                _inputBackend.KeyPress(InputKey.Enter);
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

        private bool IsSelectableNode(TreeViewNode? node)
        {
            return node != null && !string.IsNullOrEmpty(node.StoreName) && node.StoreName != "FAIL_SEPARATOR";
        }

        private int FindBestSelectionIndex(int preferredIndex)
        {
            if (_flatNodeList.Count == 0)
            {
                return -1;
            }

            int start = preferredIndex;
            if (start < 0) start = 0;
            if (start >= _flatNodeList.Count) start = _flatNodeList.Count - 1;

            if (IsSelectableNode(_flatNodeList[start]))
            {
                return start;
            }

            for (int i = start + 1; i < _flatNodeList.Count; i++)
            {
                if (IsSelectableNode(_flatNodeList[i]))
                {
                    return i;
                }
            }

            for (int i = start - 1; i >= 0; i--)
            {
                if (IsSelectableNode(_flatNodeList[i]))
                {
                    return i;
                }
            }

            // 兜底：全是分隔符时至少选中一个可见项，而不是置空
            return start;
        }

        private bool SelectBestNode(int preferredIndex)
        {
            int index = FindBestSelectionIndex(preferredIndex);
            if (index < 0 || index >= _flatNodeList.Count)
            {
                return false;
            }

            _currentSelectedIndex = index;
            var node = _flatNodeList[index];
            FocusAndSelectItem(node);
            _currentSelectedNode = node;
            return true;
        }

        /// <summary>
        /// 自动化专用：前进到下一项；如果已经到达末尾则返回 false，并保留一个合适的选中项。
        /// </summary>
        private bool TryNavigateToNextNodeForAuto()
        {
            if (_flatNodeList.Count == 0)
            {
                RebuildFlatNodeList();
            }

            if (_flatNodeList.Count == 0)
            {
                _currentSelectedIndex = -1;
                _currentSelectedNode = null;
                return false;
            }

            // 兜底：索引异常时尝试由当前节点反查索引
            if ((_currentSelectedIndex < 0 || _currentSelectedIndex >= _flatNodeList.Count) && _currentSelectedNode != null)
            {
                _currentSelectedIndex = _flatNodeList.IndexOf(_currentSelectedNode);
            }

            int nextIndex = _currentSelectedIndex + 1;
            if (_currentSelectedIndex < 0)
            {
                nextIndex = 0;
            }

            if (nextIndex < 0 || nextIndex >= _flatNodeList.Count)
            {
                SelectBestNode(_flatNodeList.Count - 1);
                return false;
            }

            for (int i = nextIndex; i < _flatNodeList.Count; i++)
            {
                var node = _flatNodeList[i];
                if (!IsSelectableNode(node))
                {
                    continue;
                }

                _currentSelectedIndex = i;
                FocusAndSelectItem(node);
                _currentSelectedNode = node;
                return true;
            }

            SelectBestNode(_flatNodeList.Count - 1);
            return false;
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
                    searchText = NormalizeStoreNameForSearch(_currentSelectedNode.StoreName);
                    if (string.IsNullOrWhiteSpace(searchText))
                    {
                        searchText = _currentSelectedNode.StoreName?.Trim() ?? string.Empty;
                    }
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
            _childParentMap.Clear();
            // ✔️ 修复: 改用 IEnumerable 兼容 ObservableCollection
            if (StoreTreeView.ItemsSource is IEnumerable<TreeViewNode> nodes)
            {
                foreach (var node in nodes)
                {
                    if (!string.IsNullOrEmpty(node.StoreName))
                    {
                        _flatNodeList.Add(node);
                        // 同步建立子节点 -> 父节点映射
                        if (node.Children != null)
                        {
                            foreach (var child in node.Children)
                            {
                                _childParentMap[child] = node;
                            }
                        }
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

            if (_isAutoRunning ||
                Volatile.Read(ref _clipboardSearchGuard) > 0 ||
                Volatile.Read(ref _selectionCopyGuard) > 0)
            {
                return;
            }

            _currentSelectedNode = node;
            ResetSearchState();

            if (Interlocked.CompareExchange(ref _copyingFlag, 1, 0) == 1) return;
            // 主列表复制内容跟随搜索模式：群名模式优先群名，商家模式用商家名
            CopyPreferredSearchText(node);
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
                    if (!TryGetPreferredSearchCopyText(_currentSelectedNode, out string copyText, out string copyType))
                    {
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "无法获取可复制内容");
                        return;
                    }

                    if (await SetClipboardWithRetryAsync(copyText))
                    {

                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            SimulatePaste();
                            StatusTextBlock.Text = $"已粘贴{copyType}: '{copyText}'";
                        });
                    }
                    else
                    {
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "无法复制到剪贴板");
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
            bool isWework = "企业微信".Equals(_currentSelectedNode?.Source);
            switch (_currentSelectedNode.Strategy)
            {
                case SendStrategy.FileExcel:
                    PasteExcelFile(storeName);
                    break;
                case SendStrategy.TextSegmented:
                    _ = PasteStoreInfoInSegmentsAsync(storeName, isWework);
                    break;
                default:
                    PasteFullStoreInfo(storeName);
                    break;
            }
        }




        // MainWindow.xaml.cs

        private async Task<bool> PasteExcelFileAsync(string storeName, bool isWework, CancellationToken token = default)
        {
            Action<string> Log = (msg) => System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [粘贴文件] {msg}");
            Log($"开始处理文件: {storeName}, isWework: {isWework}");

            if (string.IsNullOrEmpty(storeName)) return false;
            if (token.IsCancellationRequested) return false;

            IntPtr targetHwnd = GetForegroundWindow();
            Log($"当前窗口句柄: {targetHwnd}");

            string filePath;
            lock (_dataLock) { if (!_exportedFilePaths.TryGetValue(storeName, out filePath)) return false; }
            if (!File.Exists(filePath))
            {
                Log($"❌ 文件不存在: {filePath}");
                return false;
            }

            StorePayloadMode payloadMode = StorePayloadMode.Normal;
            lock (_dataLock)
            {
                if (_storeData.TryGetValue(storeName, out var rowsForMode))
                {
                    payloadMode = ResolveStorePayloadMode(storeName, rowsForMode);
                }
            }
            bool shouldAppendFixedMessage = payloadMode == StorePayloadMode.Normal;
            string displayStoreName = storeName;
            Log($"[MODE] payloadMode={payloadMode}, shouldAppendFixedMessage={shouldAppendFixedMessage}");

            // 剪贴板
            Log("设置文件到剪贴板...");

            // ✅ 准备数据对象
            var data = new DataObject();
            data.SetData(DataFormats.FileDrop, new string[] { filePath });

            // ✅ 使用重试机制
            if (!await SetClipboardWithRetryAsync(data))
            {
                Log("❌ 剪贴板设置失败");
                return false;
            }
            Log("剪贴板就绪");

            try { await Task.Delay(50, token); } catch (OperationCanceledException) { return false; }

            var innerTask = await Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                if (token.IsCancellationRequested) return false;

                if (targetHwnd != GetForegroundWindow())
                {
                    Log("⚠️ 窗口失焦，抢回焦点...");
                    RobustActivateWindow(targetHwnd);
                    try { await Task.Delay(50, token); } catch (OperationCanceledException) { return false; }
                }

                var inputRes = await _screenshotHelper.GetInputBoxClickCoordinatesAsync(targetHwnd, isWework);
                if (inputRes.success)
                {
                    int clickX = inputRes.x;
                    int clickY = inputRes.y;
                    Log($"点击坐标: ({clickX}, {clickY})");
                    await MouseHelper.HumanLikeClickAsync(clickX, clickY, 85);
                    try { await Task.Delay(35, token); } catch (OperationCanceledException) { return false; }
                }
                else
                {
                    Log("⚠️ 坐标检测失败");
                }

                bool result = false;
                // ✅ Fix: 如果处于 F1 自动化模式 (_isAutoRunning)，强制启用自动发送逻辑
                if (AutoSendCheckBox.IsChecked == true || _isAutoRunning)
                {
                    string sendAction = _isAutoRunning ? "自动化发送" : "自动发送";
                    Log("执行自动发送...");
                    result = await PasteAndVerifySendAsync(filePath, true, token);
                    if (result)
                    {
                        _currentItemPasted = true;
                        _lastPastedStoreName = storeName;
                        MarkStoreAsSent(storeName, sendAction, "文件发送成功");
                        StatusTextBlock.Text = $"✅ [自动] 已发送文件: {displayStoreName}";
                        Log("✅ 发送成功");

                        // 仅普通未发货模式追加固定话术
                        if (shouldAppendFixedMessage)
                        {
                            try { await Task.Delay(200, token); } catch (OperationCanceledException) { return false; }
                            Log("➕ 追加发送未发货预警...");
                            StatusTextBlock.Text += " + 正在追加预警...";
                            await PasteAndVerifySendAsync(GetCurrentTailMessage(), false, token);
                        }
                        // ✅ [重构] 自定义话术模式不再走文件发送路径，已废弃此分支
                    }
                    else
                    {
                        Log("❌ 发送失败");
                        RecordStoreSendHistory(storeName, sendAction, false, "文件发送失败");
                    }
                }
                else
                {
                    Log("执行手动粘贴...");
                    SimulatePaste();
                    Log("Ctrl+V 已发送");

                    if (isWework)
                    {
                         // ✅ 修复：手动模式下不自动点击"使用原文件"，让用户自己操作
                         StatusTextBlock.Text = $"📋 [手动] 已粘贴文件至企业微信: {displayStoreName}";
                    }
                    else
                    {
                        // ✅ 修复：手动模式仅粘贴，不自动发送话术
                        StatusTextBlock.Text = $"📋 [手动] 已粘贴文件: {displayStoreName}";
                    }
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

        private async Task<bool> PasteFullStoreInfoAsync(string storeName, bool isWework, CancellationToken token = default)
        {
            Action<string> Log = (msg) => System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [粘贴文本] {msg}");
            Log($"开始处理商家: {storeName}, isWework: {isWework}");

            if (string.IsNullOrEmpty(storeName)) return false;
            if (token.IsCancellationRequested) return false;

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

            var payloadMode = ResolveStorePayloadMode(storeName, trackingNumbers);
            bool isNormalMode = payloadMode == StorePayloadMode.Normal;
            string displayStoreName = storeName;

            var sb = new StringBuilder();
            
            // 🔍 [DEBUG] 输出当前模式，帮助调试粘贴内容
            Log($"[MODE] payloadMode={payloadMode}, _isCustomMessageMode={_isCustomMessageMode}, _isIssueMode={_isIssueMode}");
            
            // 统一规则：仅普通(2列)模式在开头追加店铺名；4列/5列均不追加
            if (isNormalMode)
            {
                sb.AppendLine(displayStoreName);
            }
            
            foreach (var num in trackingNumbers) sb.AppendLine(num);
            
            if (isNormalMode)
            {
                sb.AppendLine(GetCurrentTailMessage());
            }
            string fullText = sb.ToString();

            // 3. 剪贴板
            Log("正在设置剪贴板...");
            if (!await SetClipboardWithRetryAsync(fullText))
            {
                Log("❌ 剪贴板设置失败");
                return false;
            }
            Log("剪贴板设置成功");

            try { await Task.Delay(50, token); } catch (OperationCanceledException) { return false; }

            // 4. UI线程操作
            var innerTask = await Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                if (token.IsCancellationRequested) return false;

                // 焦点复查
                if (targetHwnd != GetForegroundWindow())
                {
                    Log($"⚠️ 窗口失焦 (当前: {GetForegroundWindow()})，尝试抢回...");
                    RobustActivateWindow(targetHwnd);
                    try { await Task.Delay(50, token); } catch (OperationCanceledException) { return false; }
                }

                // 5. 坐标计算与点击
                var inputRes = await _screenshotHelper.GetInputBoxClickCoordinatesAsync(targetHwnd, isWework);
                Log($"获取点击坐标: {inputRes.success}, X={inputRes.x}, Y={inputRes.y}");

                if (inputRes.success)
                {
                    int clickX = inputRes.x;
                    int clickY = inputRes.y;
                    Log($"执行鼠标点击: ({clickX}, {clickY})");
                    await MouseHelper.HumanLikeClickAsync(clickX, clickY, 85);
                    try { await Task.Delay(35, token); } catch (OperationCanceledException) { return false; }
                }
                else
                {
                    Log("⚠️ 无法计算坐标 (可能窗口最小化或句柄无效)");
                }

                // 6. 粘贴与发送
                bool result = false;
                // ✅ Fix: 如果处于 F1 自动化模式 (_isAutoRunning)，强制启用自动发送逻辑
                if (AutoSendCheckBox.IsChecked == true || _isAutoRunning)
                {
                    string sendAction = _isAutoRunning ? "自动化发送" : "自动发送";
                    Log("模式: 自动发送");
                    // 注意：PasteAndVerifySendAsync 内部日志未在此处展示，需确保该函数也正常
                    result = await PasteAndVerifySendAsync(fullText, false, token);
                    if (result)
                    {
                        _currentItemPasted = true;
                        _lastPastedStoreName = storeName;
                        MarkStoreAsSent(storeName, sendAction, "文本发送成功");
                        StatusTextBlock.Text = $"✅ [自动] 已发送: {displayStoreName}";
                        Log("✅ 发送流程完成 (PasteAndVerifySendAsync 返回 true)");
                    }
                    else
                    {
                        Log("❌ 发送流程失败 (PasteAndVerifySendAsync 返回 false)");
                        RecordStoreSendHistory(storeName, sendAction, false, "文本发送失败");
                    }
                }
                else
                {
                    Log("模式: 手动发送 (仅粘贴)");
                    SimulatePaste();
                    Log("已模拟 Ctrl+V");

                    StatusTextBlock.Text = $"📋 [自动] 已粘贴: {displayStoreName} (等待发送)";
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

        /// <summary>
        /// ✅ [重构] 分段发送商家信息（支持中断和断点续传）
        /// 将数据按 SegmentSize 分段，逐段粘贴发送。如果失败或中断，保留未发送的数据。
        /// </summary>
        private async Task<bool> PasteStoreInfoInSegmentsAsync(string storeName, bool isWework, CancellationToken token = default, bool isAutoMode = false)
        {
            Action<string> Log = (msg) =>
            {
                System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [分段发送] {msg}");
                DebugLogManager.Log("分段发送", msg);
            };
            Log($"开始分段发送: {storeName}, isWework: {isWework}");

            if (string.IsNullOrEmpty(storeName)) return false;

            bool autoSendChecked = false;
            if (Application.Current?.Dispatcher != null)
            {
                autoSendChecked = await Application.Current.Dispatcher.InvokeAsync(() => AutoSendCheckBox.IsChecked == true);
            }

            bool shouldAutoSend = isAutoMode || autoSendChecked;
            Log($"分段执行模式: {(shouldAutoSend ? "发送" : "仅粘贴")} (isAutoMode={isAutoMode}, autoSendChecked={autoSendChecked})");

            // 1. 获取数据 (复制)
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

            StorePayloadMode payloadMode = ResolveStorePayloadMode(storeName, trackingNumbers);
            int segmentSize = GetSegmentSizeForPayloadMode(payloadMode);
            int totalSegments = (int)Math.Ceiling((double)trackingNumbers.Count / segmentSize);
            Log($"共 {trackingNumbers.Count} 条，分 {totalSegments} 段，每段 {segmentSize} 条");

            if (totalSegments <= 0)
            {
                ClearSegmentFailureInfo(storeName);
                return true;
            }

            int startSegmentIndex = 0;
            int sentCount = 0;
            if (TryGetSegmentFailureInfo(storeName, out var resumeInfo) && resumeInfo != null)
            {
                int sentSegmentsFromState = Math.Max(0, Math.Min(resumeInfo.SentSegments, totalSegments));
                int sentItemsFromState = Math.Max(0, Math.Min(resumeInfo.SentItems, trackingNumbers.Count));
                int resumeFromByFailedSegment = Math.Max(0, Math.Min(totalSegments, resumeInfo.FailedSegment - 1));
                int resumeFromBySentSegment = Math.Max(0, Math.Min(totalSegments, sentSegmentsFromState));
                int resumeFrom = Math.Max(resumeFromByFailedSegment, resumeFromBySentSegment);
                bool hasResumeState =
                    resumeInfo.FailedSegment > 1 ||
                    sentSegmentsFromState > 0 ||
                    sentItemsFromState > 0;

                if (hasResumeState)
                {
                    bool isCompletedState = IsSegmentCompleted(resumeInfo, totalSegments) || resumeFrom >= totalSegments;
                    if (isCompletedState)
                    {
                        sentCount = trackingNumbers.Count;
                        UpdateSegmentProgressVisual(
                            storeName,
                            sentSegments: totalSegments,
                            totalSegments: totalSegments,
                            sentItems: sentCount,
                            totalItems: trackingNumbers.Count,
                            reason: "发送完成");
                        Application.Current.Dispatcher.Invoke(() =>
                            StatusTextBlock.Text = $"⏭️ 商家 '{storeName}' 分段已全部发送，自动跳过。");
                        Log("⏭️ 检测到分段已全部完成，跳过该商家。");
                        return true;
                    }

                    startSegmentIndex = resumeFrom;
                    int bySegmentCount = Math.Max(0, Math.Min(trackingNumbers.Count, startSegmentIndex * segmentSize));
                    sentCount = Math.Max(bySegmentCount, sentItemsFromState);

                    Log($"♻️ 检测到断点进度：从第 {startSegmentIndex + 1}/{totalSegments} 段继续，已发送 {sentCount}/{trackingNumbers.Count} 条");
                    UpdateSegmentProgressVisual(
                        storeName,
                        sentSegments: startSegmentIndex,
                        totalSegments: totalSegments,
                        sentItems: sentCount,
                        totalItems: trackingNumbers.Count,
                        reason: "发送中(断点续发)");
                }
            }

            try
            {
                for (int i = startSegmentIndex; i < totalSegments; i++)
                {
                    int segNum = i + 1;

                    // ✅ 支持中断 (F2 按 ESC 或 token)
                    bool isEscPressed = false;
                    // 仅在手动模式下(token为None)检查键盘
                    // ✅ [FIX] 使用全局API检测 ESC，解决窗口失焦无法停止的问题
                    if (token == CancellationToken.None)
                    {
                         isEscPressed = MouseHelper.IsEscPressed();
                    }
                    
                    if (token.IsCancellationRequested || isEscPressed)
                    {
                        Log("🛑 用户中止发送");
                        Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "🛑 发送已中止，剩余片段已保留");
                        await HandlePartialFailure(
                            storeName,
                            sentCount,
                            sentSegments: segNum - 1,
                            failedSegment: segNum,
                            totalSegments: totalSegments,
                            totalItems: trackingNumbers.Count,
                            reason: "用户停止(F2/ESC)",
                            moveToRetryArea: !isAutoMode,
                            historyAction: isAutoMode ? "自动化发送" : "分段发送");
                        return false;
                    }

                    int startIdx = i * segmentSize;
                    var segmentItems = trackingNumbers.Skip(startIdx).Take(segmentSize).ToList();
                    string segmentContent = string.Join("\n", segmentItems);

                    Application.Current.Dispatcher.Invoke(() =>
                        StatusTextBlock.Text = shouldAutoSend
                            ? $"📤 正在发送第 {segNum}/{totalSegments} 段 ({segmentItems.Count} 条)..."
                            : $"📋 正在粘贴第 {segNum}/{totalSegments} 段 ({segmentItems.Count} 条)...");

                    bool success = shouldAutoSend
                        ? await PasteAndVerifySendAsync(segmentContent, false, token)
                        : await PasteSegmentWithoutAutoSendAsync(segmentContent, isWework, token);
                    
                    if (!success)
                    {
                        string actionText = shouldAutoSend ? "发送" : "粘贴";
                        Log($"❌ 第 {segNum} 段{actionText}失败");
                        Application.Current.Dispatcher.Invoke(() =>
                            StatusTextBlock.Text = shouldAutoSend
                                ? $"❌ 第 {segNum} 段发送失败，剩余已保留到列表末尾"
                                : $"❌ 第 {segNum} 段粘贴失败，剩余已保留到列表末尾");
                        await HandlePartialFailure(
                            storeName,
                            sentCount,
                            sentSegments: segNum - 1,
                            failedSegment: segNum,
                            totalSegments: totalSegments,
                            totalItems: trackingNumbers.Count,
                            reason: shouldAutoSend ? "发送失败" : "粘贴失败",
                            moveToRetryArea: !isAutoMode,
                            historyAction: isAutoMode ? "自动化发送" : (shouldAutoSend ? "分段发送" : "分段粘贴"));
                        return false;
                    }

                    sentCount += segmentItems.Count;
                    UpdateSegmentProgressVisual(
                        storeName,
                        sentSegments: segNum,
                        totalSegments: totalSegments,
                        sentItems: sentCount,
                        totalItems: trackingNumbers.Count,
                        reason: shouldAutoSend ? "发送中" : "发送中(待手动发送)");

                    if (!shouldAutoSend)
                    {
                        await Application.Current.Dispatcher.InvokeAsync(() => SaveFileState());

                        bool isLastSegment = segNum >= totalSegments;
                        if (isLastSegment)
                        {
                            UpdateSegmentProgressVisual(
                                storeName,
                                sentSegments: totalSegments,
                                totalSegments: totalSegments,
                                sentItems: trackingNumbers.Count,
                                totalItems: trackingNumbers.Count,
                                reason: "发送完成");
                            _currentItemPasted = true;
                            _lastPastedStoreName = storeName;
                            ClearStoreSentMark(storeName, refreshHeader: true);
                            Application.Current.Dispatcher.Invoke(() =>
                                StatusTextBlock.Text = $"📋 分段已全部粘贴: {storeName}（未自动发送）");
                        }
                        else
                        {
                            _currentItemPasted = false;
                            _lastPastedStoreName = null;
                            Application.Current.Dispatcher.Invoke(() =>
                                StatusTextBlock.Text = $"📋 已粘贴第 {segNum}/{totalSegments} 段（未自动发送），请手动发送后再按 Ctrl+Enter 继续。");
                        }

                        return true;
                    }
                    
                    // Delay
                    if (i < totalSegments - 1)
                    {
                        try
                        {
                            await Task.Delay(_searchConfig.SegmentDelayMs, token);
                        }
                        catch (OperationCanceledException)
                        {
                            int nextSeg = Math.Min(totalSegments, segNum + 1);
                            await HandlePartialFailure(
                                storeName,
                                sentCount,
                                sentSegments: segNum,
                                failedSegment: nextSeg,
                                totalSegments: totalSegments,
                                totalItems: trackingNumbers.Count,
                                reason: "用户停止(F2/ESC)",
                                moveToRetryArea: !isAutoMode,
                                historyAction: isAutoMode ? "自动化发送" : "分段发送");
                            return false;
                        }
                    }
                }

                Log("✅ 分段发送完成");
                UpdateSegmentProgressVisual(
                    storeName,
                    sentSegments: totalSegments,
                    totalSegments: totalSegments,
                    sentItems: trackingNumbers.Count,
                    totalItems: trackingNumbers.Count,
                    reason: "发送完成");
                await Application.Current.Dispatcher.InvokeAsync(() => SaveFileState());
                _currentItemPasted = true;
                _lastPastedStoreName = storeName;
                MarkStoreAsSent(storeName, isAutoMode ? "自动化发送" : "分段发送", "分段发送完成");
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"✅ 分段发送完成: {storeName}");
                return true;
            }
            catch (OperationCanceledException)
            {
                int sentSegments = Math.Min(totalSegments, (int)Math.Ceiling((double)sentCount / segmentSize));
                int failedSegment = Math.Min(totalSegments, sentSegments + 1);
                await HandlePartialFailure(
                    storeName,
                    sentCount,
                    sentSegments: sentSegments,
                    failedSegment: failedSegment,
                    totalSegments: totalSegments,
                    totalItems: trackingNumbers.Count,
                    reason: "用户停止(F2/ESC)",
                    moveToRetryArea: !isAutoMode,
                    historyAction: isAutoMode ? "自动化发送" : "分段发送");
                return false;
            }
            catch (Exception ex)
            {
                Log($"💥 异常: {ex.Message}");
                int sentSegments = Math.Min(totalSegments, (int)Math.Ceiling((double)sentCount / segmentSize));
                int failedSegment = Math.Min(totalSegments, sentSegments + 1);
                await HandlePartialFailure(
                    storeName,
                    sentCount,
                    sentSegments: sentSegments,
                    failedSegment: failedSegment,
                    totalSegments: totalSegments,
                    totalItems: trackingNumbers.Count,
                    reason: "异常中断",
                    moveToRetryArea: !isAutoMode,
                    historyAction: isAutoMode ? "自动化发送" : "分段发送");
                return false;
            }
        }

        private async Task<bool> PasteSegmentWithoutAutoSendAsync(string segmentContent, bool isWework, CancellationToken token)
        {
            if (token.IsCancellationRequested)
            {
                return false;
            }

            if (!await SetClipboardWithRetryAsync(segmentContent))
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ 剪贴板被占用");
                return false;
            }

            try { await Task.Delay(50, token); } catch (OperationCanceledException) { return false; }

            IntPtr targetHwnd = GetForegroundWindow();
            if (targetHwnd == IntPtr.Zero)
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ 未找到目标聊天窗口");
                return false;
            }

            if (!RobustActivateWindow(targetHwnd))
            {
                Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = "❌ 窗口无法激活，粘贴中止");
                return false;
            }

            var inputRes = await _screenshotHelper.GetInputBoxClickCoordinatesAsync(targetHwnd, isWework);
            if (inputRes.success)
            {
                await MouseHelper.HumanLikeClickAsync(inputRes.x, inputRes.y, 85);
                try { await Task.Delay(35, token); } catch (OperationCanceledException) { return false; }
            }

            SimulatePaste();
            return true;
        }

        // 辅助：处理部分失败，记录失败进度，并按需要移动到失败列表
        private async Task HandlePartialFailure(
            string storeName,
            int sentCount,
            int sentSegments,
            int failedSegment,
            int totalSegments,
            int totalItems,
            string reason,
            bool moveToRetryArea,
            string historyAction = "发送")
        {
            if (string.IsNullOrWhiteSpace(storeName))
            {
                return;
            }

            int safeTotalSegments = Math.Max(1, totalSegments);
            int safeSentSegments = Math.Max(0, Math.Min(sentSegments, safeTotalSegments));
            int safeFailedSegment = Math.Max(1, Math.Min(failedSegment, safeTotalSegments));
            int safeSentItems = Math.Max(0, Math.Min(sentCount, totalItems));
            int safeTotalItems = Math.Max(0, totalItems);
            string safeReason = string.IsNullOrWhiteSpace(reason) ? "发送失败" : reason.Trim();

            DebugLogManager.Log(
                "分段异常",
                $"商家={storeName}, 失败段={safeFailedSegment}/{safeTotalSegments}, 已发送段={safeSentSegments}, 已发送条数={safeSentItems}/{safeTotalItems}, 原因={safeReason}");
            RecordStoreSendHistory(storeName, historyAction, false, $"分段异常: {safeReason}");

            lock (_segmentFailureLock)
            {
                _segmentFailureInfos[storeName] = new SegmentFailureInfo
                {
                    FailedSegment = safeFailedSegment,
                    TotalSegments = safeTotalSegments,
                    SentSegments = safeSentSegments,
                    SentItems = safeSentItems,
                    TotalItems = safeTotalItems,
                    Reason = safeReason
                };
            }

            if (moveToRetryArea)
            {
                int preferredIndex = _currentSelectedIndex;

                if (!_failedStores.Contains(storeName))
                {
                    _failedStores.Add(storeName);
                }

                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    SaveFileState();
                    ProcessAndDisplayData();
                    SelectBestNode(preferredIndex);
                });
            }
            else
            {
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    if (_currentSelectedNode != null &&
                        string.Equals(_currentSelectedNode.StoreName, storeName, StringComparison.Ordinal) &&
                        _currentSelectedNode.Strategy == SendStrategy.TextSegmented)
                    {
                        ApplySegmentFailureProgressToNode(_currentSelectedNode);
                    }
                });
            }
        }

        private void ClearSegmentFailureInfo(string storeName)
        {
            if (string.IsNullOrWhiteSpace(storeName))
            {
                return;
            }

            lock (_segmentFailureLock)
            {
                _segmentFailureInfos.Remove(storeName);
            }
        }

        private bool TryGetSegmentFailureInfo(string storeName, out SegmentFailureInfo? info)
        {
            lock (_segmentFailureLock)
            {
                return _segmentFailureInfos.TryGetValue(storeName, out info);
            }
        }

        private void EmitRecentStoreHistoryToDebugLog(int take = 30)
        {
            try
            {
                var recent = StoreSendHistoryRepository.GetRecent(take);
                if (recent.Count == 0)
                {
                    DebugLogManager.Log("历史", "未检测到持久化发送历史。");
                    return;
                }

                DebugLogManager.Log("历史", $"已加载最近 {recent.Count} 条发送历史。");
                foreach (var entry in recent)
                {
                    if (entry == null || string.IsNullOrWhiteSpace(entry.StoreName))
                    {
                        continue;
                    }

                    string resultText = entry.IsSuccess ? "成功" : "失败";
                    string detailText = string.IsNullOrWhiteSpace(entry.Detail) ? string.Empty : $" | 说明={entry.Detail}";
                    string fileName = string.IsNullOrWhiteSpace(entry.FilePath) ? string.Empty : $" | 文件={Path.GetFileName(entry.FilePath)}";
                    string logSource = ResolveHistoryLogSource(entry.Action);
                    DebugLogManager.Log(
                        logSource,
                        $"[恢复] {entry.Timestamp:MM-dd HH:mm:ss} | 店铺={entry.StoreName} | 操作={entry.Action} | 结果={resultText}{detailText}{fileName}");
                }
            }
            catch (Exception ex)
            {
                DebugLogManager.Log("历史", $"加载历史失败: {ex.Message}");
            }
        }

        private void RecordStoreSelectionHistory(TreeViewNode selectedNode)
        {
            if (selectedNode == null || string.IsNullOrWhiteSpace(selectedNode.StoreName) || selectedNode.StoreName == "FAIL_SEPARATOR")
            {
                return;
            }

            TreeViewNode rootNode = selectedNode;
            if (TryResolveRootNode(selectedNode, out TreeViewNode resolvedRoot) && resolvedRoot != null)
            {
                rootNode = resolvedRoot;
            }

            string storeName = rootNode.StoreName?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(storeName) || storeName == "FAIL_SEPARATOR")
            {
                return;
            }

            DateTime now = DateTime.Now;
            if (string.Equals(_lastSelectionHistoryStoreName, storeName, StringComparison.Ordinal) &&
                (now - _lastSelectionHistoryTime).TotalSeconds < 2)
            {
                return;
            }

            _lastSelectionHistoryStoreName = storeName;
            _lastSelectionHistoryTime = now;

            string detail = _isAutoRunning ? "自动流程选中" : "手动选中";
            RecordStoreSendHistory(storeName, "点击店铺", true, detail);
        }

        private void RecordStoreSendHistory(string storeName, string action, bool success, string detail = "")
        {
            if (string.IsNullOrWhiteSpace(storeName) || storeName == "FAIL_SEPARATOR")
            {
                return;
            }

            string normalizedStoreName = NormalizeStoreNameForSearch(storeName);
            if (string.IsNullOrWhiteSpace(normalizedStoreName))
            {
                normalizedStoreName = storeName.Trim();
            }

            string filePath = _lastLoadedFilePath;
            if (string.IsNullOrWhiteSpace(filePath))
            {
                filePath = _searchConfig?.LastOpenedFilePath?.Trim() ?? string.Empty;
            }

            var historyEntry = new StoreSendHistoryEntry
            {
                Timestamp = DateTime.Now,
                StoreName = normalizedStoreName,
                Action = string.IsNullOrWhiteSpace(action) ? "Unknown" : action.Trim(),
                IsSuccess = success,
                Detail = detail ?? string.Empty,
                FilePath = filePath
            };

            StoreSendHistoryRepository.Append(historyEntry);

            string resultText = success ? "成功" : "失败";
            string detailText = string.IsNullOrWhiteSpace(historyEntry.Detail) ? string.Empty : $" | 说明={historyEntry.Detail}";
            string fileName = string.IsNullOrWhiteSpace(historyEntry.FilePath) ? string.Empty : $" | 文件={Path.GetFileName(historyEntry.FilePath)}";
            string logSource = ResolveHistoryLogSource(historyEntry.Action);
            DebugLogManager.Log(
                logSource,
                $"店铺={historyEntry.StoreName} | 操作={historyEntry.Action} | 结果={resultText}{detailText}{fileName}");
        }

        private static string ResolveHistoryLogSource(string action)
        {
            if (string.IsNullOrWhiteSpace(action))
            {
                return "历史";
            }

            string normalizedAction = action.Trim();
            if (normalizedAction.Contains("点击", StringComparison.Ordinal))
            {
                return "点击历史";
            }

            if (normalizedAction.Contains("发送", StringComparison.Ordinal) ||
                normalizedAction.Contains("粘贴", StringComparison.Ordinal) ||
                normalizedAction.Contains("分段", StringComparison.Ordinal) ||
                normalizedAction.Contains("自动", StringComparison.Ordinal))
            {
                return "发送历史";
            }

            return "历史";
        }

        private bool IsStoreMarkedSent(string storeName)
        {
            if (string.IsNullOrWhiteSpace(storeName))
            {
                return false;
            }

            lock (_sentStoreLock)
            {
                return _sentStores.Contains(storeName);
            }
        }

        private void MarkStoreAsSent(string storeName, string action = "发送", string detail = "发送成功")
        {
            if (string.IsNullOrWhiteSpace(storeName))
            {
                return;
            }

            ClearStoreSentMark(storeName);

            lock (_sentStoreLock)
            {
                _sentStores.Add(storeName);
            }

            RecordStoreSendHistory(storeName, action, true, detail);
            RefreshStoreNodeHeader(storeName);
        }

        private void ClearStoreSentMark(string storeName, bool refreshHeader = false)
        {
            if (string.IsNullOrWhiteSpace(storeName))
            {
                return;
            }

            lock (_sentStoreLock)
            {
                _sentStores.Remove(storeName);
            }

            if (refreshHeader)
            {
                RefreshStoreNodeHeader(storeName);
            }
        }

        private string BuildStoreHeaderText(string storeName, int itemCount, SendStrategy strategy)
        {
            string prefix = string.Empty;
            if (_manualReviewStores.Contains(storeName))
            {
                prefix = "❌ [需人工] ";
            }

            if (strategy == SendStrategy.TextSegmented &&
                TryGetSegmentFailureInfo(storeName, out var segmentFailureInfo) &&
                segmentFailureInfo != null)
            {
                int segmentSize = GetSegmentSizeForStore(storeName);
                int totalSegments = itemCount > 0
                    ? (int)Math.Ceiling((double)itemCount / segmentSize)
                    : 0;
                prefix += BuildSegmentProgressPrefix(segmentFailureInfo, totalSegments);
            }

            return $"{prefix}{storeName} ({itemCount}条)";
        }

        private void RefreshStoreNodeHeader(string storeName)
        {
            if (string.IsNullOrWhiteSpace(storeName))
            {
                return;
            }

            void RefreshAction()
            {
                if (StoreTreeView.ItemsSource is not IEnumerable<TreeViewNode> nodes)
                {
                    return;
                }

                var targetNode = nodes.FirstOrDefault(n => n.StoreName == storeName);
                if (targetNode == null)
                {
                    return;
                }

                RefreshStoreNodeHeader(targetNode);
            }

            if (Application.Current?.Dispatcher == null)
            {
                return;
            }

            if (Application.Current.Dispatcher.CheckAccess())
            {
                RefreshAction();
            }
            else
            {
                Application.Current.Dispatcher.Invoke(RefreshAction);
            }
        }

        private void RefreshStoreNodeHeader(TreeViewNode node)
        {
            if (node == null || string.IsNullOrWhiteSpace(node.StoreName) || node.StoreName == "FAIL_SEPARATOR")
            {
                return;
            }

            int itemCount = 0;
            lock (_dataLock)
            {
                if (_storeData.TryGetValue(node.StoreName, out var rows))
                {
                    itemCount = rows.Count;
                }
            }

            node.Header = BuildStoreHeaderText(node.StoreName, itemCount, node.Strategy);
        }

        private static bool IsSegmentSending(SegmentFailureInfo info)
        {
            return !string.IsNullOrWhiteSpace(info.Reason) &&
                   info.Reason.Contains("发送中", StringComparison.Ordinal);
        }

        private static bool IsSegmentCompleted(SegmentFailureInfo info, int totalSegments)
        {
            int safeTotal = Math.Max(1, totalSegments > 0 ? totalSegments : info.TotalSegments);
            int safeSent = Math.Max(0, Math.Min(info.SentSegments, safeTotal));
            bool reasonCompleted = !string.IsNullOrWhiteSpace(info.Reason) &&
                                   info.Reason.Contains("发送完成", StringComparison.Ordinal);
            return reasonCompleted || safeSent >= safeTotal;
        }

        private string BuildSegmentProgressPrefix(SegmentFailureInfo info, int totalSegments)
        {
            int safeTotal = Math.Max(1, totalSegments);
            if (IsSegmentCompleted(info, safeTotal))
            {
                return string.Empty;
            }

            if (IsSegmentSending(info))
            {
                int safeSent = Math.Max(0, Math.Min(info.SentSegments, safeTotal));
                return $"⏳ [分段进度 {safeSent}/{safeTotal}] ";
            }

            return "⚠️ [分段异常] ";
        }

        private static string BuildSegmentChildLabel(int segmentNumber, int totalSegments, int itemCount, SegmentFailureInfo? info)
        {
            string baseText = $"第 {segmentNumber}/{totalSegments} 段 (共 {itemCount} 条) - 点击复制此段";
            if (info == null)
            {
                return $"📄 {baseText}";
            }

            int safeTotal = Math.Max(1, totalSegments);
            if (IsSegmentCompleted(info, safeTotal))
            {
                return $"✅ {baseText}";
            }

            if (IsSegmentSending(info))
            {
                int sentSegments = Math.Max(0, Math.Min(info.SentSegments, safeTotal));
                int nextSegment = Math.Min(safeTotal, sentSegments + 1);

                if (segmentNumber <= sentSegments)
                {
                    return $"✅ {baseText}";
                }

                if (segmentNumber == nextSegment)
                {
                    return $"⏳ {baseText}";
                }

                return $"⏸ {baseText}";
            }

            if (segmentNumber < info.FailedSegment)
            {
                return $"✅ {baseText}";
            }

            if (segmentNumber == info.FailedSegment)
            {
                return $"⚠️ {baseText}";
            }

            return $"⏸ {baseText}";
        }

        private void UpdateSegmentProgressVisual(string storeName, int sentSegments, int totalSegments, int sentItems, int totalItems, string reason = "发送中")
        {
            if (string.IsNullOrWhiteSpace(storeName))
            {
                return;
            }

            int safeTotalSegments = Math.Max(1, totalSegments);
            int safeSentSegments = Math.Max(0, Math.Min(sentSegments, safeTotalSegments));
            int nextSegment = Math.Max(1, Math.Min(safeTotalSegments, safeSentSegments + 1));
            int safeTotalItems = Math.Max(0, totalItems);
            int safeSentItems = Math.Max(0, Math.Min(sentItems, safeTotalItems));
            string safeReason = string.IsNullOrWhiteSpace(reason) ? "发送中" : reason;

            lock (_segmentFailureLock)
            {
                _segmentFailureInfos[storeName] = new SegmentFailureInfo
                {
                    FailedSegment = nextSegment,
                    TotalSegments = safeTotalSegments,
                    SentSegments = safeSentSegments,
                    SentItems = safeSentItems,
                    TotalItems = safeTotalItems,
                    Reason = safeReason
                };
            }

            if (Application.Current?.Dispatcher == null)
            {
                return;
            }

            Action refreshAction = () =>
            {
                if (StoreTreeView.ItemsSource is not IEnumerable<TreeViewNode> nodes)
                {
                    return;
                }

                var targetNode = nodes.FirstOrDefault(n =>
                    n != null &&
                    string.Equals(n.StoreName, storeName, StringComparison.Ordinal) &&
                    n.Strategy == SendStrategy.TextSegmented);

                if (targetNode != null)
                {
                    ApplySegmentFailureProgressToNode(targetNode);
                }
            };

            if (Application.Current.Dispatcher.CheckAccess())
            {
                refreshAction();
            }
            else
            {
                Application.Current.Dispatcher.Invoke(refreshAction);
            }
        }

        private void ApplySegmentFailureProgressToNode(TreeViewNode node)
        {
            if (node == null || node.Strategy != SendStrategy.TextSegmented || string.IsNullOrWhiteSpace(node.StoreName))
            {
                return;
            }

            if (!TryGetSegmentFailureInfo(node.StoreName, out var info) || info == null)
            {
                return;
            }

            int totalSegments = Math.Max(1, node.Children.Count);
            int itemCount = 0;
            lock (_dataLock)
            {
                if (_storeData.TryGetValue(node.StoreName, out var rows))
                {
                    itemCount = rows.Count;
                }
            }

            node.Header = BuildStoreHeaderText(node.StoreName, itemCount, node.Strategy);

            for (int i = 0; i < node.Children.Count; i++)
            {
                var child = node.Children[i];
                int childItemCount = child.RawData?
                    .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Length ?? 0;
                child.Text = BuildSegmentChildLabel(i + 1, totalSegments, childItemCount, info);
            }
        }














        private void SimulatePaste()
        {
            try
            {
                _inputBackend.KeyChord(InputKey.LeftControl, InputKey.V);
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
                _inputBackend.KeyChord(InputKey.LeftAlt, InputKey.S);
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
            // ✅ 获取文件的最后修改时间
            DateTime fileLastModified = File.GetLastWriteTime(filePath);
            
            // ✅ 延迟状态恢复：需要先检测模式，再决定从哪个 FileState 恢复
            // 这里只做变量声明，实际恢复在模式检测后执行
            bool shouldRestoreState = false;
            FileState stateToRestore = null;

            // 1. 清除内存中的旧数据
            lock (_dataLock)
            {
                _storeData.Clear();
                _exportedFilePaths.Clear();
            }
            
            // ✅ 清空运行时状态集合
            _failedStores.Clear();
            _manualReviewStores.Clear();
            lock (_segmentFailureLock) { _segmentFailureInfos.Clear(); }
            lock (_sentStoreLock) { _sentStores.Clear(); }
            
            // ✅ [Fix] 重置模式标志，确保每次加载文件时从干净状态开始
            _isIssueMode = false;
            _isCustomMessageMode = false;
            _activeIssueSegmentStartCount = Math.Max(2, GetDefaultSegmentSize());



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
                    int colCount = worksheet.Dimension.End.Column;
                    _lastLoadedFilePath = filePath;
                    _lastLoadedColumnCount = colCount;
                    
                    // 🔍 [DEBUG] 输出检测到的列数，帮助调试
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Excel 检测: rowCount={rowCount}, colCount={colCount}");

                    FileParseOverride? fileParseOverride = FindFileParseOverride(filePath);
                    string parseMode = FileParseModes.Normalize(fileParseOverride?.ParseMode);
                    bool useManualTableParse = false;
                    int manualTrackingColumn = 1;
                    int manualStoreColumn = 2;
                    int manualColumnCount = 0;
                    _activeTailMessage = _searchConfig?.FixedMessage ?? string.Empty;
                    _activeIssueSegmentStartCount = ResolveIssueSegmentStartCount(fileParseOverride);
                    DebugLogManager.Log("ParseMode", $"当前问题件分段起始条数={_activeIssueSegmentStartCount} | 文件={Path.GetFileName(filePath)}");

                    if (string.Equals(parseMode, FileParseModes.Magician, StringComparison.Ordinal))
                    {
                        _isIssueMode = false;
                        _isCustomMessageMode = false;
                        _activeTailMessage = fileParseOverride?.TailMessage?.Trim();
                        if (string.IsNullOrWhiteSpace(_activeTailMessage))
                        {
                            _activeTailMessage = _searchConfig?.FixedMessage ?? string.Empty;
                        }

                        System.Diagnostics.Debug.WriteLine("[模式识别] 命中手动规则：强制【魔术师格式（两列表格）】");
                    }
                    else if (string.Equals(parseMode, FileParseModes.Issue, StringComparison.Ordinal))
                    {
                        manualTrackingColumn = Math.Min(Math.Max(1, fileParseOverride?.TrackingColumn ?? 1), Math.Max(1, colCount));
                        manualStoreColumn = Math.Min(Math.Max(1, fileParseOverride?.StoreColumn ?? 2), Math.Max(1, colCount));

                        manualColumnCount = colCount;
                        _isIssueMode = true;
                        _isCustomMessageMode = false;
                        useManualTableParse = manualColumnCount >= 2;
                        if (!useManualTableParse)
                        {
                            _isIssueMode = false;
                        }

                        if (useManualTableParse)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                $"[模式识别] 命中手动规则：问题件格式，运单列={manualTrackingColumn}，店铺列={manualStoreColumn}，分段起始条数={_activeIssueSegmentStartCount}，有效列数={manualColumnCount}");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine(
                                "[模式识别] 问题件规则列数不足2列，回退为【未发货模式】");
                        }
                    }
                    else
                    {
                        // ✅ 自动识别模式（稳健版）
                        // 优先级1：文件名关键字匹配（仲裁/赔付/监控/遗失）→ 问题件格式（运单号=第1列，商家名=第3列）
                        string fileNameOnly = Path.GetFileNameWithoutExtension(filePath);
                        bool matchedByFileName =
                            fileNameOnly.Contains("仲裁") ||
                            fileNameOnly.Contains("赔付") ||
                            fileNameOnly.Contains("监控") ||
                            fileNameOnly.Contains("遗失");

                        if (matchedByFileName)
                        {
                            manualTrackingColumn = Math.Min(1, Math.Max(1, colCount));
                            manualStoreColumn    = Math.Min(3, Math.Max(1, colCount));
                            manualColumnCount    = colCount;
                            _isIssueMode         = true;
                            _isCustomMessageMode = false;
                            useManualTableParse  = manualColumnCount >= 2;
                            if (!useManualTableParse) _isIssueMode = false;
                            System.Diagnostics.Debug.WriteLine(
                                $"[模式识别] 文件名命中关键字（{fileNameOnly}）：自动切换为【问题件格式】，运单列=1，店铺列=3");
                        }
                        else
                        {
                        // 优先级2：列数/表头内容检测
                        // - 问题件：5列且有问题件表头特征，或第5列存在真实数据
                        // - 自定义话术：至少4列且第3/4列存在真实数据
                        // - 其它：未发货模式
                        var header2 = GetSafeText(worksheet.Cells[1, 2].Value).Trim();
                        var header3 = GetSafeText(worksheet.Cells[1, 3].Value).Trim();
                        var header4 = GetSafeText(worksheet.Cells[1, 4].Value).Trim();
                        bool looksLikeIssueHeader =
                            header2.Contains("问题") || header2.Contains("类型") ||
                            header3.Contains("问题") || header3.Contains("原因") ||
                            header4.Contains("店铺");

                        bool hasDataInCol3Or4 = false;
                        bool hasDataInCol5 = false;
                        int sampleEndRow = Math.Min(rowCount, 30);
                        for (int row = 2; row <= sampleEndRow; row++)
                        {
                            if (!hasDataInCol3Or4)
                            {
                                string col3 = GetSafeText(worksheet.Cells[row, 3].Value).Trim();
                                string col4 = GetSafeText(worksheet.Cells[row, 4].Value).Trim();
                                hasDataInCol3Or4 = !string.IsNullOrWhiteSpace(col3) || !string.IsNullOrWhiteSpace(col4);
                            }

                            if (!hasDataInCol5)
                            {
                                string col5 = GetSafeText(worksheet.Cells[row, 5].Value).Trim();
                                hasDataInCol5 = !string.IsNullOrWhiteSpace(col5);
                            }

                            if (hasDataInCol3Or4 && hasDataInCol5)
                            {
                                break;
                            }
                        }

                        if (colCount >= 5 && (looksLikeIssueHeader || hasDataInCol5))
                        {
                            _isIssueMode = true;
                            _isCustomMessageMode = false;
                            _activeTailMessage = _searchConfig?.FixedMessage ?? string.Empty;
                            System.Diagnostics.Debug.WriteLine("[模式识别] 检测到问题件结构，切换为【问题件模式】");
                        }
                        else if (colCount >= 4 && hasDataInCol3Or4)
                        {
                            _isCustomMessageMode = true;
                            _isIssueMode = false;
                            _activeTailMessage = _searchConfig?.FixedMessage ?? string.Empty;
                            System.Diagnostics.Debug.WriteLine("[模式识别] 检测到4列自定义结构，切换为【自定义话术模式】");
                        }
                        else
                        {
                            _isIssueMode = false;
                            _isCustomMessageMode = false;
                            _activeTailMessage = _searchConfig?.FixedMessage ?? string.Empty;
                            System.Diagnostics.Debug.WriteLine("[模式识别] 未匹配问题件/自定义结构，使用【未发货模式】");
                        }
                        } // end else（列数/表头检测分支）
                    }

                    string resolvedModeName = _isIssueMode ? "Issue" : (_isCustomMessageMode ? "CustomMessage" : "Normal");
                    string overrideModeName = fileParseOverride == null ? "None" : FileParseModes.Normalize(fileParseOverride.ParseMode);
                    DebugLogManager.Log(
                        "ParseMode",
                        $"生效模式={resolvedModeName} | 覆盖规则={overrideModeName} | 问题件分段起始条数={_activeIssueSegmentStartCount} | 默认分段={GetDefaultSegmentSize()}");

                    // ✅ 根据模式选择对应的 FileState 进行状态恢复
                    if (_searchConfig != null)
                    {
                        string modeName = _isIssueMode ? "问题件" : (_isCustomMessageMode ? "自定义话术" : "未发货");
                        FileState savedState;
                        if (_isIssueMode)
                        {
                            _searchConfig.LastIssueFileState ??= new FileState();
                            savedState = _searchConfig.LastIssueFileState;
                        }
                        else if (_isCustomMessageMode)
                        {
                            _searchConfig.LastCustomMessageFileState ??= new FileState();
                            savedState = _searchConfig.LastCustomMessageFileState;
                        }
                        else
                        {
                            _searchConfig.LastFileState ??= new FileState();
                            savedState = _searchConfig.LastFileState;
                        }
                        
                        // 判断是否为"相同版本文件"：路径相同 且 修改时间相同 且 模式匹配
                        if (savedState != null &&
                            !string.IsNullOrEmpty(savedState.FilePath) &&
                            filePath.Equals(savedState.FilePath, StringComparison.OrdinalIgnoreCase) &&
                            Math.Abs((fileLastModified - savedState.LastModifiedTime).TotalSeconds) < 2 &&
                            savedState.IsIssueMode == _isIssueMode &&
                            savedState.IsCustomMessageMode == _isCustomMessageMode)
                        {
                            shouldRestoreState = true;
                            stateToRestore = savedState;
                            System.Diagnostics.Debug.WriteLine($"[文件状态] 检测到相同版本文件（模式={modeName}），将恢复状态");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[文件状态] 文件已更新或为新文件，不恢复旧状态");
                        }

                        // 创建新的状态对象并保存
                        var newState = new FileState
                        {
                            FilePath = filePath,
                            LastModifiedTime = fileLastModified,
                            LastSelectedStoreName = "",
                            FailedStores = new List<string>(),
                            ManualReviewStores = new List<string>(),
                            SegmentFailures = new List<SegmentFailureState>(),
                            DeletedStores = new List<string>(),
                            IsIssueMode = _isIssueMode,
                            IsCustomMessageMode = _isCustomMessageMode
                        };
                        
                        if (_isIssueMode)
                        {
                            _searchConfig.LastIssueFileState = newState;
                        }
                        else if (_isCustomMessageMode)
                        {
                            _searchConfig.LastCustomMessageFileState = newState;
                        }
                        else
                        {
                            _searchConfig.LastFileState = newState;
                        }
                        
                        // 向后兼容：同步更新旧字段
                        _searchConfig.LastOpenedFilePath = filePath;
                        _searchConfig.Save();
                    }
                    
                    // ✅ 如果需要恢复状态，加载重试区和需人工列表
                    if (shouldRestoreState && stateToRestore != null)
                    {
                        if (stateToRestore.FailedStores != null)
                            foreach (var s in stateToRestore.FailedStores) _failedStores.Add(s);
                        if (stateToRestore.ManualReviewStores != null)
                            foreach (var s in stateToRestore.ManualReviewStores) _manualReviewStores.Add(s);
                        if (stateToRestore.SegmentFailures != null)
                        {
                            lock (_segmentFailureLock)
                            {
                                foreach (var seg in stateToRestore.SegmentFailures)
                                {
                                    if (string.IsNullOrWhiteSpace(seg.StoreName))
                                    {
                                        continue;
                                    }

                                    _segmentFailureInfos[seg.StoreName] = new SegmentFailureInfo
                                    {
                                        FailedSegment = Math.Max(1, seg.FailedSegment),
                                        TotalSegments = Math.Max(1, seg.TotalSegments),
                                        SentSegments = Math.Max(0, seg.SentSegments),
                                        SentItems = Math.Max(0, seg.SentItems),
                                        TotalItems = Math.Max(0, seg.TotalItems),
                                        Reason = string.IsNullOrWhiteSpace(seg.Reason) ? "发送失败" : seg.Reason
                                    };
                                }
                            }
                        }
                        System.Diagnostics.Debug.WriteLine($"[文件状态] 已恢复：重试区 {_failedStores.Count} 个，需人工 {_manualReviewStores.Count} 个");
                    }

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string storeName = "";
                        string content = "";

                        if (useManualTableParse)
                        {
                            storeName = GetSafeText(worksheet.Cells[row, manualStoreColumn].Value).Trim();
                            if (string.IsNullOrWhiteSpace(storeName))
                            {
                                continue;
                            }

                            string trackingNo = GetSafeText(worksheet.Cells[row, manualTrackingColumn].Value).Trim();
                            if (string.IsNullOrWhiteSpace(trackingNo))
                            {
                                continue;
                            }

                            var parts = new List<string>(manualColumnCount);
                            for (int c = 1; c <= manualColumnCount; c++)
                            {
                                parts.Add(GetSafeText(worksheet.Cells[row, c].Value).Trim());
                            }
                            content = string.Join("\t", parts);
                        }
                        else if (_isIssueMode)
                        {
                            // 🆕 问题件模式 (5列)
                            // 格式：运单号(1) | 类型(2) | 原因(3) | 店铺(4) | 业务员(5)
                            storeName = GetSafeText(worksheet.Cells[row, 4].Value).Trim();
                            
                            // 拼接完整行 (Tab分隔，Excel粘贴友好)
                            var parts = new List<string>();
                            for (int c = 1; c <= 5; c++)
                            {
                                parts.Add(GetSafeText(worksheet.Cells[row, c].Value).Trim());
                            }
                            content = string.Join("\t", parts);
                        }
                        else if (_isCustomMessageMode)
                        {
                            // ✅ [重构] 自定义话术模式 (4列)
                            // 格式：运单号(1) | 话术(2) | 店铺(3) | 网点(4)
                            // 直接用商家名（第3列）分组，不再按话术拆分
                            storeName = GetSafeText(worksheet.Cells[row, 3].Value).Trim();

                            if (string.IsNullOrWhiteSpace(storeName))
                            {
                                continue;
                            }
                            
                            // 拼接完整行 (Tab分隔)
                            var parts = new List<string>();
                            for (int c = 1; c <= 4; c++)
                            {
                                parts.Add(GetSafeText(worksheet.Cells[row, c].Value).Trim());
                            }
                            content = string.Join("\t", parts);
                        }
                        else
                        {
                            // 🔙 原模式 (2列)
                            // 格式：运单号(1) | 店铺(2)
                            content = GetSafeText(worksheet.Cells[row, 1].Value).Trim(); // 只存运单号
                            storeName = GetSafeText(worksheet.Cells[row, 2].Value).Trim();
                        }

                        if (string.IsNullOrEmpty(storeName)) continue;

                        lock (_dataLock)
                        {
                            if (!_storeData.ContainsKey(storeName))
                            {
                                _storeData[storeName] = new List<string>();
                            }
                            _storeData[storeName].Add(content);
                        }
                    }
                }

                // 重置选中状态，避免指向不存在的旧索引
                _currentSelectedIndex = -1;
                _currentSelectedNode = null;

                // 处理并显示数据
                ProcessAndDisplayData();
                
                // ✅ 如果是相同版本文件，恢复上次选中的商家位置
                if (shouldRestoreState && stateToRestore != null && !string.IsNullOrEmpty(stateToRestore.LastSelectedStoreName))
                {
                    string storeNameToRestore = stateToRestore.LastSelectedStoreName;
                    Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        bool restored = RestoreSelection(storeNameToRestore, 0);
                        if (restored)
                        {
                            StatusTextBlock.Text = $"已恢复到上次选中: {storeNameToRestore} (包括重试区 {_failedStores.Count} 个)";
                        }
                    }, System.Windows.Threading.DispatcherPriority.Loaded);
                }

            }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = $"文件读取错误: {ex.Message}";
                    StoreTreeView.ItemsSource = null;
                    UpdateListProgressStatus();
                });
            }
        }

        /// <summary>
        /// ✅ 保存当前文件的完整状态（选中项、重试区、需人工、已删除）
        /// 根据当前模式保存到对应的 FileState（未发货/问题件/自定义话术）
        /// </summary>
        private void SaveFileState(string selectedStoreName = null)
        {
            if (_searchConfig == null) return;
            
            // ✅ 根据当前模式选择正确的状态对象
            FileState state;
            if (_isIssueMode)
            {
                _searchConfig.LastIssueFileState ??= new FileState();
                state = _searchConfig.LastIssueFileState;
            }
            else if (_isCustomMessageMode)
            {
                _searchConfig.LastCustomMessageFileState ??= new FileState();
                state = _searchConfig.LastCustomMessageFileState;
            }
            else
            {
                _searchConfig.LastFileState ??= new FileState();
                state = _searchConfig.LastFileState;
            }
            if (state == null) return;
            
            // 更新选中项（如果提供）
            if (!string.IsNullOrEmpty(selectedStoreName))
            {
                state.LastSelectedStoreName = selectedStoreName;
            }
            
            // 同步当前运行时状态到持久化对象
            state.IsIssueMode = _isIssueMode;
            state.IsCustomMessageMode = _isCustomMessageMode;
            state.FailedStores = _failedStores.ToList();
            state.ManualReviewStores = _manualReviewStores.ToList();
            lock (_segmentFailureLock)
            {
                state.SegmentFailures = _segmentFailureInfos
                    .Select(kvp => new SegmentFailureState
                    {
                        StoreName = kvp.Key,
                        FailedSegment = kvp.Value.FailedSegment,
                        TotalSegments = kvp.Value.TotalSegments,
                        SentSegments = kvp.Value.SentSegments,
                        SentItems = kvp.Value.SentItems,
                        TotalItems = kvp.Value.TotalItems,
                        Reason = kvp.Value.Reason
                    })
                    .ToList();
            }
            
            // 向后兼容：同步旧字段
            _searchConfig.LastSelectedStoreName = state.LastSelectedStoreName;
            
            // ✅ [优化] 防抖保存：延迟 1秒 执行写入，避免连续点击卡顿
            _saveDebounceTimer?.Stop();
            _saveDebounceTimer?.Start();
        }





        /// <summary>
        /// ✅ [修改版] 处理显示数据，并初始化“失败归档区”
        /// </summary>


        private void ProcessAndDisplayData()
        {
            List<KeyValuePair<string, List<string>>> sortedStores;

            var infoMap = _businessInfoList
                .Select(b => new
                {
                    Key = NormalizeStoreNameForBusinessInfo(b.StoreName),
                    Info = b
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.Key))
                .GroupBy(x => x.Key)
                .ToDictionary(g => g.Key, g => g.Select(x => x.Info).FirstOrDefault());

            lock (_dataLock)
            {
                sortedStores = _storeData
                    .Select(kvp => 
                    {
                        string realKey = kvp.Key;
                        var info = infoMap.ContainsKey(realKey) ? infoMap[realKey] : null;
                        return new { Kvp = kvp, Info = info };
                    })
                    .OrderByDescending(x => !string.IsNullOrEmpty(x.Info?.GroupName))
                    .ThenByDescending(x => x.Kvp.Value.Count > 100)
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

            var normalList = new List<KeyValuePair<string, List<string>>>();
            var failedList = new List<KeyValuePair<string, List<string>>>();

            foreach (var item in sortedStores)
            {
                if (_failedStores.Contains(item.Key))
                    failedList.Add(item);
                else
                    normalList.Add(item);
            }

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

            // --- 内部辅助函数：创建节点 ---
                TreeViewNode CreateNode(string storeName, List<string> trackingNumbers)
                {
                    var payloadMode = ResolveStorePayloadMode(storeName, trackingNumbers);
                    var strategy = ResolveSendStrategy(payloadMode, trackingNumbers.Count);
                    int itemCount = trackingNumbers.Count;

                    // ✅ [重构] 不再解析复合键，storeName 就是纯商家名
                    string displayStoreName = storeName;

                    SegmentFailureInfo? segmentFailureInfo = null;

                    if (strategy == SendStrategy.TextSegmented &&
                        TryGetSegmentFailureInfo(storeName, out var tmpFailureInfo) &&
                        tmpFailureInfo != null)
                    {
                        segmentFailureInfo = tmpFailureInfo;
                    }

                    var parentNode = new TreeViewNode
                    {
                        Header = BuildStoreHeaderText(storeName, itemCount, strategy),
                        StoreName = storeName,
                        Strategy = strategy
                    };

                    // 关联商家信息（群名、来源）
                    string lookupName = displayStoreName;
                    var busInfo = infoMap.ContainsKey(lookupName) ? infoMap[lookupName] : null;
                if (busInfo != null)
                {
                    parentNode.Source = busInfo.Source;
                    parentNode.GroupName = busInfo.GroupName;
                }

                switch (strategy)
                {
                    case SendStrategy.FileExcel:
                        // 2列模式 >100条：保持 Excel 文件发送
                        try
                        {
                            string filePath = null;
                            bool needGenerate = true;
                            lock (_dataLock)
                            {
                                if (_exportedFilePaths.TryGetValue(storeName, out string cachedPath) && File.Exists(cachedPath))
                                {
                                    filePath = cachedPath;
                                    needGenerate = false;
                                }
                            }

                            if (needGenerate)
                            {
                                filePath = CreateExcelFile(storeName, trackingNumbers, _exportDirectory);
                                lock (_dataLock) { _exportedFilePaths[storeName] = filePath; }
                            }

                            parentNode.Children.Add(new TreeViewNode 
                            { 
                                Text = "(单击复制名称，拖拽可导出文件)", 
                                RawData = filePath,
                                Strategy = SendStrategy.FileExcel,
                                Source = parentNode.Source // 传递 Source
                            });
                        }
                        catch (Exception ex)
                        {
                            parentNode.Children.Add(new TreeViewNode { Text = $"(文件创建失败: {ex.Message})" });
                        }
                        break;

                    case SendStrategy.TextSegmented:
                        // ✅ [新功能] 分段显示子节点
                        int segmentSize = GetSegmentSizeForPayloadMode(payloadMode);
                        int totalSegments = (int)Math.Ceiling((double)itemCount / segmentSize);
                        
                        for (int i = 0; i < totalSegments; i++)
                        {
                            int startIdx = i * segmentSize;
                            int count = Math.Min(segmentSize, itemCount - startIdx);
                            var segment = trackingNumbers.GetRange(startIdx, count);
                            
                            // 构建该段的完整文本
                            var sb = new StringBuilder();
                            foreach (var line in segment) sb.AppendLine(line);
                            string segmentText = sb.ToString();

                            parentNode.Children.Add(new TreeViewNode
                            {
                                Text = BuildSegmentChildLabel(i + 1, totalSegments, count, segmentFailureInfo),
                                StoreName = storeName,
                                RawData = segmentText, // ✅ 存入段数据
                                Strategy = SendStrategy.TextDirect,
                                Source = parentNode.Source // ✅ 传递 Source，确保点击子节点时能识别窗口类型
                            });
                        }

                        break;

                    default: // TextDirect
                        // 全部展开显示
                        foreach (var number in trackingNumbers)
                        {
                            string displayText = number;
                            // 适配多列数据的显示
                            if (number.Contains("\t"))
                            {
                                var parts = number.Split('\t');
                                if (parts.Length > 0)
                                {
                                    string info = "";
                                    if (parts.Length >= 2) info = parts[1];
                                    if (_isIssueMode && parts.Length > 2) info = parts[2];
                                    
                                    displayText = string.IsNullOrEmpty(info) ? parts[0] : $"{parts[0]} - {info}";
                                    
                                    if (displayText.Length > 60)
                                        displayText = displayText.Substring(0, 57) + "...";
                                }
                            }
                            parentNode.Children.Add(new TreeViewNode 
                            { 
                                Text = displayText, 
                                RawData = number, 
                                StoreName = storeName,
                                Strategy = SendStrategy.TextDirect,
                                Source = parentNode.Source // ✅ 传递 Source
                            });
                        }

                        break;
                }
                return parentNode;
            }

            // 1. 正常列表
            foreach (var kvp in normalList)
            {
                _treeViewCollection.Add(CreateNode(kvp.Key, kvp.Value));
            }

            // 2. 分隔符
            _failureNode = new TreeViewNode
            {
                Header = "========== 🚫 自动重试区 ==========",
                StoreName = "FAIL_SEPARATOR",
                Children = new ObservableCollection<TreeViewNode>()
            };
            _treeViewCollection.Add(_failureNode);

            // 3. 失败/重试列表
            foreach (var kvp in failedList)
            {
                _treeViewCollection.Add(CreateNode(kvp.Key, kvp.Value));
            }

            Application.Current.Dispatcher.Invoke(() =>
            {
                StoreTreeView.ItemsSource = _treeViewCollection;
                RebuildFlatNodeList();

                string filterInfo = _currentFilter.Count > 0 ? $"（已筛选 {_currentFilter.Count} 个关键词）" : "";
                StatusTextBlock.Text = $"处理完成，共显示 {sortedStores.Count} 个商家{filterInfo}";
                UpdateListProgressStatus();
            });
        }



        private string CreateExcelFile(string storeName, List<string> trackingNumbers, string outputDir)
        {
            var payloadMode = ResolveStorePayloadMode(storeName, trackingNumbers);

            string displayStoreName = storeName;

            string safeFileName = string.Join("_", displayStoreName.Split(Path.GetInvalidFileNameChars())).Trim();
            if (string.IsNullOrWhiteSpace(safeFileName))
            {
                safeFileName = "未命名商家";
            }

            string suffix = payloadMode switch
            {
                StorePayloadMode.Issue => "问题件明细",
                StorePayloadMode.CustomMessage => "",  // ✅ 自定义话术模式：文件名仅用店铺名
                _ => "未发货明细"
            };
            string worksheetName = SanitizeWorksheetName(payloadMode == StorePayloadMode.CustomMessage ? "明细" : suffix, "明细");

            // ✅ 统计条数用于文件名
            int itemCount = trackingNumbers?.Count ?? 0;
            string countSuffix = $"(共{itemCount}条)";

            string fileName = $"{safeFileName}{suffix}{countSuffix}.xlsx";

            string filePath = Path.Combine(outputDir, fileName);
            // 若文件已存在，直接删除旧文件（实现覆盖更新），防止产生 (2) 等重复副本
            if (File.Exists(filePath))
            {
                try
                {
                    File.Delete(filePath);
                }
                catch (IOException ioEx)
                {
                    // 若旧文件被占用，则可能需要加后缀，或者直接抛出异常提示用户关掉旧文件
                    // 在这里为了尽量保证不报错退出，还是保留一个兜底的回退（(新时间戳)）或者直接原样异常抛出
                    // 用户希望的是不要生成 (2)，可以尝试改用时间戳，或只允许覆盖，此处选择直接抛异常让用户在UI上得知占用。
                    throw new InvalidOperationException($"无法覆盖旧文件，可能该文件正在被其他程序（如 Excel）打开，请先关闭该文件：{fileName}", ioEx);
                }
            }

            using (var package = new ExcelPackage())
            {
                // 先用稳定名称创建，再重命名，避免动态名称导致创建失败
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                if (!string.Equals(worksheetName, "Sheet1", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        worksheet.Name = worksheetName;
                    }
                    catch (Exception renameEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[导出告警] 工作表重命名失败，回退 Sheet1。name='{worksheetName}', ex={renameEx.Message}");
                    }
                }
                
                if (payloadMode == StorePayloadMode.Issue)
                {
                    // ✅ 问题件模式：5列表头
                    worksheet.Cells[1, 1].Value = "运单号";
                    worksheet.Cells[1, 2].Value = "问题件类型";
                    worksheet.Cells[1, 3].Value = "问题件原因";
                    worksheet.Cells[1, 4].Value = "店铺";
                    worksheet.Cells[1, 5].Value = "业务员";
                    using (var headerRange = worksheet.Cells[1, 1, 1, 5])
                    {
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }
                    
                    // ✅ 问题件模式：解析 Tab 分隔的数据并写入5列
                    for (int i = 0; i < trackingNumbers.Count; i++)
                    {
                        var parts = trackingNumbers[i].Split('\t');
                        for (int col = 0; col < Math.Min(parts.Length, 5); col++)
                        {
                            worksheet.Cells[i + 2, col + 1].Value = parts[col];
                        }
                    }
                    
                    // ✅ 自动调整列宽
                    for (int col = 1; col <= 5; col++)
                    {
                        worksheet.Column(col).AutoFit(12);
                    }
                }
                else if (payloadMode == StorePayloadMode.CustomMessage)
                {
                    // ✅ 自定义话术模式：4列表头
                    worksheet.Cells[1, 1].Value = "运单号";
                    // 第2列表头按需求留空（仅用于内部话术列）
                    worksheet.Cells[1, 2].Value = string.Empty;
                    worksheet.Cells[1, 3].Value = "店铺";
                    worksheet.Cells[1, 4].Value = "网点";
                    
                    using (var headerRange = worksheet.Cells[1, 1, 1, 4])
                    {
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }
                    
                    // 写入数据
                    for (int i = 0; i < trackingNumbers.Count; i++)
                    {
                        var parts = trackingNumbers[i].Split('\t');
                        for (int col = 0; col < Math.Min(parts.Length, 4); col++)
                        {
                            worksheet.Cells[i + 2, col + 1].Value = parts[col];
                        }
                    }
                    
                    for (int col = 1; col <= 4; col++)
                    {
                        worksheet.Column(col).AutoFit(12);
                    }
                }
                else
                {
                    // ✅ 未发货模式：保持原有的2列格式
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
                        worksheet.Cells[i + 2, 2].Value = displayStoreName;
                    }
                    worksheet.Column(1).AutoFit(15);
                    worksheet.Column(2).AutoFit(20);
                }
                
                try
                {
                    package.SaveAs(new FileInfo(filePath));
                }
                catch (Exception saveEx)
                {
                    throw new InvalidOperationException(
                        $"保存Excel失败: mode={payloadMode}, sheet='{worksheet.Name}', path='{filePath}', store='{displayStoreName}', reason={saveEx.Message}",
                        saveEx);
                }
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

            // ✅ 进入筛选前，保存当前选中项（仅在首次进入筛选时保存）
            if (string.IsNullOrEmpty(_preFilterSelectedStoreName))
            {
                _preFilterSelectedStoreName = GetCurrentSelectedStoreName(out int selectedIndex);
                _preFilterSelectedIndex = selectedIndex;
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

            // ✅ 保存需要恢复的商家名
            string storeNameToRestore = _preFilterSelectedStoreName;
            int indexToRestore = _preFilterSelectedIndex;
            
            // ✅ 清空筛选前保存的商家名（已使用完毕）
            _preFilterSelectedStoreName = null;
            _preFilterSelectedIndex = -1;

            // ✅ 先重置选中状态
            _currentSelectedIndex = -1;
            _currentSelectedNode = null;

            ProcessAndDisplayData();

            // ✅ 清除筛选后，恢复到进入筛选前的选中项
            if (!string.IsNullOrEmpty(storeNameToRestore))
            {
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    bool restored = RestoreSelection(storeNameToRestore, indexToRestore);
                    StatusTextBlock.Text = restored
                        ? $"已恢复选中: '{storeNameToRestore}'"
                        : "已清空筛选，但未能恢复到先前选中项";
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }


        private void DeleteStoreButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string storeName)
            {
                // 1. 记录删除前的索引，用于恢复焦点
                // 如果当前没有选中项（极少情况），默认尝试选中第一个(0)
                int indexToRestore = _currentSelectedIndex >= 0 ? _currentSelectedIndex : 0;

                lock (_dataLock)
                {
                    _storeData.Remove(storeName);
                    _exportedFilePaths.Remove(storeName);

                    // ✅ [修复] 同步清理所有状态，防止幽灵数据
                    _failedStores.Remove(storeName);
                    _manualReviewStores.Remove(storeName);
                    lock (_segmentFailureLock)
                    {
                        _segmentFailureInfos.Remove(storeName);
                    }
                    lock (_sentStoreLock)
                    {
                        _sentStores.Remove(storeName);
                    }
                }

                // ✅ 保存删除后的状态
                SaveFileState();

                // 刷新列表
                ProcessAndDisplayData();

                // 2. 恢复选中状态 (此时 _flatNodeList 已经被 ProcessAndDisplayData 重建)
                // 使用 Dispatcher 确保在 UI 刷新完成后执行
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    if (_flatNodeList.Count > 0)
                    {
                        // 如果之前的索引超过了现在的最大索引（例如删除了最后一项），则指向新的最后一项
                        if (indexToRestore >= _flatNodeList.Count)
                        {
                            indexToRestore = _flatNodeList.Count - 1;
                        }

                        // 再次防止负数索引
                        if (indexToRestore < 0) indexToRestore = 0;

                        var nodeToSelect = _flatNodeList[indexToRestore];

                        // 执行选中
                        FocusAndSelectItem(nodeToSelect);
                        _currentSelectedNode = nodeToSelect;
                    }
                }, System.Windows.Threading.DispatcherPriority.Loaded);

                StatusTextBlock.Text = $"已删除商家: '{storeName}'";
            }
        }


        private void TreeViewItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(null);

            if (sender is FrameworkElement element && element.DataContext is TreeViewNode node && !string.IsNullOrWhiteSpace(node.StoreName))
            {
                if (FindVisualParent<TreeViewItem>(element) is TreeViewItem treeViewItem)
                {
                    if (!treeViewItem.IsSelected)
                    {
                        treeViewItem.IsSelected = true;
                    }
                    treeViewItem.Focus();
                }

                _currentSelectedNode = node;
                if (_flatNodeList.Count == 0)
                {
                    RebuildFlatNodeList();
                }
                if (_flatNodeList.Contains(node))
                {
                    _currentSelectedIndex = _flatNodeList.IndexOf(node);
                }
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
                    if (sender is FrameworkElement element && element.DataContext is TreeViewNode node && node.Strategy == SendStrategy.FileExcel)
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

        private async void TrackingNumber_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBlock textBlock && !string.IsNullOrEmpty(textBlock.Text))
            {
                string textToCopy = textBlock.Text.Trim();
                TreeViewNode? boundNode = textBlock.DataContext as TreeViewNode;
                // 如果 DataContext 有 RawData，优先复制 RawData（如：B列话术完整内容）
                if (boundNode != null && !string.IsNullOrEmpty(boundNode.RawData))
                {
                    textToCopy = boundNode.RawData;
                }
                else if (textToCopy.StartsWith("(") && textToCopy.EndsWith(")"))
                {
                    // 说明性文本（如“单击复制名称，拖拽可导出文件”）不参与复制
                    return;
                }

                if (Interlocked.CompareExchange(ref _childNodeCopyInProgress, 1, 0) == 1)
                {
                    e.Handled = true;
                    return;
                }

                try
                {
                    // 子项点击以“切换流畅”为优先，使用更轻量的剪贴板重试策略。
                    if (await SetClipboardWithRetryAsync(textToCopy, maxAttempts: 10, retryDelayMs: 18))
                    {
                        int lineCount = CountContentLines(textToCopy);
                        if (lineCount > 1 || textToCopy.Length > 120)
                        {
                            StatusTextBlock.Text = $"✅ 已复制内容（{lineCount}行，右键可预览）";
                        }
                        else
                        {
                            StatusTextBlock.Text = $"✅ 已复制: {BuildShortCopyPreview(textToCopy)}";
                        }
                    }
                    else
                    {
                        StatusTextBlock.Text = "复制失败";
                    }
                }
                catch (Exception ex)
                {
                    StatusTextBlock.Text = $"复制失败: {ex.Message}";
                }
                finally
                {
                    Interlocked.Exchange(ref _childNodeCopyInProgress, 0);
                }
                e.Handled = true;
            }
        }

        private void TrackingNumber_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not TextBlock textBlock || string.IsNullOrWhiteSpace(textBlock.Text))
            {
                return;
            }

            TreeViewNode? node = textBlock.DataContext as TreeViewNode;
            if (node == null)
            {
                return;
            }

            // 右键时先选中当前子项，避免仍停留在上一次左键选中项
            if (FindVisualParent<TreeViewItem>(textBlock) is TreeViewItem treeViewItem)
            {
                treeViewItem.IsSelected = true;
                treeViewItem.Focus();
            }
            _currentSelectedNode = node;
            if (_flatNodeList.Contains(node))
            {
                _currentSelectedIndex = _flatNodeList.IndexOf(node);
            }

            string content = node.RawData;
            if (string.IsNullOrWhiteSpace(content))
            {
                content = textBlock.Text.Trim();
            }

            if (string.IsNullOrWhiteSpace(content) ||
                (content.StartsWith("(") && content.EndsWith(")")))
            {
                return;
            }

            string title = string.IsNullOrWhiteSpace(node.Text) ? "内容预览" : node.Text;
            string? allContent = null;
            if (TryResolveRootNode(node, out TreeViewNode rootNode))
            {
                string allPreviewContent = BuildAllChildPreviewContent(rootNode);
                if (!string.IsNullOrWhiteSpace(allPreviewContent) &&
                    !string.Equals(allPreviewContent.Trim(), content.Trim(), StringComparison.Ordinal))
                {
                    allContent = allPreviewContent;
                }
            }

            OpenSegmentPreviewWindow(title, content, allContent);
            e.Handled = true;
        }

        private void OpenSegmentPreviewWindow(string title, string content, string? allContent = null)
        {
            int lineCount = CountContentLines(content);
            int allLineCount = CountContentLines(allContent);

            var previewWindow = new SegmentPreviewWindow(title, content, allContent)
            {
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };
            previewWindow.Show();
            previewWindow.Activate();

            if (allLineCount > lineCount)
            {
                StatusTextBlock.Text = $"📝 已打开内容预览（当前{lineCount}行，可显示全部{allLineCount}行）";
            }
            else
            {
                StatusTextBlock.Text = $"📝 已打开内容预览（{lineCount}行）";
            }
        }

        private static string BuildAllChildPreviewContent(TreeViewNode rootNode)
        {
            if (rootNode?.Children == null || rootNode.Children.Count == 0)
            {
                return string.Empty;
            }

            var sections = new List<string>();
            for (int i = 0; i < rootNode.Children.Count; i++)
            {
                TreeViewNode child = rootNode.Children[i];
                string sectionContent = ResolvePreviewContentForNode(child);
                if (string.IsNullOrWhiteSpace(sectionContent))
                {
                    continue;
                }

                sections.Add(sectionContent.Trim());
            }

            // 按用户要求：所有内容预览仅直接列出全部项，不添加分段标题或分隔符
            return string.Join(Environment.NewLine, sections).Trim();
        }

        private static string ResolvePreviewContentForNode(TreeViewNode node)
        {
            if (node == null)
            {
                return string.Empty;
            }

            string content = node.RawData;
            if (string.IsNullOrWhiteSpace(content))
            {
                content = node.Text?.Trim() ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                return string.Empty;
            }

            string trimmed = content.Trim();
            if (trimmed.StartsWith("(") && trimmed.EndsWith(")"))
            {
                return string.Empty;
            }

            return content;
        }

        private static int CountContentLines(string? content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return 0;
            }

            return content
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
                .Count(line => !string.IsNullOrWhiteSpace(line));
        }

        private static string BuildShortCopyPreview(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return string.Empty;
            }

            string oneLine = content.Replace("\r", " ").Replace("\n", " ").Trim();
            const int maxLength = 36;
            if (oneLine.Length <= maxLength)
            {
                return oneLine;
            }

            return oneLine.Substring(0, maxLength) + "...";
        }

        private void SuppressNextSelectionOsd(int count = 1)
        {
            if (count <= 0)
            {
                return;
            }

            Interlocked.Add(ref _suppressSelectionOsdCount, count);
        }

        private bool TryConsumeSuppressedSelectionOsd()
        {
            while (true)
            {
                int current = Volatile.Read(ref _suppressSelectionOsdCount);
                if (current <= 0)
                {
                    return false;
                }

                if (Interlocked.CompareExchange(ref _suppressSelectionOsdCount, current - 1, current) == current)
                {
                    return true;
                }
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
                bool suppressOsdForThisSelection = TryConsumeSuppressedSelectionOsd();

                if (node.StoreName != "FAIL_SEPARATOR")
                {
                    if (_searchConfig.EnableOsdWindow && !suppressOsdForThisSelection)
                    {
                        string seqPrefix = string.Empty;
                        string lastTracking = string.Empty;
                        string displayStoreName = node.StoreName;
                        string groupName = node.GroupName;

                        // O(1) 直接查表找父节点
                        _childParentMap.TryGetValue(node, out TreeViewNode parentNode);

                        if (parentNode != null)
                        {
                            // 当前是子节点（分段），取父节点的商家名和群名
                            displayStoreName = parentNode.StoreName;
                            groupName = parentNode.GroupName;

                            int listIdx = _flatNodeList.IndexOf(parentNode);
                            int childIdx = parentNode.Children.IndexOf(node);
                            int totalSeg = parentNode.Children.Count;

                            if (childIdx >= 0 && totalSeg > 0)
                            {
                                string listPart = listIdx >= 0 ? $"[{listIdx + 1}] " : string.Empty;
                                seqPrefix = $"{listPart}第 {childIdx + 1}/{totalSeg} 段";

                                // 提取该分段末条运单号
                                string rawData = node.RawData ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(rawData))
                                {
                                    var segLines = rawData.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                                    if (segLines.Length > 0)
                                    {
                                        string lastLine = segLines[segLines.Length - 1];
                                        int tabIdx = lastLine.IndexOf('\t');
                                        lastTracking = tabIdx > 0 ? lastLine.Substring(0, tabIdx).Trim() : lastLine.Trim();
                                    }
                                }
                            }
                        }
                        else
                        {
                            // 主节点：只显示列表序号
                            int midx = _flatNodeList.IndexOf(node);
                            if (midx >= 0)
                            {
                                seqPrefix = $"[{midx + 1}] ";
                            }
                        }

                        OsdWindow.ShowMessage(displayStoreName, seqPrefix, groupName, lastTracking);
                    }
                }

                // ✅ 切换到新项时，检查是否需要重置粘贴状态
                if (_lastPastedStoreName != node.StoreName)
                {
                    _currentItemPasted = false;
                }

                if (_flatNodeList.Contains(node))
                {
                    _currentSelectedIndex = _flatNodeList.IndexOf(node);
                }
                UpdateListProgressStatus();
                
                // ✅ 取消以前的同步调用，改用防抖保存
                if (node.StoreName != "FAIL_SEPARATOR")
                {
                    _pendingSaveStoreName = node.StoreName;
                    _selectionSaveDebounceTimer.Stop();
                    _selectionSaveDebounceTimer.Start();
                }

                SyncBusInfoManagerWithCurrentSelection();

                // 启动复制文字的防抖触发
                _pendingCopyNode = node;
                _selectionCopyDebounceTimer.Stop();
                _selectionCopyDebounceTimer.Start();
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

                        string displayStoreName = storeName;
                        targetNode.Header = BuildStoreHeaderText(storeName, trackingCount, targetNode.Strategy);

                        StatusTextBlock.Text = $"[OCR] ✅ 已更新商家 '{displayStoreName}' 的群名为: {groupName}";
                    }
                }
            });
        }






        #region 剪贴板操作

        private bool TryGetPreferredSearchCopyText(TreeViewNode? node, out string copyText, out string copyType)
        {
            copyText = string.Empty;
            copyType = "商家名";

            if (node == null)
            {
                return false;
            }

            TreeViewNode sourceNode = node;
            if (TryResolveRootNode(node, out TreeViewNode rootNode))
            {
                sourceNode = rootNode;
            }

            string storeName = NormalizeStoreNameForSearch(sourceNode.StoreName);
            if (string.IsNullOrWhiteSpace(storeName))
            {
                storeName = sourceNode.StoreName?.Trim() ?? string.Empty;
            }

            string groupName = sourceNode.GroupName?.Trim() ?? string.Empty;
            bool hasGroupName = !string.IsNullOrWhiteSpace(groupName);

            if (!_isStoreMode && hasGroupName)
            {
                copyText = groupName;
                copyType = "群名";
                return true;
            }

            copyText = storeName;
            copyType = "商家名";
            return !string.IsNullOrWhiteSpace(copyText);
        }

        private void CopyPreferredSearchText(TreeViewNode node)
        {
            Task.Run(async () =>
            {
                try
                {
                    if (_isAutoRunning || Volatile.Read(ref _clipboardSearchGuard) > 0) return;

                    if (!TryGetPreferredSearchCopyText(node, out string copyText, out string copyType))
                    {
                        throw new Exception("无可复制内容");
                    }

                    bool fallbackToStoreName = !_isStoreMode &&
                                               string.Equals(copyType, "商家名", StringComparison.Ordinal);

                    if (!await SetClipboardWithRetryAsync(copyText)) throw new Exception("剪贴板被占用");
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        string fallbackTip = fallbackToStoreName ? "（当前群名模式，无群名已回退商家名）" : string.Empty;
                        StatusTextBlock.Text = $"已复制{copyType}: '{copyText}'{fallbackTip}";
                    });
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
                    if (_isAutoRunning || Volatile.Read(ref _clipboardSearchGuard) > 0) return;

                    List<string> trackingNumbers;
                    lock (_dataLock)
                    {
                        if (!_storeData.TryGetValue(storeName, out trackingNumbers)) throw new Exception("未找到商家数据");
                        trackingNumbers = trackingNumbers.ToList();
                    }

                    var payloadMode = ResolveStorePayloadMode(storeName, trackingNumbers);
                    bool isNormalMode = payloadMode == StorePayloadMode.Normal;
                    string displayStoreName = storeName;

                    var sb = new StringBuilder();
                    // 统一规则：仅普通(2列)模式在开头追加店铺名；4列/5列均不追加
                    if (isNormalMode)
                    {
                        sb.AppendLine(displayStoreName);
                    }

                    foreach (var num in trackingNumbers) sb.AppendLine(num);
                    
                    // ✅ 仅普通(2列)模式追加固定话术
                    if (isNormalMode)
                    {
                        sb.AppendLine(GetCurrentTailMessage());
                    }

                    if (!await SetClipboardWithRetryAsync(sb.ToString())) throw new Exception("剪贴板被占用");

                    Application.Current.Dispatcher.Invoke(() => StatusTextBlock.Text = $"✅ 已复制 '{displayStoreName}' 的完整信息 ({trackingNumbers.Count} 条单号)");
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

        private async Task<bool> SetClipboardWithRetryAsync(object data, int maxAttempts = 25, int retryDelayMs = 20)
        {
            if (maxAttempts <= 0)
            {
                maxAttempts = 1;
            }

            if (retryDelayMs < 0)
            {
                retryDelayMs = 0;
            }

            for (int i = 0; i < maxAttempts; i++)
            {
                try
                {
                    bool success = await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        try
                        {
                            Clipboard.SetDataObject(data, false);
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
                await Task.Delay(retryDelayMs);
            }
            return false;
        }

        // ✅ 新增：选中事件自动滚动
        private void TreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem item && item.DataContext is TreeViewNode node)
            {
                _currentSelectedNode = node;
                if (_flatNodeList.Count == 0)
                {
                    RebuildFlatNodeList();
                }
                if (_flatNodeList.Contains(node))
                {
                    _currentSelectedIndex = _flatNodeList.IndexOf(node);
                }
            }
        }

        /// <summary>
        /// 滚动时防止虚拟化导致当前选中项状态被意外清空。
        /// </summary>
        private void StoreTreeView_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            // 不在滚动时强制恢复旧选中项，避免拖动滚动条后锁死在上一个节点。
            if (StoreTreeView.SelectedItem is TreeViewNode selectedNode &&
                !string.IsNullOrWhiteSpace(selectedNode.StoreName))
            {
                _currentSelectedNode = selectedNode;
                if (_flatNodeList.Count == 0)
                {
                    RebuildFlatNodeList();
                }
                if (_flatNodeList.Contains(selectedNode))
                {
                    _currentSelectedIndex = _flatNodeList.IndexOf(selectedNode);
                }
            }
        }

        #endregion

        #region UI辅助

        private void FilterToggleButton_Checked(object sender, RoutedEventArgs e)
        {
            FilterPanel.Visibility = Visibility.Visible;

            // ✅ 修复：展开筛选面板后也需要同步选中状态
            if (_currentSelectedNode != null)
            {
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    SyncTreeViewSelection(_currentSelectedNode);
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }

        private void FilterToggleButton_Unchecked(object sender, RoutedEventArgs e)
        {
            FilterPanel.Visibility = Visibility.Collapsed;

            // ✅ 修复：收起筛选面板后重新同步选中状态
            // 由于 TreeView 使用虚拟化，收起面板后高度变化会导致容器重建，需要重新设置选中项
            if (_currentSelectedNode != null)
            {
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    SyncTreeViewSelection(_currentSelectedNode);
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }



        /// <summary>
        /// ✅ 修复：直接操作数据模型清除选中状态
        /// </summary>
        private void ClearAllTreeViewSelections()
        {
            foreach (var node in EnumerateAllTreeNodes())
            {
                if (node != _currentSelectedNode && node.IsSelected)
                {
                    node.IsSelected = false;
                }
            }
        }

        private IEnumerable<TreeViewNode> EnumerateAllTreeNodes()
        {
            if (_treeViewCollection == null)
            {
                yield break;
            }

            foreach (var root in _treeViewCollection)
            {
                foreach (var node in EnumerateTreeNodesRecursive(root))
                {
                    yield return node;
                }
            }
        }

        private static IEnumerable<TreeViewNode> EnumerateTreeNodesRecursive(TreeViewNode? node)
        {
            if (node == null)
            {
                yield break;
            }

            yield return node;

            if (node.Children == null)
            {
                yield break;
            }

            foreach (var child in node.Children)
            {
                foreach (var nested in EnumerateTreeNodesRecursive(child))
                {
                    yield return nested;
                }
            }
        }

        private void UpdateListProgressStatus()
        {
            void UpdateCore()
            {
                if (ListProgressTextBlock == null)
                {
                    return;
                }

                List<TreeViewNode> storeNodes = GetTopLevelStoreNodesSnapshot();
                int totalStores = storeNodes.Count;
                if (totalStores <= 0)
                {
                    ListProgressTextBlock.Text = "后续 0 项 / 共 0 商家";
                    return;
                }

                TreeViewNode? currentRoot = ResolveCurrentRootNodeForProgress(storeNodes);
                int currentIndex = currentRoot == null ? -1 : storeNodes.IndexOf(currentRoot);
                if (currentIndex < 0 && currentRoot != null)
                {
                    currentIndex = storeNodes.FindIndex(n =>
                        string.Equals(n.StoreName, currentRoot.StoreName, StringComparison.Ordinal));
                }

                int remaining = currentIndex >= 0
                    ? Math.Max(0, totalStores - currentIndex - 1)
                    : totalStores;
                ListProgressTextBlock.Text = $"后续 {remaining} 项 / 共 {totalStores} 商家";
            }

            if (Application.Current?.Dispatcher == null)
            {
                return;
            }

            if (Application.Current.Dispatcher.CheckAccess())
            {
                UpdateCore();
            }
            else
            {
                Application.Current.Dispatcher.Invoke(UpdateCore);
            }
        }

        private List<TreeViewNode> GetTopLevelStoreNodesSnapshot()
        {
            IEnumerable<TreeViewNode>? source = _treeViewCollection;
            if (source == null && StoreTreeView.ItemsSource is IEnumerable<TreeViewNode> uiSource)
            {
                source = uiSource;
            }

            if (source == null)
            {
                return new List<TreeViewNode>();
            }

            return source
                .Where(node => IsSelectableNode(node))
                .ToList();
        }

        private TreeViewNode? ResolveCurrentRootNodeForProgress(List<TreeViewNode> storeNodes)
        {
            TreeViewNode? currentNode = StoreTreeView.SelectedItem as TreeViewNode ?? _currentSelectedNode;
            if ((currentNode == null || !IsSelectableNode(currentNode)) &&
                _currentSelectedIndex >= 0 &&
                _currentSelectedIndex < _flatNodeList.Count)
            {
                currentNode = _flatNodeList[_currentSelectedIndex];
            }

            if (currentNode == null || !IsSelectableNode(currentNode))
            {
                return null;
            }

            if (storeNodes.Contains(currentNode))
            {
                return currentNode;
            }

            if (TryResolveRootNode(currentNode, out TreeViewNode rootNode) && IsSelectableNode(rootNode))
            {
                return rootNode;
            }

            return storeNodes.FirstOrDefault(n =>
                string.Equals(n.StoreName, currentNode.StoreName, StringComparison.Ordinal));
        }

        /// <summary>
        /// ✅ 修复：仅负责滚动到视图
        /// </summary>
        private void BringNodeIntoViewByIndex(int index)
        {
            if (index < 0 || index >= _flatNodeList.Count) return;

            var node = _flatNodeList[index];
            node.IsSelected = true; // 确保数据层面被选中

            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                StoreTreeView.UpdateLayout();
                // ✅ 修复：使用支持虚拟化的滚动方法
                ScrollToNode(node);
            }, System.Windows.Threading.DispatcherPriority.Background);
        }

        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            return parentObject as T ?? FindVisualParent<T>(parentObject);
        }

        // ✅ 新增：查找子元素辅助方法
        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T t) return t;
                var result = FindVisualChild<T>(child);
                if (result != null) return result;
            }
            return null;
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
                    process.ProcessName.Equals("Weixin", StringComparison.OrdinalIgnoreCase) ||
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

        private static bool IsProcessNameForApp(string processName, bool isWework)
        {
            if (string.IsNullOrWhiteSpace(processName))
            {
                return false;
            }

            if (isWework)
            {
                return processName.Equals("WXWork", StringComparison.OrdinalIgnoreCase);
            }

            return processName.Equals("WeChat", StringComparison.OrdinalIgnoreCase) ||
                   processName.Equals("Weixin", StringComparison.OrdinalIgnoreCase);
        }



        private static string NormalizeStoreNameForSearch(string storeName)
        {
            if (string.IsNullOrWhiteSpace(storeName))
            {
                return string.Empty;
            }

            return storeName.Trim();
        }

        private static string NormalizeStoreNameForBusinessInfo(string storeName)
        {
            return NormalizeStoreNameForSearch(storeName);
        }



        private static int GetTabColumnCount(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return 0;
            }

            return value.Split('\t').Length;
        }

        /// <summary>
        /// 根据数据内容判断负载模式（纯靠 Tab 列数，不再依赖复合键）
        /// </summary>
        private StorePayloadMode ResolveStorePayloadMode(string storeName, List<string> rows)
        {
            // 手动/自动已识别为问题件时，优先按问题件处理（确保使用问题件分段阈值）
            if (_isIssueMode)
            {
                return StorePayloadMode.Issue;
            }

            if (_isCustomMessageMode)
            {
                return StorePayloadMode.CustomMessage;
            }

            string? firstRow = rows?.FirstOrDefault(r => !string.IsNullOrWhiteSpace(r));
            int columnCount = GetTabColumnCount(firstRow);

            if (columnCount >= 5)
            {
                return StorePayloadMode.Issue;
            }

            if (columnCount == 4)
            {
                return StorePayloadMode.CustomMessage;
            }

            return StorePayloadMode.Normal;
        }

        /// <summary>
        /// 根据模式和数据量决定发送策略
        /// </summary>
        private SendStrategy ResolveSendStrategy(StorePayloadMode mode, int itemCount)
        {
            // 2列模式：超100条走文件
            if (mode == StorePayloadMode.Normal && itemCount > 100)
                return SendStrategy.FileExcel;

            // 4列/5列模式：超过对应分段起始条数走分段（问题件可由解析设置覆盖）
            if ((mode == StorePayloadMode.CustomMessage || mode == StorePayloadMode.Issue)
                && itemCount > GetSegmentSizeForPayloadMode(mode))
                return SendStrategy.TextSegmented;

            // 其余：一次性文本
            return SendStrategy.TextDirect;
        }

        private static string SanitizeWorksheetName(string worksheetName, string fallback)
        {
            string safeName = (worksheetName ?? string.Empty).Trim();
            foreach (char c in new[] { ':', '\\', '/', '?', '*', '[', ']' })
            {
                safeName = safeName.Replace(c, '_');
            }

            if (string.IsNullOrWhiteSpace(safeName))
            {
                safeName = (fallback ?? string.Empty).Trim();
                foreach (char c in new[] { ':', '\\', '/', '?', '*', '[', ']' })
                {
                    safeName = safeName.Replace(c, '_');
                }
            }

            if (safeName.Length > 31)
            {
                safeName = safeName.Substring(0, 31);
            }

            if (string.IsNullOrWhiteSpace(safeName))
            {
                safeName = "Sheet1";
            }

            return safeName;
        }

        private static string BuildCustomMessageFileName(string message, int maxLength = 80)
        {
            string fileName = SanitizeFileNamePart(message, "未命名话术");
            return TruncateWithEllipsis(fileName, maxLength);
        }

        private static string BuildCustomStoreMessageFileName(string storeName, string message, int maxLength = 120)
        {
            string safeStoreName = SanitizeFileNamePart(storeName, "未命名商家");
            string safeMessage = SanitizeFileNamePart(message, string.Empty);

            if (string.IsNullOrWhiteSpace(safeMessage))
            {
                return TruncateWithEllipsis(safeStoreName, maxLength);
            }

            string combined = $"{safeStoreName}-{safeMessage}";
            if (combined.Length <= maxLength)
            {
                return combined;
            }

            // 优先保留完整店铺名，仅截断话术部分。
            int messageMaxLength = maxLength - safeStoreName.Length - 1;
            if (messageMaxLength > 0)
            {
                safeMessage = TruncateWithEllipsis(safeMessage, messageMaxLength);
                return $"{safeStoreName}-{safeMessage}";
            }

            return TruncateWithEllipsis(combined, maxLength);
        }

        private static string SanitizeFileNamePart(string text, string fallback)
        {
            string safe = (text ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();

            safe = string.Join("_", safe.Split(Path.GetInvalidFileNameChars())).Trim();
            if (string.IsNullOrWhiteSpace(safe))
            {
                safe = fallback;
            }

            return safe;
        }

        private static string TruncateWithEllipsis(string text, int maxLength)
        {
            string safe = text ?? string.Empty;
            if (maxLength <= 0)
            {
                return string.Empty;
            }

            if (safe.Length <= maxLength)
            {
                return safe;
            }

            if (maxLength <= 3)
            {
                return safe.Substring(0, maxLength);
            }

            return safe.Substring(0, maxLength - 3).TrimEnd() + "...";
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

            string storeName = _currentSelectedNode.StoreName; // 节点键（可能为 "商家名##话术"）
            string businessStoreName = NormalizeStoreNameForBusinessInfo(storeName);
            if (string.IsNullOrWhiteSpace(businessStoreName))
            {
                businessStoreName = storeName?.Trim() ?? string.Empty;
            }

            var existingInfo = _businessInfoList.FirstOrDefault(b => b.StoreName == businessStoreName);
            if (existingInfo == null && !string.Equals(storeName, businessStoreName, StringComparison.Ordinal))
            {
                // 兼容旧数据：历史上可能存了复合键
                existingInfo = _businessInfoList.FirstOrDefault(b => b.StoreName == storeName);
            }

            // 使用副本编辑，避免取消时意外污染内存对象
            BusinessInfo infoToEdit = existingInfo != null
                ? new BusinessInfo
                {
                    StoreName = businessStoreName,
                    GroupName = existingInfo.GroupName,
                    Source = existingInfo.Source
                }
                : new BusinessInfo { StoreName = businessStoreName };

            var editWindow = new EditBusInfoWindow(infoToEdit);
            bool? result = editWindow.ShowDialog();

            if (result == true)
            {
                BusinessInfo updatedInfo = editWindow.Info;

                // ✅ 更新业务信息列表
                _businessInfoList.RemoveAll(b =>
                    string.Equals(b.StoreName, updatedInfo.StoreName, StringComparison.Ordinal) ||
                    string.Equals(b.StoreName, storeName, StringComparison.Ordinal));

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

                StatusTextBlock.Text = $"✅ 已更新商家 '{updatedInfo.StoreName}' 的信息。";
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
                        string sanitizedJson = SanitizeJsonControlCharsInStrings(json, out bool hasSanitizedControlChars);

                        var deserializeOptions = new JsonSerializerOptions
                        {
                            AllowTrailingCommas = true,
                            ReadCommentHandling = JsonCommentHandling.Skip
                        };

                        _businessInfoList = JsonSerializer.Deserialize<List<BusinessInfo>>(sanitizedJson, deserializeOptions) ?? new List<BusinessInfo>();

                        if (hasSanitizedControlChars)
                        {
                            StatusTextBlock.Text = $"✅ 成功加载 {_businessInfoList.Count} 条商家群聊信息（已自动修复不可见控制字符）。";
                            return;
                        }
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

        private static string SanitizeJsonControlCharsInStrings(string json, out bool changed)
        {
            changed = false;
            if (string.IsNullOrEmpty(json))
            {
                return json;
            }

            var builder = new StringBuilder(json.Length);
            bool inString = false;
            bool isEscaped = false;

            foreach (char ch in json)
            {
                if (!inString)
                {
                    builder.Append(ch);
                    if (ch == '"')
                    {
                        inString = true;
                    }
                    continue;
                }

                if (isEscaped)
                {
                    builder.Append(ch);
                    isEscaped = false;
                    continue;
                }

                if (ch == '\\')
                {
                    builder.Append(ch);
                    isEscaped = true;
                    continue;
                }

                if (ch == '"')
                {
                    builder.Append(ch);
                    inString = false;
                    continue;
                }

                if (ch < 0x20)
                {
                    changed = true;
                    AppendEscapedControlChar(builder, ch);
                    continue;
                }

                builder.Append(ch);
            }

            return changed ? builder.ToString() : json;
        }

        private static void AppendEscapedControlChar(StringBuilder builder, char ch)
        {
            switch (ch)
            {
                case '\b':
                    builder.Append("\\b");
                    break;
                case '\t':
                    builder.Append("\\t");
                    break;
                case '\n':
                    builder.Append("\\n");
                    break;
                case '\f':
                    builder.Append("\\f");
                    break;
                case '\r':
                    builder.Append("\\r");
                    break;
                default:
                    builder.Append("\\u");
                    builder.Append(((int)ch).ToString("x4"));
                    break;
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

                    var storeName = ocrResult.StoreName; // 节点键（可能为 "商家名##话术"）
                    var businessStoreName = NormalizeStoreNameForBusinessInfo(storeName);
                    if (string.IsNullOrWhiteSpace(businessStoreName))
                    {
                        businessStoreName = storeName?.Trim() ?? string.Empty;
                    }

                    var existingInfo = _businessInfoList.FirstOrDefault(b => b.StoreName == businessStoreName);
                    if (existingInfo == null && !string.Equals(storeName, businessStoreName, StringComparison.Ordinal))
                    {
                        // 兼容旧数据：历史上可能存了复合键
                        existingInfo = _businessInfoList.FirstOrDefault(b => b.StoreName == storeName);
                    }

                    if (existingInfo != null && !string.IsNullOrWhiteSpace(existingInfo.GroupName))
                    {
                        //StatusTextBlock.Text = $"[OCR] 商家 '{storeName}' 已有群名，本次识别结果被忽略。";
                        return;
                    }

                    BusinessInfo targetInfo = existingInfo;
                    if (existingInfo != null)
                    {
                        existingInfo.StoreName = businessStoreName;
                        existingInfo.GroupName = ocrResult.GroupName;
                        existingInfo.Source = ocrResult.Source;
                    }
                    else
                    {
                        targetInfo = new BusinessInfo
                        {
                            StoreName = businessStoreName,
                            GroupName = ocrResult.GroupName,
                            Source = ocrResult.Source
                        };
                        _businessInfoList.Add(targetInfo);
                    }

                    _businessInfoList.RemoveAll(b =>
                        b != targetInfo &&
                        (string.Equals(b.StoreName, businessStoreName, StringComparison.Ordinal) ||
                         string.Equals(b.StoreName, storeName, StringComparison.Ordinal)));

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

                    StatusTextBlock.Text = $"[OCR] ✅ 已保存商家 '{businessStoreName}' 的群名";
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
        private bool RestoreSelection(string storeName, int fallbackIndex)
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
                    return true;
                }
                return false;
            }

            // 重建扁平列表
            RebuildFlatNodeList();

            // 在新列表中查找同名商家
            var targetNode = _flatNodeList.FirstOrDefault(n => n.StoreName == storeName);

            if (targetNode != null)
            {
                _currentSelectedIndex = _flatNodeList.IndexOf(targetNode);
                _currentSelectedNode = targetNode;

                FocusAndSelectItem(targetNode);
                return true;
            }

            if (fallbackIndex >= 0 && fallbackIndex < _flatNodeList.Count)
            {
                // 如果找不到原商家，使用备用索引
                _currentSelectedIndex = fallbackIndex;
                var node = _flatNodeList[fallbackIndex];
                _currentSelectedNode = node;
                FocusAndSelectItem(node);
                return true;
            }

            return false;
        }

        /// <summary>
        /// 获取当前可恢复的商家选中快照（优先当前节点，失败则使用当前索引）。
        /// </summary>
        private string GetCurrentSelectedStoreName(out int selectedIndex)
        {
            selectedIndex = -1;

            TreeViewNode selectedNode = _currentSelectedNode;
            if (selectedNode == null || string.IsNullOrEmpty(selectedNode.StoreName) || selectedNode.StoreName == "FAIL_SEPARATOR")
            {
                if (_flatNodeList.Count == 0)
                {
                    RebuildFlatNodeList();
                }

                if (_currentSelectedIndex >= 0 && _currentSelectedIndex < _flatNodeList.Count)
                {
                    selectedNode = _flatNodeList[_currentSelectedIndex];
                }
            }

            if (selectedNode == null || string.IsNullOrEmpty(selectedNode.StoreName) || selectedNode.StoreName == "FAIL_SEPARATOR")
            {
                return null;
            }

            selectedIndex = _flatNodeList.IndexOf(selectedNode);
            return selectedNode.StoreName;
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

            // 2. 数据驱动：设置选中状态
            ClearAllTreeViewSelections();
            node.IsSelected = true;

            // 3. 延迟执行滚动和焦点操作
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    // 延迟到界面空闲时再执行滚动
                    Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        try
                        {
                            StoreTreeView.UpdateLayout();
                            ScrollToNode(node);
                            
                            var container = StoreTreeView.ItemContainerGenerator.ContainerFromItem(node) as TreeViewItem;
                            if (container != null)
                            {
                                container.Focus();
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"选中项滚动失败: {ex.Message}");
                        }
                    }, System.Windows.Threading.DispatcherPriority.ContextIdle);

                    // 触发后续的业务逻辑（如复制）
                    TriggerCopyOperation(node);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"选中项定位失败: {ex.Message}");
                }

            }, System.Windows.Threading.DispatcherPriority.Loaded);
        }

        /// <summary>
        /// ✅ 新增：在虚拟化开启时安全滚动到指定节点
        /// </summary>
        private void ScrollToNode(TreeViewNode node)
        {
            if (node == null) return;
            
            // 尝试在整体树状结构中递归寻找并展开到该节点
            var tvi = FindAndExpandTreeViewItem(StoreTreeView, node);
            if (tvi != null)
            {
                tvi.BringIntoView();
            }
            else
            {
                // 如果实在没找到（如虚拟化极为深层且未生成），提供一个兜底方案：
                // 使用原先尝试提取 VirtualizingStackPanel 并反射调用 BringIndexIntoView 的方法
                int index = _flatNodeList.IndexOf(node);
                if (index >= 0)
                {
                    var vsp = FindVisualChild<VirtualizingStackPanel>(StoreTreeView);
                    if (vsp != null)
                    {
                        try
                        {
                            var method = typeof(VirtualizingStackPanel).GetMethod("BringIndexIntoView", 
                                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public);
                            if (method != null)
                            {
                                method.Invoke(vsp, new object[] { index });
                                StoreTreeView.UpdateLayout();
                                // 再次获取看能不能得到
                                var containerAfterScroll = StoreTreeView.ItemContainerGenerator.ContainerFromItem(node) as TreeViewItem;
                                containerAfterScroll?.BringIntoView();
                            }
                        }
                        catch { /* 忽略反射调用时的异常 */ }
                    }
                }
            }
        }

        /// <summary>
        /// 递归遍历 ItemsControl 寻找指定数据绑定的 TreeViewItem 并逐层展开
        /// </summary>
        private TreeViewItem? FindAndExpandTreeViewItem(ItemsControl itemsControl, TreeViewNode targetNode)
        {
            if (itemsControl == null || targetNode == null) return null;

            try
            {
                // 检查当前 itemsControl 是否已经包含该项
                var container = itemsControl.ItemContainerGenerator.ContainerFromItem(targetNode) as TreeViewItem;
                if (container != null) return container;

                // 还没找到，遍历子项
                for (int i = 0; i < itemsControl.Items.Count; i++)
                {
                    var childContainer = itemsControl.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                    if (childContainer == null)
                    {
                        // 虚拟化未生成的话，先 UpdateLayout 生成看看
                        itemsControl.UpdateLayout();
                        childContainer = itemsControl.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                    }

                    if (childContainer != null)
                    {
                        // 检查此子容器内部是否可能包含目标项 (即该节点在它的 Children 列表里)
                        if (childContainer.DataContext is TreeViewNode parentNode && ContainsNode(parentNode, targetNode))
                        {
                            // 发现目标节点在这个分支下，展开它
                            if (!childContainer.IsExpanded)
                            {
                                childContainer.IsExpanded = true;
                                childContainer.UpdateLayout(); // 展开后强制布局以生成内部节点
                            }

                            // 递归查找
                            var result = FindAndExpandTreeViewItem(childContainer, targetNode);
                            if (result != null) return result;
                        }
                    }
                }
            }
            catch (InvalidOperationException)
            {
                // 容器正在生成中（"无法在正在进行内容生成时调用 StartAt"），安全忽略
                Debug.WriteLine("[FindAndExpandTreeViewItem] 被跳过：容器正在生成中");
            }

            return null;
        }

        private bool ContainsNode(TreeViewNode parentNode, TreeViewNode targetNode)
        {
            if (parentNode == null || parentNode.Children == null) return false;
            foreach (var child in parentNode.Children)
            {
                if (child == targetNode || ContainsNode(child, targetNode)) return true;
            }
            return false;
        }

        private void SyncTreeViewSelection(TreeViewNode targetNode)
        {
            if (targetNode == null) return;

            try
            {
                // 1. 数据层：确保只有目标节点被选中
                ClearAllTreeViewSelections();
                targetNode.IsSelected = true;

                // 2. UI层：确保目标节点可见，放在空闲时执行避免渲染树冲突
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    try
                    {
                        StoreTreeView.UpdateLayout();
                        ScrollToNode(targetNode); 
                    }
                    catch (Exception iex)
                    {
                        Debug.WriteLine($"UI 更新失败: {iex.Message}");
                    }
                }, System.Windows.Threading.DispatcherPriority.ContextIdle);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"同步选中状态失败: {ex.Message}");
            }
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
        private string _text = string.Empty;
        private string _groupName;
        private string _source;
        private bool _isSelected; // ✅ 新增：选中状态字段

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

        public string Text
        {
            get => _text;
            set
            {
                if (_text != value)
                {
                    _text = value;
                    OnPropertyChanged(nameof(Text));
                }
            }
        }
        public string RawData { get; set; }  // ✅ 新增：保存原始完整数据，用于复制
        public string StoreName { get; set; }
        public SendStrategy Strategy { get; set; } = SendStrategy.TextDirect;
        
        // ✅ 新增：IsSelected 属性 (支持双向绑定)
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

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

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }







    #endregion
}
