using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.Windows.Controls.Primitives;
using System.Text.Json;
using System.IO;

namespace moshushou
{
    public partial class OsdWindow : Window
    {
        private static OsdWindow _instance;
        private DispatcherTimer _timer;
        private bool _isEditMode = false;
        private int _bgOpacityPercent = 60; // 背景透明度百分比（0=全透明，100=不透明）
        private const string ConfigFilePath = "osd_config.json";

        // Win32 常量用于点击穿透
        private const int WS_EX_TRANSPARENT = 0x00000020;
        private const int GWL_EXSTYLE = -20;

        public static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex)
        {
            if (IntPtr.Size == 8) return GetWindowLongPtr64(hWnd, nIndex);
            return new IntPtr(GetWindowLong32(hWnd, nIndex));
        }

        public static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 8) return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
            return new IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32()));
        }

        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        public OsdWindow()
        {
            InitializeComponent();
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(1.5); // 停留1.5秒
            _timer.Tick += Timer_Tick;

            // 监听右键按下，关闭编辑模式
            this.MouseRightButtonDown += OsdWindow_MouseRightButtonDown;
            LoadConfig();
            ApplyBgOpacity(_bgOpacityPercent);
        }

        private void OsdWindow_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_isEditMode)
            {
                SetEditMode(false);
            }
        }

        private void ExitEditModeButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isEditMode)
            {
                SetEditMode(false);
            }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            // 初始时设置为穿透模式
            SetWindowExTransparent(true);
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            // 拦截关闭事件，防止 WPF 把窗口彻底回收，下次才能安全重新 Show()
            e.Cancel = true;
            this.Hide();
        }

        private void SetWindowExTransparent(bool isTransparent)
        {
            IntPtr hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd == IntPtr.Zero) return;

            long extendedStyle = GetWindowLongPtr(hwnd, GWL_EXSTYLE).ToInt64();
            if (isTransparent)
            {
                extendedStyle |= WS_EX_TRANSPARENT;
            }
            else
            {
                extendedStyle &= ~WS_EX_TRANSPARENT;
            }
            SetWindowLongPtr(hwnd, GWL_EXSTYLE, new IntPtr(extendedStyle));
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (double.IsNaN(this.Left) || double.IsNaN(this.Top))
            {
                CenterWindowOnScreen();
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            _timer.Stop();
            if (!_isEditMode)
            {
                this.Hide();
            }
        }

        private void CenterWindowOnScreen()
        {
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            double screenHeight = SystemParameters.PrimaryScreenHeight;
            this.Left = (screenWidth - this.ActualWidth) / 2;
            this.Top = screenHeight - this.ActualHeight - 200; // 悬浮在屏幕偏下方
        }

        public static void ShowMessage(string message, string sequenceInfo = "")
        {
            // 使用 InvokeAsync 避免在 VirtualizingStackPanel 布局过程中同步重入导致 InvalidOperationException
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    EnsureInstance();

                    _instance.SequenceRun.Text = sequenceInfo;
                    _instance.MessageRun.Text = message;
                    
                    _instance.Opacity = 1; // 确保可见
                    _instance.Show();

                    if (!_instance._isEditMode)
                    {
                        _instance._timer.Stop();
                        _instance._timer.Start();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"OSD 显示异常: {ex.Message}");
                }
            });
        }

        public static void ToggleEditMode()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                EnsureInstance();
                _instance.SetEditMode(!_instance._isEditMode);
            });
        }

        private static void EnsureInstance()
        {
            if (_instance == null)
            {
                _instance = new OsdWindow();
            }
        }

        private void SetEditMode(bool enable)
        {
            _isEditMode = enable;
            _timer.Stop();

            if (enable)
            {
                this.Opacity = 1;
                this.Show();
                EditOverlay.Visibility = Visibility.Visible;
                ResizeThumb.Visibility = Visibility.Visible;
                // 同步 Slider 到当前透明度值
                OpacitySlider.Value = _bgOpacityPercent;
                SetWindowExTransparent(false); // 取消穿透，允许响应鼠标操作
            }
            else
            {
                EditOverlay.Visibility = Visibility.Collapsed;
                ResizeThumb.Visibility = Visibility.Collapsed;
                SetWindowExTransparent(true); // 恢复穿透模式
                SaveConfig();
                this.Hide();
            }
        }

        /// <summary>
        /// Slider 值改变时实时更新背景透明度
        /// </summary>
        private void OpacitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            int percent = (int)Math.Round(e.NewValue);
            _bgOpacityPercent = percent;
            ApplyBgOpacity(percent);
            if (OpacityValueText != null)
            {
                OpacityValueText.Text = $"{percent}%";
            }
        }

        /// <summary>
        /// 将百分比值应用到 BgBorder 的背景色 Alpha 通道
        /// </summary>
        private void ApplyBgOpacity(int percent)
        {
            if (BgBorder == null) return;
            byte alpha = (byte)(percent * 255 / 100);
            BgBorder.Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromArgb(alpha, 0, 0, 0));
        }

        private void EditOverlay_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_isEditMode)
            {
                this.DragMove();
            }
        }

        private void ResizeThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (_isEditMode)
            {
                double newWidth = this.Width + e.HorizontalChange;
                double newHeight = this.Height + e.VerticalChange;

                if (newWidth > 100) this.Width = newWidth;
                if (newHeight > 50) this.Height = newHeight;
            }
        }

        private void SaveConfig()
        {
            try
            {
                var config = new { Left = this.Left, Top = this.Top, Width = this.Width, Height = this.Height, BgOpacity = _bgOpacityPercent };
                string json = JsonSerializer.Serialize(config);
                File.WriteAllText(ConfigFilePath, json);
            }
            catch { }
        }

        private void LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    string json = File.ReadAllText(ConfigFilePath);
                    var config = JsonSerializer.Deserialize<JsonElement>(json);
                    
                    if (config.TryGetProperty("Left", out var left)) this.Left = left.GetDouble();
                    if (config.TryGetProperty("Top", out var top)) this.Top = top.GetDouble();
                    if (config.TryGetProperty("Width", out var width)) this.Width = width.GetDouble();
                    if (config.TryGetProperty("Height", out var height)) this.Height = height.GetDouble();
                    if (config.TryGetProperty("BgOpacity", out var opacity))
                    {
                        _bgOpacityPercent = Math.Clamp(opacity.GetInt32(), 0, 100);
                    }
                }
            }
            catch { }
        }
    }
}
