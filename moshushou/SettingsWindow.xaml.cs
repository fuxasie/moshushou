using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace moshushou
{
    public partial class SettingsWindow : Window
    {
        private sealed class ParseModeOption
        {
            public string Mode { get; init; } = FileParseModes.Auto;
            public string Label { get; init; } = string.Empty;
            public override string ToString() => Label;
        }

        private readonly SearchConfig _config;
        private readonly MainWindow _mainWindow;
        private readonly List<ParseModeOption> _parseModeOptions;
        private MainWindow.ParseOverrideDebugContext? _parseContext;

        public SettingsWindow(SearchConfig config, MainWindow mainWindow)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _mainWindow = mainWindow ?? throw new ArgumentNullException(nameof(mainWindow));

            _parseModeOptions = new List<ParseModeOption>
            {
                new ParseModeOption { Mode = FileParseModes.Auto, Label = "自动识别（恢复默认）" },
                new ParseModeOption { Mode = FileParseModes.Magician, Label = "魔术师格式（两列表格）" },
                new ParseModeOption { Mode = FileParseModes.Issue, Label = "问题件格式（整行 Tab 发送）" }
            };

            InitializeComponent();

            // 初始化基础控件
            EnableOsdCheckBox.IsChecked = _config.EnableOsdWindow;
            SkipNextOnCtrlSpaceCheckBox.IsChecked = _config.SkipNextOnCtrlSpace;

            // 初始化解析模式下拉
            ModeComboBox.ItemsSource = _parseModeOptions;

            Loaded += SettingsWindow_Loaded;
        }

        private void SettingsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 尝试获取当前文件解析上下文
            _parseContext = _mainWindow.GetCurrentFileParseContext();
            RefreshParseSection();
        }

        private void RefreshParseSection()
        {
            if (_parseContext == null)
            {
                CurrentFileTextBlock.Text = "（未加载文件，请先在主窗口加载 Excel）";
                ApplyParseButton.IsEnabled = false;
                ModeComboBox.IsEnabled = false;
                return;
            }

            CurrentFileTextBlock.Text = _parseContext.FilePath;
            ModeComboBox.IsEnabled = true;
            ApplyParseButton.IsEnabled = true;

            string selectedMode = FileParseModes.Normalize(_parseContext.ParseMode);
            ModeComboBox.SelectedItem = _parseModeOptions.FirstOrDefault(o => o.Mode == selectedMode) ?? _parseModeOptions[0];

            TrackingColumnTextBox.Text = Math.Max(1, _parseContext.TrackingColumn).ToString();
            StoreColumnTextBox.Text = Math.Max(1, _parseContext.StoreColumn).ToString();
            IssueSegmentStartCountTextBox.Text = Math.Max(2, _parseContext.IssueSegmentStartCount).ToString();
            TailMessageTextBox.Text = _parseContext.TailMessage ?? string.Empty;

            UpdateParseModeInputState();
        }

        private string GetSelectedParseMode()
        {
            if (ModeComboBox.SelectedItem is ParseModeOption opt)
            {
                return opt.Mode;
            }
            return FileParseModes.Auto;
        }

        private void UpdateParseModeInputState()
        {
            string mode = GetSelectedParseMode();
            bool isIssue = string.Equals(mode, FileParseModes.Issue, StringComparison.Ordinal);
            bool isMagician = string.Equals(mode, FileParseModes.Magician, StringComparison.Ordinal);

            IssueConfigPanel.Visibility = isIssue ? Visibility.Visible : Visibility.Collapsed;
            MagicianConfigPanel.Visibility = isMagician ? Visibility.Visible : Visibility.Collapsed;

            if (isIssue)
            {
                string detectedCols = _parseContext?.DetectedColumnCount > 0
                    ? $"（检测到 {_parseContext.DetectedColumnCount} 列）"
                    : string.Empty;
                ParseHintTextBlock.Text = $"按整行内容（Tab 分割）发送，需指定运单号列和商家名列{detectedCols}。";
            }
            else if (isMagician)
            {
                ParseHintTextBlock.Text = "发送顺序：店铺名 → 运单号列表 → 尾部自定义话术。";
            }
            else
            {
                ParseHintTextBlock.Text = "按程序原有列数识别逻辑自动解析，无需额外配置。";
            }
        }

        private void ModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateParseModeInputState();
        }

        // ====== OSD 事件 ======

        private void EnableOsdCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            _config.EnableOsdWindow = true;
            _config.Save();
        }

        private void EnableOsdCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            _config.EnableOsdWindow = false;
            _config.Save();
        }

        private void OsdSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            OsdWindow.ToggleEditMode();
        }

        // ====== 快捷键事件 ======

        private void SkipNextOnCtrlSpaceCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            _config.SkipNextOnCtrlSpace = true;
            _config.Save();
        }

        private void SkipNextOnCtrlSpaceCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            _config.SkipNextOnCtrlSpace = false;
            _config.Save();
        }

        // ====== 解析设置事件 ======

        private async void ApplyParseButton_Click(object sender, RoutedEventArgs e)
        {
            if (_parseContext == null)
            {
                MessageBox.Show(this, "请先在主窗口加载 Excel 文件，再设置解析方式。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string parseMode = GetSelectedParseMode();
            int trackingColumn = Math.Max(1, _parseContext.TrackingColumn);
            int storeColumn = Math.Max(1, _parseContext.StoreColumn);
            int issueSegmentStartCount = Math.Max(2, _parseContext.IssueSegmentStartCount);
            string tailMessage = TailMessageTextBox.Text?.Trim() ?? string.Empty;

            if (string.Equals(parseMode, FileParseModes.Issue, StringComparison.Ordinal))
            {
                if (!int.TryParse(TrackingColumnTextBox.Text?.Trim(), out trackingColumn) || trackingColumn <= 0)
                {
                    MessageBox.Show(this, "运单号列必须是大于 0 的整数。", "输入错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (!int.TryParse(StoreColumnTextBox.Text?.Trim(), out storeColumn) || storeColumn <= 0)
                {
                    MessageBox.Show(this, "商家名列必须是大于 0 的整数。", "输入错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (!int.TryParse(IssueSegmentStartCountTextBox.Text?.Trim(), out issueSegmentStartCount) || issueSegmentStartCount < 2)
                {
                    MessageBox.Show(this, "分段起始条数必须是大于等于 2 的整数。", "输入错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            else if (string.Equals(parseMode, FileParseModes.Magician, StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(tailMessage))
                {
                    MessageBox.Show(this, "魔术师格式必须填写尾部自定义话术。", "输入错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            var parseOverride = new FileParseOverride
            {
                FilePath = _parseContext.FilePath,
                ParseMode = parseMode,
                TrackingColumn = trackingColumn,
                StoreColumn = storeColumn,
                IssueSegmentStartCount = issueSegmentStartCount,
                TailMessage = tailMessage
            };

            ApplyParseButton.IsEnabled = false;
            try
            {
                await _mainWindow.ApplyCurrentFileParseOverrideAsync(parseOverride);
                // 重新获取上下文，刷新界面
                _parseContext = _mainWindow.GetCurrentFileParseContext();
                RefreshParseSection();
            }
            finally
            {
                ApplyParseButton.IsEnabled = true;
            }
        }

        // ====== 关闭 ======

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
