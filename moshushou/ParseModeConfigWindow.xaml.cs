using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace moshushou
{
    public partial class ParseModeConfigWindow : Window
    {
        private sealed class ParseModeOption
        {
            public string Mode { get; init; } = FileParseModes.Auto;
            public string Label { get; init; } = string.Empty;
            public override string ToString() => Label;
        }

        private readonly MainWindow.ParseOverrideDebugContext _context;
        private readonly List<ParseModeOption> _options;

        public FileParseOverride? ParseOverride { get; private set; }

        public ParseModeConfigWindow(MainWindow.ParseOverrideDebugContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _options = new List<ParseModeOption>
            {
                new ParseModeOption { Mode = FileParseModes.Auto, Label = "自动识别（恢复默认）" },
                new ParseModeOption { Mode = FileParseModes.Magician, Label = "魔术师格式（两列表格）" },
                new ParseModeOption { Mode = FileParseModes.Issue, Label = "问题件格式（整行Tab发送）" }
            };

            InitializeComponent();
            FilePathTextBlock.Text = _context.FilePath;

            ModeComboBox.ItemsSource = _options;
            string selectedMode = FileParseModes.Normalize(_context.ParseMode);
            ModeComboBox.SelectedItem = _options.FirstOrDefault(item => item.Mode == selectedMode) ?? _options[0];

            TrackingColumnTextBox.Text = Math.Max(1, _context.TrackingColumn).ToString();
            StoreColumnTextBox.Text = Math.Max(1, _context.StoreColumn).ToString();
            IssueSegmentStartCountTextBox.Text = Math.Max(2, _context.IssueSegmentStartCount).ToString();
            TailMessageTextBox.Text = _context.TailMessage ?? string.Empty;

            UpdateInputState();
        }

        private string GetSelectedMode()
        {
            if (ModeComboBox.SelectedItem is ParseModeOption option)
            {
                return option.Mode;
            }

            return FileParseModes.Auto;
        }

        private void UpdateInputState()
        {
            string mode = GetSelectedMode();
            bool isIssueMode = string.Equals(mode, FileParseModes.Issue, StringComparison.Ordinal);
            bool isMagicianMode = string.Equals(mode, FileParseModes.Magician, StringComparison.Ordinal);

            IssueConfigPanel.Visibility = isIssueMode ? Visibility.Visible : Visibility.Collapsed;
            MagicianConfigPanel.Visibility = isMagicianMode ? Visibility.Visible : Visibility.Collapsed;

            if (isIssueMode)
            {
                if (_context.DetectedColumnCount > 0)
                {
                    HintTextBlock.Text = $"问题件格式：将按整行内容发送（Tab分割）。请填写运单号列、商家名列和分段起始条数（>=2）。当前检测列数：{_context.DetectedColumnCount}。";
                }
                else
                {
                    HintTextBlock.Text = "问题件格式：将按整行内容发送（Tab分割）。请填写运单号列、商家名列和分段起始条数（>=2）。";
                }
            }
            else if (isMagicianMode)
            {
                HintTextBlock.Text = "魔术师格式：发送顺序为 店铺名 -> 运单号列表 -> 尾部自定义话术。";
            }
            else
            {
                HintTextBlock.Text = "自动识别：按程序原有列数识别逻辑解析。";
            }
        }

        private void ModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateInputState();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            string parseMode = GetSelectedMode();
            int trackingColumn = Math.Max(1, _context.TrackingColumn);
            int storeColumn = Math.Max(1, _context.StoreColumn);
            int issueSegmentStartCount = Math.Max(2, _context.IssueSegmentStartCount);
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

            ParseOverride = new FileParseOverride
            {
                FilePath = _context.FilePath,
                ParseMode = parseMode,
                TrackingColumn = trackingColumn,
                StoreColumn = storeColumn,
                IssueSegmentStartCount = issueSegmentStartCount,
                TailMessage = tailMessage
            };

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
