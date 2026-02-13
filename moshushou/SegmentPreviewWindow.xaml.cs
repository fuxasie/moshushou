using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;

namespace moshushou
{
    public partial class SegmentPreviewWindow : Window
    {
        private readonly string? _allContent;
        private readonly bool _hasAllContent;
        private bool _isShowingAll;

        public SegmentPreviewWindow(string segmentTitle, string content, string? allContent = null)
        {
            InitializeComponent();

            string safeTitle = string.IsNullOrWhiteSpace(segmentTitle) ? "内容" : segmentTitle.Trim();
            string currentContent = content ?? string.Empty;
            _allContent = string.IsNullOrWhiteSpace(allContent) ? null : allContent;
            _hasAllContent = !string.IsNullOrWhiteSpace(_allContent);
            _isShowingAll = false;

            Title = $"内容预览 - {safeTitle}";
            ShowAllButton.Visibility = _hasAllContent ? Visibility.Visible : Visibility.Collapsed;
            UpdateDisplayedContent(currentContent);

            Loaded += (_, _) =>
            {
                ContentTextBox.Focus();
                ContentTextBox.Select(0, 0);
                ContentTextBox.ScrollToHome();
            };
        }

        private static int CountNonEmptyLines(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return 0;
            }

            return content
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
                .Count(line => !string.IsNullOrWhiteSpace(line));
        }

        private async void CopyAllButton_Click(object sender, RoutedEventArgs e)
        {
            bool copied = await TrySetClipboardTextAsync(ContentTextBox.Text ?? string.Empty);
            SummaryTextBlock.Text = copied
                ? $"已复制当前显示内容（{CountNonEmptyLines(ContentTextBox.Text ?? string.Empty)}行）"
                : "剪贴板忙，复制失败";
        }

        private async void CopyNumbersButton_Click(object sender, RoutedEventArgs e)
        {
            var numbers = ExtractTrackingNumbers(ContentTextBox.Text ?? string.Empty);
            if (numbers.Count == 0)
            {
                SummaryTextBlock.Text = "未识别到可复制的单号";
                return;
            }

            bool copied = await TrySetClipboardTextAsync(string.Join(Environment.NewLine, numbers));
            SummaryTextBlock.Text = copied
                ? $"已复制单号（{numbers.Count}个）"
                : "剪贴板忙，复制失败";
        }

        private void ShowAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_hasAllContent || string.IsNullOrWhiteSpace(_allContent))
            {
                SummaryTextBlock.Text = "没有可显示的全部内容";
                return;
            }

            _isShowingAll = true;
            UpdateDisplayedContent(_allContent);
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private static List<string> ExtractTrackingNumbers(string content)
        {
            var results = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(content))
            {
                return results;
            }

            var lines = content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            foreach (string rawLine in lines)
            {
                string line = rawLine?.Trim() ?? string.Empty;
                if (line.Length == 0)
                {
                    continue;
                }

                string firstCell = line;
                int tabIndex = firstCell.IndexOf('\t');
                if (tabIndex > 0)
                {
                    firstCell = firstCell[..tabIndex];
                }
                else
                {
                    int commaIndex = firstCell.IndexOf(',');
                    if (commaIndex > 0)
                    {
                        firstCell = firstCell[..commaIndex];
                    }

                    int spaceIndex = firstCell.IndexOfAny(new[] { ' ', '\u3000' });
                    if (spaceIndex > 0)
                    {
                        firstCell = firstCell[..spaceIndex];
                    }
                }

                string candidate = firstCell.Trim().Trim('"', '\'', '，', ',', ';', '；');
                if (candidate.Length == 0)
                {
                    continue;
                }

                if (seen.Add(candidate))
                {
                    results.Add(candidate);
                }
            }

            return results;
        }

        private void UpdateDisplayedContent(string content)
        {
            string safeContent = content ?? string.Empty;
            ContentTextBox.Text = safeContent;
            int lineCount = CountNonEmptyLines(safeContent);
            string tag = _isShowingAll ? "全部" : "当前";
            SummaryTextBlock.Text = $"{tag}内容：{lineCount}行，{safeContent.Length}字";
        }

        private static async Task<bool> TrySetClipboardTextAsync(string text, int maxAttempts = 12, int baseDelayMs = 25)
        {
            string safeText = text ?? string.Empty;

            for (int i = 1; i <= maxAttempts; i++)
            {
                try
                {
                    Clipboard.SetText(safeText);
                    return true;
                }
                catch (COMException) when (i < maxAttempts)
                {
                    await Task.Delay(baseDelayMs + i * 15);
                }
                catch (ExternalException) when (i < maxAttempts)
                {
                    await Task.Delay(baseDelayMs + i * 15);
                }
                catch (InvalidOperationException) when (i < maxAttempts)
                {
                    await Task.Delay(baseDelayMs + i * 15);
                }
            }

            return false;
        }
    }
}
