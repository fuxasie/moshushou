using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace moshushou
{
    public partial class DebugLogWindow : Window
    {
        private const int MaxLines = 20000;
        private readonly ObservableCollection<string> _logs = new ObservableCollection<string>();
        private readonly List<string> _allLogs = new List<string>();
        private string _filterKeyword = string.Empty;

        public DebugLogWindow()
        {
            InitializeComponent();
            LogListBox.ItemsSource = _logs;

            Loaded += DebugLogWindow_Loaded;
            Closed += DebugLogWindow_Closed;
        }

        private void DebugLogWindow_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshSnapshot();
            DebugLogManager.LogAdded += DebugLogManager_LogAdded;
            DebugLogManager.LogsCleared += DebugLogManager_LogsCleared;
        }

        private void DebugLogWindow_Closed(object? sender, EventArgs e)
        {
            DebugLogManager.LogAdded -= DebugLogManager_LogAdded;
            DebugLogManager.LogsCleared -= DebugLogManager_LogsCleared;
        }

        private void DebugLogManager_LogAdded(string line)
        {
            Dispatcher.InvokeAsync(() =>
            {
                _allLogs.Add(line);
                if (_allLogs.Count > MaxLines)
                {
                    _allLogs.RemoveAt(0);
                }

                if (PassesFilters(line))
                {
                    _logs.Add(line);
                }

                UpdateCounter();
                ScrollToEndIfNeeded();
            });
        }

        private void DebugLogManager_LogsCleared()
        {
            Dispatcher.InvokeAsync(() =>
            {
                _allLogs.Clear();
                _logs.Clear();
                UpdateCounter();
            });
        }

        private void RefreshSnapshot()
        {
            _allLogs.Clear();
            _allLogs.AddRange(DebugLogManager.GetSnapshot());
            if (_allLogs.Count > MaxLines)
            {
                _allLogs.RemoveRange(0, _allLogs.Count - MaxLines);
            }

            ApplyFilters();
        }

        private void UpdateCounter()
        {
            CounterTextBlock.Text = $"{_logs.Count}/{_allLogs.Count} lines";
        }

        private void ScrollToEndIfNeeded()
        {
            if (AutoScrollCheckBox.IsChecked != true || _logs.Count == 0)
            {
                return;
            }

            var last = _logs.LastOrDefault();
            if (last != null)
            {
                LogListBox.ScrollIntoView(last);
            }
        }

        private void ClearLogButton_Click(object sender, RoutedEventArgs e)
        {
            DebugLogManager.Clear();
        }

        private void CopyAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (_logs.Count == 0)
            {
                return;
            }

            Clipboard.SetText(string.Join(Environment.NewLine, _logs));
        }

        private void ClearHistoryButton_Click(object sender, RoutedEventArgs e)
        {
            StoreSendHistoryRepository.Clear();
            DebugLogManager.Log("历史", "已清空持久化历史。");
        }

        private void FilterTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _filterKeyword = FilterTextBox.Text?.Trim() ?? string.Empty;
            ApplyFilters();
        }

        private void HistoryOnlyCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            _logs.Clear();
            foreach (var line in _allLogs)
            {
                if (PassesFilters(line))
                {
                    _logs.Add(line);
                }
            }

            UpdateCounter();
            ScrollToEndIfNeeded();
        }

        private bool PassesFilters(string line)
        {
            if (HistoryOnlyCheckBox.IsChecked == true &&
                !IsHistoryLine(line))
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(_filterKeyword) &&
                line.IndexOf(_filterKeyword, StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }

            return true;
        }

        private static bool IsHistoryLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            return line.Contains("[历史]", StringComparison.Ordinal) ||
                   line.Contains("[发送历史]", StringComparison.Ordinal) ||
                   line.Contains("[点击历史]", StringComparison.Ordinal);
        }

    }
}
