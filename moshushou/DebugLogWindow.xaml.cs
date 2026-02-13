using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace moshushou
{
    public partial class DebugLogWindow : Window
    {
        private readonly ObservableCollection<string> _logs = new ObservableCollection<string>();

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
                _logs.Add(line);
                if (_logs.Count > 12000)
                {
                    _logs.RemoveAt(0);
                }

                UpdateCounter();
                ScrollToEndIfNeeded();
            });
        }

        private void DebugLogManager_LogsCleared()
        {
            Dispatcher.InvokeAsync(() =>
            {
                _logs.Clear();
                UpdateCounter();
            });
        }

        private void RefreshSnapshot()
        {
            _logs.Clear();
            foreach (var line in DebugLogManager.GetSnapshot())
            {
                _logs.Add(line);
            }

            UpdateCounter();
            ScrollToEndIfNeeded();
        }

        private void UpdateCounter()
        {
            CounterTextBlock.Text = $"{_logs.Count} lines";
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
    }
}
