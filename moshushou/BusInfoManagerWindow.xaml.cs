using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace moshushou
{
    public partial class BusInfoManagerWindow : Window
    {
        private sealed class BusInfoRow : INotifyPropertyChanged
        {
            private string _storeName = string.Empty;
            private string _groupName = string.Empty;
            private string _source = string.Empty;

            public string StoreName
            {
                get => _storeName;
                set
                {
                    string next = value?.Trim() ?? string.Empty;
                    if (!string.Equals(_storeName, next, StringComparison.Ordinal))
                    {
                        _storeName = next;
                        OnPropertyChanged(nameof(StoreName));
                    }
                }
            }

            public string GroupName
            {
                get => _groupName;
                set
                {
                    string next = value?.Trim() ?? string.Empty;
                    if (!string.Equals(_groupName, next, StringComparison.Ordinal))
                    {
                        _groupName = next;
                        OnPropertyChanged(nameof(GroupName));
                    }
                }
            }

            public string Source
            {
                get => _source;
                set
                {
                    string next = NormalizeSource(value);
                    if (!string.Equals(_source, next, StringComparison.Ordinal))
                    {
                        _source = next;
                        OnPropertyChanged(nameof(Source));
                    }
                }
            }

            public event PropertyChangedEventHandler? PropertyChanged;

            private void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private readonly ObservableCollection<BusInfoRow> _rows = new ObservableCollection<BusInfoRow>();
        private readonly ICollectionView _rowsView;
        private string _filterKeyword = string.Empty;

        public List<BusinessInfo>? UpdatedBusinessInfos { get; private set; }
        public event Action<List<BusinessInfo>>? Saved;

        public BusInfoManagerWindow(IEnumerable<BusinessInfo>? sourceItems)
        {
            InitializeComponent();

            if (sourceItems != null)
            {
                foreach (BusinessInfo item in sourceItems)
                {
                    string storeName = NormalizeStoreName(item?.StoreName);
                    if (string.IsNullOrWhiteSpace(storeName))
                    {
                        continue;
                    }

                    _rows.Add(new BusInfoRow
                    {
                        StoreName = storeName,
                        GroupName = item?.GroupName?.Trim() ?? string.Empty,
                        Source = NormalizeSource(item?.Source)
                    });
                }
            }

            _rowsView = CollectionViewSource.GetDefaultView(_rows);
            _rowsView.Filter = FilterRow;

            MappingDataGrid.ItemsSource = _rowsView;
            RefreshSummary();
        }

        private static string NormalizeStoreName(string? storeName)
        {
            return (storeName ?? string.Empty).Trim();
        }

        private static string NormalizeSource(string? source)
        {
            string value = (source ?? string.Empty).Trim();
            if (string.Equals(value, "企业微信", StringComparison.Ordinal) ||
                string.Equals(value, "浼佷笟寰俊", StringComparison.Ordinal) ||
                string.Equals(value, "wxwork", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "wework", StringComparison.OrdinalIgnoreCase))
            {
                return "企业微信";
            }

            if (string.Equals(value, "微信", StringComparison.Ordinal) ||
                string.Equals(value, "寰俊", StringComparison.Ordinal) ||
                string.Equals(value, "wechat", StringComparison.OrdinalIgnoreCase))
            {
                return "微信";
            }

            return string.Empty;
        }

        private bool FilterRow(object obj)
        {
            if (obj is not BusInfoRow row)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(_filterKeyword))
            {
                return true;
            }

            return row.StoreName.Contains(_filterKeyword, StringComparison.OrdinalIgnoreCase) ||
                   row.GroupName.Contains(_filterKeyword, StringComparison.OrdinalIgnoreCase) ||
                   row.Source.Contains(_filterKeyword, StringComparison.OrdinalIgnoreCase);
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _filterKeyword = SearchTextBox.Text?.Trim() ?? string.Empty;
            _rowsView.Refresh();
            RefreshSummary();
        }

        private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
        {
            SearchTextBox.Text = string.Empty;
            SearchTextBox.Focus();
        }

        private void MappingDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MappingDataGrid.SelectedItem is not BusInfoRow row)
            {
                return;
            }

            EditorStoreNameTextBox.Text = row.StoreName;
            EditorGroupNameTextBox.Text = row.GroupName;
            SetSourceComboValue(row.Source);
        }

        private void SetSourceComboValue(string source)
        {
            string normalized = NormalizeSource(source);
            foreach (object item in EditorSourceComboBox.Items)
            {
                if (item is ComboBoxItem comboItem &&
                    string.Equals(comboItem.Content?.ToString() ?? string.Empty, normalized, StringComparison.Ordinal))
                {
                    EditorSourceComboBox.SelectedItem = comboItem;
                    return;
                }
            }

            EditorSourceComboBox.SelectedIndex = 0;
        }

        private string GetEditorSourceValue()
        {
            if (EditorSourceComboBox.SelectedItem is ComboBoxItem selected)
            {
                return NormalizeSource(selected.Content?.ToString());
            }

            return string.Empty;
        }

        private bool TryBuildEditorValues(out string storeName, out string groupName, out string source)
        {
            storeName = NormalizeStoreName(EditorStoreNameTextBox.Text);
            groupName = (EditorGroupNameTextBox.Text ?? string.Empty).Trim();
            source = GetEditorSourceValue();

            if (string.IsNullOrWhiteSpace(storeName))
            {
                HintTextBlock.Text = "提示：商家名不能为空。";
                return false;
            }

            if (string.IsNullOrWhiteSpace(groupName))
            {
                HintTextBlock.Text = "提示：群名不能为空。";
                return false;
            }

            return true;
        }

        private void AddOrMergeButton_Click(object sender, RoutedEventArgs e)
        {
            if (!TryBuildEditorValues(out string storeName, out string groupName, out string source))
            {
                return;
            }

            BusInfoRow? existing = _rows.FirstOrDefault(item =>
                string.Equals(item.StoreName, storeName, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.StoreName = storeName;
                existing.GroupName = groupName;
                existing.Source = source;
                HintTextBlock.Text = $"提示：已覆盖商家“{storeName}”的映射。";
                SelectAndFocusRow(existing);
            }
            else
            {
                var row = new BusInfoRow
                {
                    StoreName = storeName,
                    GroupName = groupName,
                    Source = source
                };
                _rows.Add(row);
                HintTextBlock.Text = $"提示：已新增商家“{storeName}”。";
                SelectAndFocusRow(row);
            }

            _rowsView.Refresh();
            RefreshSummary();
        }

        private void UpdateButton_Click(object sender, RoutedEventArgs e)
        {
            if (MappingDataGrid.SelectedItem is not BusInfoRow selected)
            {
                HintTextBlock.Text = "提示：请先在列表中选择一条记录。";
                return;
            }

            if (!TryBuildEditorValues(out string storeName, out string groupName, out string source))
            {
                return;
            }

            BusInfoRow? duplicate = _rows.FirstOrDefault(item =>
                !ReferenceEquals(item, selected) &&
                string.Equals(item.StoreName, storeName, StringComparison.OrdinalIgnoreCase));
            if (duplicate != null)
            {
                HintTextBlock.Text = $"提示：商家“{storeName}”已存在，请使用“新增/覆盖”或换名。";
                return;
            }

            selected.StoreName = storeName;
            selected.GroupName = groupName;
            selected.Source = source;

            _rowsView.Refresh();
            RefreshSummary();
            HintTextBlock.Text = $"提示：已修改商家“{storeName}”。";
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (MappingDataGrid.SelectedItem is not BusInfoRow selected)
            {
                HintTextBlock.Text = "提示：请先在列表中选择一条记录。";
                return;
            }

            string removedStore = selected.StoreName;
            _rows.Remove(selected);
            MappingDataGrid.SelectedItem = null;

            _rowsView.Refresh();
            RefreshSummary();
            HintTextBlock.Text = $"提示：已删除商家“{removedStore}”。";
        }

        private void ClearEditorButton_Click(object sender, RoutedEventArgs e)
        {
            EditorStoreNameTextBox.Text = string.Empty;
            EditorGroupNameTextBox.Text = string.Empty;
            EditorSourceComboBox.SelectedIndex = 0;
            MappingDataGrid.SelectedItem = null;
            HintTextBlock.Text = "提示：已清空输入。";
            EditorStoreNameTextBox.Focus();
        }

        private void SelectAndFocusRow(BusInfoRow row)
        {
            MappingDataGrid.SelectedItem = row;
            MappingDataGrid.ScrollIntoView(row);
        }

        private static string BuildShortHintStoreName(string storeName, int maxLength = 24)
        {
            string value = NormalizeStoreName(storeName);
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            if (value.Length <= maxLength)
            {
                return value;
            }

            if (maxLength <= 3)
            {
                return value.Substring(0, maxLength);
            }

            return value.Substring(0, maxLength - 3) + "...";
        }

        private void RefreshSummary()
        {
            int filtered = _rowsView.Cast<object>().Count();
            SummaryTextBlock.Text = $"显示 {filtered} / 共 {_rows.Count} 条";
        }

        public void SyncFromMainSelection(string storeName, string groupName, string source)
        {
            string normalizedStoreName = NormalizeStoreName(storeName);
            if (string.IsNullOrWhiteSpace(normalizedStoreName))
            {
                return;
            }

            string normalizedGroupName = (groupName ?? string.Empty).Trim();
            string normalizedSource = NormalizeSource(source);

            BusInfoRow? existing = _rows.FirstOrDefault(item =>
                string.Equals(item.StoreName, normalizedStoreName, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                EditorStoreNameTextBox.Text = existing.StoreName;
                EditorGroupNameTextBox.Text = existing.GroupName;
                SetSourceComboValue(existing.Source);
                SelectAndFocusRow(existing);
                HintTextBlock.Text = $"已联动：{BuildShortHintStoreName(normalizedStoreName)}";
                return;
            }

            MappingDataGrid.SelectedItem = null;
            EditorStoreNameTextBox.Text = normalizedStoreName;
            EditorGroupNameTextBox.Text = normalizedGroupName;
            SetSourceComboValue(normalizedSource);
            EditorStoreNameTextBox.Focus();
            EditorStoreNameTextBox.CaretIndex = EditorStoreNameTextBox.Text.Length;
            HintTextBlock.Text = $"待新增：{BuildShortHintStoreName(normalizedStoreName)}";
        }

        private bool TryCommitPendingEditorChangesOnSave()
        {
            bool hasStoreInput = !string.IsNullOrWhiteSpace(EditorStoreNameTextBox.Text);
            bool hasGroupInput = !string.IsNullOrWhiteSpace(EditorGroupNameTextBox.Text);
            bool hasSourceInput = !string.IsNullOrWhiteSpace(GetEditorSourceValue());
            bool hasAnyEditorInput = hasStoreInput || hasGroupInput || hasSourceInput;

            if (!hasAnyEditorInput)
            {
                return true;
            }

            if (!TryBuildEditorValues(out string storeName, out string groupName, out string source))
            {
                return false;
            }

            if (MappingDataGrid.SelectedItem is BusInfoRow selected)
            {
                BusInfoRow? duplicate = _rows.FirstOrDefault(item =>
                    !ReferenceEquals(item, selected) &&
                    string.Equals(item.StoreName, storeName, StringComparison.OrdinalIgnoreCase));
                if (duplicate != null)
                {
                    HintTextBlock.Text = $"提示：商家“{storeName}”已存在，请先处理重名项。";
                    return false;
                }

                selected.StoreName = storeName;
                selected.GroupName = groupName;
                selected.Source = source;
                _rowsView.Refresh();
                SelectAndFocusRow(selected);
                RefreshSummary();
                return true;
            }

            BusInfoRow? existing = _rows.FirstOrDefault(item =>
                string.Equals(item.StoreName, storeName, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.StoreName = storeName;
                existing.GroupName = groupName;
                existing.Source = source;
                _rowsView.Refresh();
                SelectAndFocusRow(existing);
                RefreshSummary();
                return true;
            }

            var row = new BusInfoRow
            {
                StoreName = storeName,
                GroupName = groupName,
                Source = source
            };
            _rows.Add(row);
            _rowsView.Refresh();
            SelectAndFocusRow(row);
            RefreshSummary();
            return true;
        }

        private void SaveAndCloseButton_Click(object sender, RoutedEventArgs e)
        {
            if (!TryCommitPendingEditorChangesOnSave())
            {
                return;
            }

            UpdatedBusinessInfos = _rows
                .Where(item =>
                    !string.IsNullOrWhiteSpace(item.StoreName) &&
                    !string.IsNullOrWhiteSpace(item.GroupName))
                .GroupBy(item => item.StoreName.Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(group => group.Last())
                .OrderBy(item => item.StoreName, StringComparer.OrdinalIgnoreCase)
                .Select(item => new BusinessInfo
                {
                    StoreName = item.StoreName.Trim(),
                    GroupName = item.GroupName.Trim(),
                    Source = NormalizeSource(item.Source)
                })
                .ToList();

            Saved?.Invoke(UpdatedBusinessInfos);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}



