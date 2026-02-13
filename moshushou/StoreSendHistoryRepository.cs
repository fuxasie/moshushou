using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace moshushou
{
    public sealed class StoreSendHistoryEntry
    {
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string StoreName { get; set; } = string.Empty;
        public string Action { get; set; } = string.Empty;
        public bool IsSuccess { get; set; }
        public string Detail { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
    }

    public static class StoreSendHistoryRepository
    {
        private const int MaxEntries = 5000;
        private static readonly object SyncRoot = new object();
        private static readonly string HistoryPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "store_send_history.json");

        private static List<StoreSendHistoryEntry>? _entries;

        public static IReadOnlyList<StoreSendHistoryEntry> GetSnapshot()
        {
            lock (SyncRoot)
            {
                EnsureLoaded();
                return _entries!.ToList();
            }
        }

        public static IReadOnlyList<StoreSendHistoryEntry> GetRecent(int take)
        {
            if (take <= 0)
            {
                return Array.Empty<StoreSendHistoryEntry>();
            }

            lock (SyncRoot)
            {
                EnsureLoaded();

                int count = _entries!.Count;
                if (count <= take)
                {
                    return _entries.ToList();
                }

                return _entries.Skip(count - take).ToList();
            }
        }

        public static void Append(StoreSendHistoryEntry? entry)
        {
            if (entry == null || string.IsNullOrWhiteSpace(entry.StoreName))
            {
                return;
            }

            lock (SyncRoot)
            {
                EnsureLoaded();

                entry.StoreName = entry.StoreName.Trim();
                entry.Action = string.IsNullOrWhiteSpace(entry.Action) ? "Unknown" : entry.Action.Trim();
                entry.Detail = entry.Detail?.Trim() ?? string.Empty;
                entry.FilePath = entry.FilePath?.Trim() ?? string.Empty;
                if (entry.Timestamp == default)
                {
                    entry.Timestamp = DateTime.Now;
                }

                _entries!.Add(entry);
                if (_entries.Count > MaxEntries)
                {
                    _entries.RemoveRange(0, _entries.Count - MaxEntries);
                }

                SaveUnsafe();
            }
        }

        public static void Clear()
        {
            lock (SyncRoot)
            {
                _entries = new List<StoreSendHistoryEntry>();
                SaveUnsafe();
            }
        }

        private static void EnsureLoaded()
        {
            if (_entries != null)
            {
                return;
            }

            try
            {
                if (File.Exists(HistoryPath))
                {
                    string json = File.ReadAllText(HistoryPath, Encoding.UTF8);
                    _entries = JsonSerializer.Deserialize<List<StoreSendHistoryEntry>>(json) ?? new List<StoreSendHistoryEntry>();
                }
                else
                {
                    _entries = new List<StoreSendHistoryEntry>();
                }
            }
            catch
            {
                _entries = new List<StoreSendHistoryEntry>();
            }

            _entries = _entries
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.StoreName))
                .Select(item =>
                {
                    item.StoreName = item.StoreName.Trim();
                    item.Action = item.Action?.Trim() ?? string.Empty;
                    item.Detail = item.Detail?.Trim() ?? string.Empty;
                    item.FilePath = item.FilePath?.Trim() ?? string.Empty;
                    return item;
                })
                .ToList();
        }

        private static void SaveUnsafe()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                string json = JsonSerializer.Serialize(_entries, options);
                File.WriteAllText(HistoryPath, json, Encoding.UTF8);
            }
            catch
            {
            }
        }
    }
}
