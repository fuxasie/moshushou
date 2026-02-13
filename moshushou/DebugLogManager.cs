using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;

namespace moshushou
{
    public static class DebugLogManager
    {
        private const int MaxHistory = 8000;
        private static readonly object SyncRoot = new object();
        private static readonly List<string> History = new List<string>(MaxHistory);
        private static bool _initialized;
        private static DebugCaptureTraceListener? _listener;

        public static event Action<string>? LogAdded;
        public static event Action? LogsCleared;

        public static void Initialize()
        {
            lock (SyncRoot)
            {
                if (_initialized)
                {
                    return;
                }

                _listener = new DebugCaptureTraceListener();
                Trace.Listeners.Add(_listener);
                Trace.AutoFlush = true;
                _initialized = true;
            }

            Append("System", "Debug log capture initialized");
        }

        public static void Log(string source, string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            string safeSource = string.IsNullOrWhiteSpace(source) ? "Log" : source.Trim();
            Append(safeSource, message.Trim());
        }

        public static IReadOnlyList<string> GetSnapshot()
        {
            lock (SyncRoot)
            {
                return History.ToList();
            }
        }

        public static void Clear()
        {
            lock (SyncRoot)
            {
                History.Clear();
            }

            LogsCleared?.Invoke();
        }

        internal static void AppendFromTrace(string? message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            Append("Debug", message.Trim());
        }

        private static void Append(string source, string message)
        {
            string line = $"[{DateTime.Now:HH:mm:ss.fff}] [{source}] {message}";

            lock (SyncRoot)
            {
                History.Add(line);
                if (History.Count > MaxHistory)
                {
                    History.RemoveRange(0, History.Count - MaxHistory);
                }
            }

            LogAdded?.Invoke(line);
        }
    }

    internal sealed class DebugCaptureTraceListener : TraceListener
    {
        private readonly ThreadLocal<StringBuilder> _buffer = new ThreadLocal<StringBuilder>(() => new StringBuilder());

        public override void Write(string? message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            _buffer.Value?.Append(message);
        }

        public override void WriteLine(string? message)
        {
            var sb = _buffer.Value;
            if (sb == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(message))
            {
                sb.Append(message);
            }

            string final = sb.ToString();
            sb.Clear();

            if (!string.IsNullOrWhiteSpace(final))
            {
                DebugLogManager.AppendFromTrace(final);
            }
        }
    }
}
