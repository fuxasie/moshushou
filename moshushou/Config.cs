using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace moshushou
{
    public static class FileParseModes
    {
        public const string Auto = "Auto";
        public const string Magician = "Magician";
        public const string Issue = "Issue";

        public static string Normalize(string? mode)
        {
            if (string.IsNullOrWhiteSpace(mode))
            {
                return Auto;
            }

            string value = mode.Trim();
            if (value.Equals(Magician, StringComparison.OrdinalIgnoreCase) ||
                value.Equals("Normal2", StringComparison.OrdinalIgnoreCase))
            {
                return Magician;
            }

            if (value.Equals(Issue, StringComparison.OrdinalIgnoreCase) ||
                value.Equals("Table", StringComparison.OrdinalIgnoreCase))
            {
                return Issue;
            }

            return Auto;
        }
    }

    /// <summary>
    /// 文件状态类：保存单个文件的完整操作状态
    /// </summary>
    public class SegmentFailureState
    {
        public string StoreName { get; set; } = "";
        public int FailedSegment { get; set; }
        public int TotalSegments { get; set; }
        public int SentSegments { get; set; }
        public int SentItems { get; set; }
        public int TotalItems { get; set; }
        public string Reason { get; set; } = "发送失败";
    }

    public class FileState
    {
        public string FilePath { get; set; } = "";              // 文件完整路径
        public DateTime LastModifiedTime { get; set; }          // 文件最后修改时间
        public string LastSelectedStoreName { get; set; } = ""; // 上次选中的商家名
        public List<string> FailedStores { get; set; } = new(); // 自动重试区列表
        public List<string> ManualReviewStores { get; set; } = new(); // 需人工列表
        public List<SegmentFailureState> SegmentFailures { get; set; } = new(); // 分段发送失败进度
        public List<string> DeletedStores { get; set; } = new();     // 已删除的商家列表
        public bool IsIssueMode { get; set; } = false;               // 是否为问题件模式
        public bool IsCustomMessageMode { get; set; } = false;       // 是否为自定义话术模式(4列)
    }

    /// <summary>
    /// 单文件解析覆盖规则：用于手动指定当前文件的解析方式。
    /// </summary>
    public class FileParseOverride
    {
        public string FilePath { get; set; } = "";
        public string ParseMode { get; set; } = FileParseModes.Auto; // Auto | Magician | Issue
        public int TrackingColumn { get; set; } = 1;                 // 1-based
        public int StoreColumn { get; set; } = 2;                    // 1-based
        public int IssueSegmentStartCount { get; set; } = 30;        // 问题件：达到该条数后开始分段（>=2）
        public string TailMessage { get; set; } = "";                // 魔术师格式尾部话术
    }

    public class SearchConfig
    {

        // 新增：智能激活的超时时间（毫秒）
        public int ActivationTimeoutMs { get; set; } = 2000; // 1秒内没激活就算失败

        // ✅ 新增：微信 4.0 适配 (进程名)
        public string WeChatProcessName { get; set; } = "Weixin";

        // 优化后的延迟，可以适当缩短
        public int DelayAfterCtrlF { get; set; } = 100;
        public int DelayKeyboardAction { get; set; } = 30;

        // 这两个已经不再直接使用，但保留以防万一
        public int DelayWindowActivate { get; set; } = 100;
        public int DelayClipboard { get; set; } = 10;

        public string WechatWindowClassName { get; set; } = "WeChatMainWndForPC"; // 3.0 旧版类名 (保留备用)
        public string WeworkWindowClassName { get; set; } = "WeWorkWindow";

        // 🚫 以下属性已弃用，系统现在完全基于 YOLO 动态识别
        [Obsolete("Use YOLO detection instead")]
        public int WeChatCropLeft { get; set; } = 270;
        [Obsolete("Use YOLO detection instead")]
        public int WeChatCropHeight { get; set; } = 53;
        [Obsolete("Use YOLO detection instead")]
        public int WeChatRightCrop { get; set; } = 125;
        [Obsolete("Use YOLO detection instead")]
        public int WeWorkRightCrop { get; set; } = 100;

        [Obsolete("Use YOLO detection instead")]
        public int[] WeChatSearchResultRect { get; set; } = new int[] { 0, 60, 350, 400 }; 
        [Obsolete("Use YOLO detection instead")]
        public int[] WeWorkSearchResultRect { get; set; } = new int[] { 78, 90, 394, 58 };

        // ✅ 新增：记住上次打开的文件和选中位置 (向后兼容，加载时会迁移到 LastFileState)
        public string LastOpenedFilePath { get; set; } = "";
        public string LastSelectedStoreName { get; set; } = "";

        // ✅ 新增：文件状态持久化（包含完整操作状态）
        public FileState LastFileState { get; set; } = new FileState();

        // ✅ 新增：问题件模式的独立状态存储
        public FileState LastIssueFileState { get; set; } = new FileState();

        // ✅ 新增：自定义话术模式(4列)的独立状态存储
        public FileState LastCustomMessageFileState { get; set; } = new FileState();

        // ✅ 新增：按文件路径保存解析覆盖规则（调试历史中手动配置）
        public List<FileParseOverride> FileParseOverrides { get; set; } = new();

        // ✅ 新增：固定话术（可配置）
        public string FixedMessage { get; set; } = "现同步未发货预警，超时未交件会考核处罚，请尽快处理转出,已售后的及时发起拦截。（注：未处理售后请勿虚假拦截，核实虚假正常考核处罚。字节超时未发出总部将发起拦截）";

        /// <summary>
        /// 分段发送：每段最大条数（4列/5列模式超过此数量时分段发送）
        /// </summary>
        public int SegmentSize { get; set; } = 30;

        /// <summary>
        /// 分段发送：段间延迟毫秒数
        /// </summary>
        public int SegmentDelayMs { get; set; } = 500;

        // ✅ 新增：OSD 悬浮窗显示开关
        public bool EnableOsdWindow { get; set; } = true;
        
        // ✅ 新增：按 Ctrl+空格 发送后是否禁止自动跳下一项
        public bool SkipNextOnCtrlSpace { get; set; } = false;

        // ... Load 和 Save 方法保持不变 ...
        private static readonly string ConfigPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "search_config.json");
        public static SearchConfig Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    var loadedConfig = JsonSerializer.Deserialize<SearchConfig>(json);
                    if (loadedConfig != null)
                    {
                        bool changed = false;

                        // 企业微信文本限制：分段条数上限统一收敛为 30
                        if (loadedConfig.SegmentSize <= 0 || loadedConfig.SegmentSize > 30)
                        {
                            loadedConfig.SegmentSize = 30;
                            changed = true;
                        }

                        loadedConfig.FileParseOverrides ??= new List<FileParseOverride>();
                        int beforeCount = loadedConfig.FileParseOverrides.Count;
                        loadedConfig.FileParseOverrides = loadedConfig.FileParseOverrides
                            .Where(item => item != null && !string.IsNullOrWhiteSpace(item.FilePath))
                            .Select(item =>
                            {
                                item.FilePath = item.FilePath.Trim();
                                item.ParseMode = FileParseModes.Normalize(item.ParseMode);
                                if (item.TrackingColumn <= 0) item.TrackingColumn = 1;
                                if (item.StoreColumn <= 0) item.StoreColumn = 2;
                                if (item.IssueSegmentStartCount < 2)
                                {
                                    item.IssueSegmentStartCount = Math.Max(2, loadedConfig.SegmentSize);
                                }
                                item.TailMessage = item.TailMessage?.Trim() ?? string.Empty;
                                return item;
                            })
                            .ToList();

                        // 同一文件可能因历史路径格式差异产生多条规则，按规范化路径去重并保留最后一条。
                        loadedConfig.FileParseOverrides = loadedConfig.FileParseOverrides
                            .GroupBy(item => NormalizePathKey(item.FilePath), StringComparer.OrdinalIgnoreCase)
                            .Select(group => group.Last())
                            .ToList();

                        if (loadedConfig.FileParseOverrides.Count != beforeCount)
                        {
                            changed = true;
                        }

                        if (changed)
                        {
                            loadedConfig.Save();
                        }

                        return loadedConfig;
                    }
                }
            }
            catch { }
            var config = new SearchConfig();
            config.Save();
            return config;
        }

        public void Save()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping // 支持中文直接显示
                };
                string json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(ConfigPath, json);
            }
            catch { }
        }

        private static string NormalizePathKey(string path)
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
    }
}
