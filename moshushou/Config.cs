using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace moshushou
{
    /// <summary>
    /// 文件状态类：保存单个文件的完整操作状态
    /// </summary>
    public class FileState
    {
        public string FilePath { get; set; } = "";              // 文件完整路径
        public DateTime LastModifiedTime { get; set; }          // 文件最后修改时间
        public string LastSelectedStoreName { get; set; } = ""; // 上次选中的商家名
        public List<string> FailedStores { get; set; } = new(); // 自动重试区列表
        public List<string> ManualReviewStores { get; set; } = new(); // 需人工列表
        public List<string> DeletedStores { get; set; } = new();     // 已删除的商家列表
        public bool IsIssueMode { get; set; } = false;               // 是否为问题件模式
        public bool IsCustomMessageMode { get; set; } = false;       // 是否为自定义话术模式(4列)
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

        // ✅ 新增：固定话术（可配置）
        public string FixedMessage { get; set; } = "现同步未发货预警，超时未交件会考核处罚，请尽快处理转出,已售后的及时发起拦截。（注：未处理售后请勿虚假拦截，核实虚假正常考核处罚。字节超时未发出总部将发起拦截）";

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
                    return JsonSerializer.Deserialize<SearchConfig>(json);
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
    }
}
