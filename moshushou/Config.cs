using System;
using System.IO;
using System.Text.Json;

namespace moshushou
{
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

        // ✅ 新增：可配置的截图参数 (适配 Qt 版本)
        public int WeChatCropLeft { get; set; } = 270;
        public int WeChatCropHeight { get; set; } = 53;
        public int WeChatRightCrop { get; set; } = 125;
        public int WeWorkRightCrop { get; set; } = 100;

        // ✅ 新增：搜索结果列表坐标 (相对坐标 X, Y, W, H)
        // 微信 4.0 (Qt) 搜索下拉列表：X=0 起始，Y=60 (跳过搜索栏), W=350, H=400 (覆盖完整列表)
        public int[] WeChatSearchResultRect { get; set; } = new int[] { 0, 60, 350, 400 }; 
        public int[] WeWorkSearchResultRect { get; set; } = new int[] { 78, 90, 394, 58 };

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
                string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
            }
            catch { }
        }
    }
}