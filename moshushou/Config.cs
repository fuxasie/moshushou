using System;
using System.IO;
using System.Text.Json;

namespace moshushou
{
    public class SearchConfig
    {
        // 新增：智能激活的超时时间（毫秒）
        public int ActivationTimeoutMs { get; set; } = 1000; // 1秒内没激活就算失败

        // 优化后的延迟，可以适当缩短
        public int DelayAfterCtrlF { get; set; } = 100;
        public int DelayKeyboardAction { get; set; } = 30;

        // 这两个已经不再直接使用，但保留以防万一
        public int DelayWindowActivate { get; set; } = 100;
        public int DelayClipboard { get; set; } = 10;

        public string WechatWindowClassName { get; set; } = "WeChatMainWndForPC";
        public string WeworkWindowClassName { get; set; } = "WeWorkWindow";

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