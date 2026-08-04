using System;

namespace moshushou.Input
{
    public static class InputBackendFactory
    {
        public const string VirtualHidMode = "VirtualHid";
        public const string SendInputMode = "SendInput";

        public static IInputBackend Create(SearchConfig config, Action<string>? log = null)
        {
            string requestedMode = string.IsNullOrWhiteSpace(config.InputBackend)
                ? VirtualHidMode
                : config.InputBackend.Trim();

            if (requestedMode.Equals(SendInputMode, StringComparison.OrdinalIgnoreCase))
            {
                var sendInput = new SendInputBackend();
                log?.Invoke("输入后端: SendInput（兼容模式）");
                return sendInput;
            }

            var virtualHid = new VirtualHidBackend();
            if (virtualHid.IsAvailable)
            {
                log?.Invoke("输入后端: Virtual HID");
                return virtualHid;
            }

            if (!config.AllowSendInputFallback)
            {
                log?.Invoke("输入后端: Virtual HID 未连接；已禁止 SendInput 回退");
                return virtualHid;
            }

            virtualHid.Dispose();
            var fallback = new SendInputBackend();
            log?.Invoke("输入后端: 未发现 Virtual HID，当前回退到 SendInput");
            return fallback;
        }
    }
}
