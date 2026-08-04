using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
// using System.Windows; // 移除以避免 Point/Size 歧义
using WeChatOcr; // 确保已通过 NuGet 安装 WeChatOcr.Lite

using moshushou.Yolo;

namespace moshushou
{
    public class ScreenshotHelper : IDisposable
    {
        #region Win32 API Imports
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
        #endregion

        // *** NEW ***: 定义截图裁剪的精确参数
        // private const int LEFT_CROP = 270; // 移入 Config
        // private const int CROP_HEIGHT = 53; // 移入 Config
        // private const int WECHAT_RIGHT_CROP = 125; // 移入 Config
        // private const int WEWORK_RIGHT_CROP = 100; // 移入 Config
        private const int DEFAULT_RIGHT_CROP = 300; // 备用值
        // 搜索列表识别阈值：单独降低，避免“搜索群聊/最近搜索群聊”漏检
        private const float SEARCH_YOLO_CONF_THRESHOLD = 0.15f;
        private const float SEARCH_YOLO_IOU_THRESHOLD = 0.45f;
        // 发送失败后的布局校验阈值：略低于默认值，减少“明明在聊天窗口却漏检”的概率
        private const float LAYOUT_YOLO_CONF_THRESHOLD = 0.18f;
        private const float LAYOUT_YOLO_IOU_THRESHOLD = 0.45f;

        private readonly string _baseDirectory;
        private readonly Action<string> _logAction;
        private readonly SearchConfig _config;
        private readonly YoloWindowDetector _yoloDetector;
        private readonly SemaphoreSlim _ocrSemaphore = new SemaphoreSlim(1, 1);
        private readonly object _ocrEngineSync = new object();
        private readonly object _detectionCacheSync = new object();
        private ImageOcr? _ocrEngine;
        private bool _disposed;
        private static long _ocrRequestSeq = 0;
        private static long _ocrFailureDebugSeq = 0;
        private static int _ocrActiveCount = 0;
        private const int OCR_TIMEOUT_MS = 8000;
        private const int DETECTION_CACHE_TTL_MS = 2000;

        private sealed class OcrTextRegion
        {
            public string Text { get; init; } = string.Empty;
            public int Left { get; init; }
            public int Top { get; init; }
            public int Right { get; init; }
            public int Bottom { get; init; }
        }

        private sealed class OcrRunResult
        {
            public string Text { get; init; } = string.Empty;
            public List<OcrTextRegion> Regions { get; init; } = new List<OcrTextRegion>();
            public bool TimedOut { get; init; }
            public bool Canceled { get; init; }
        }

        private sealed class DetectionCacheEntry
        {
            public IntPtr Hwnd { get; init; }
            public RECT WindowRect { get; init; }
            public DateTime CreatedUtc { get; init; }
            public List<YoloResult> Results { get; init; } = new List<YoloResult>();
        }

        private sealed class GroupTitleOcrSample
        {
            public DateTime CapturedAt { get; init; }
            public IntPtr Hwnd { get; init; }
            public bool IsWework { get; init; }
            public int WindowWidth { get; init; }
            public int WindowHeight { get; init; }
            public Rectangle GroupBox { get; init; }
            public float YoloConfidence { get; init; }
            public string DetectionSummary { get; init; } = string.Empty;
            public string RawText { get; init; } = string.Empty;
            public string CleanedText { get; init; } = string.Empty;
            public byte[] GroupPngBytes { get; init; } = Array.Empty<byte>();
        }

        private DetectionCacheEntry? _detectionCache;
        private readonly object _lastGroupTitleOcrSampleSync = new object();
        private GroupTitleOcrSample? _lastGroupTitleOcrSample;
        private static readonly JsonSerializerOptions FailedOcrDebugJsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        public string FailedOcrDebugDirectory => Path.Combine(_baseDirectory, "OCR_Failure_Debug");


        public ScreenshotHelper(string baseStorageDirectory, SearchConfig config, Action<string> logAction = null)
        {
            _baseDirectory = baseStorageDirectory;
            _config = config;
            _logAction = logAction;
            try
            {
                _yoloDetector = new YoloWindowDetector();
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"⚠️ YOLO 模型加载失败: {ex.Message}");
            }
        }

        // [已禁用] Debug_Yolo 调试目录相关代码
        // private static string EnsureDebugYoloDir()
        // {
        //     string datePart = DateTime.Now.ToString("yyyyMMdd");
        //
        //     string primary = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Debug_Yolo", datePart);
        //     try
        //     {
        //         Directory.CreateDirectory(primary);
        //         return primary;
        //     }
        //     catch (Exception ex1)
        //     {
        //         System.Diagnostics.Debug.WriteLine($"[Debug_Yolo] 主目录创建失败: {primary}, err={ex1.Message}");
        //     }
        //
        //     string fallback = Path.Combine(Environment.CurrentDirectory, "Debug_Yolo", datePart);
        //     try
        //     {
        //         Directory.CreateDirectory(fallback);
        //         return fallback;
        //     }
        //     catch (Exception ex2)
        //     {
        //         System.Diagnostics.Debug.WriteLine($"[Debug_Yolo] 备选目录创建失败: {fallback}, err={ex2.Message}");
        //     }
        //
        //     string tempFallback = Path.Combine(Path.GetTempPath(), "moshushou", "Debug_Yolo", datePart);
        //     Directory.CreateDirectory(tempFallback);
        //     System.Diagnostics.Debug.WriteLine($"[Debug_Yolo] 使用临时目录: {tempFallback}");
        //     return tempFallback;
        // }

        private static string SanitizeDebugToken(string value, int maxLen = 32)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "empty";
            }

            string token = value.Trim();
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                token = token.Replace(c, '_');
            }
            token = token.Replace(" ", "_");
            if (token.Length > maxLen)
            {
                token = token.Substring(0, maxLen);
            }
            return string.IsNullOrWhiteSpace(token) ? "empty" : token;
        }

        private static string BuildBboxText(Rectangle bbox)
        {
            return $"({bbox.X},{bbox.Y},{bbox.Width},{bbox.Height})";
        }

        private static string BuildDetectionSummary(IEnumerable<YoloResult> results, int maxItems = 8)
        {
            if (results == null)
            {
                return "none";
            }

            var list = results
                .Take(maxItems)
                .Select(r => $"{r.LabelName}:{r.Confidence:F2}@{BuildBboxText(r.BBox)}")
                .ToList();

            if (list.Count == 0)
            {
                return "none";
            }

            return string.Join(" | ", list);
        }

        private void CacheDetections(IntPtr hwnd, RECT rect, IEnumerable<YoloResult> results)
        {
            var snapshot = results.Select(r => new YoloResult
            {
                LabelId = r.LabelId,
                LabelName = r.LabelName,
                Confidence = r.Confidence,
                BBox = r.BBox
            }).ToList();

            lock (_detectionCacheSync)
            {
                _detectionCache = new DetectionCacheEntry
                {
                    Hwnd = hwnd,
                    WindowRect = rect,
                    CreatedUtc = DateTime.UtcNow,
                    Results = snapshot
                };
            }
        }

        private bool TryGetCachedDetections(IntPtr hwnd, RECT rect, out List<YoloResult> results)
        {
            lock (_detectionCacheSync)
            {
                DetectionCacheEntry? cache = _detectionCache;
                bool sameRect = cache != null &&
                                cache.WindowRect.Left == rect.Left &&
                                cache.WindowRect.Top == rect.Top &&
                                cache.WindowRect.Right == rect.Right &&
                                cache.WindowRect.Bottom == rect.Bottom;
                bool fresh = cache != null &&
                             (DateTime.UtcNow - cache.CreatedUtc).TotalMilliseconds <= DETECTION_CACHE_TTL_MS;

                if (cache != null && cache.Hwnd == hwnd && sameRect && fresh)
                {
                    results = cache.Results;
                    return true;
                }
            }

            results = new List<YoloResult>();
            return false;
        }

        private (bool success, int x, int y) GetInputCoordinatesFromDetections(
            IEnumerable<YoloResult> results,
            int screenX,
            int screenY)
        {
            var chatBox = results
                .Where(r => r.LabelName == YoloWindowDetector.Label_ChatBox)
                .OrderByDescending(r => r.Confidence)
                .FirstOrDefault();
            if (chatBox == null) return (false, 0, 0);

            int inputCenterX = chatBox.BBox.X + chatBox.BBox.Width / 2;
            int inputCenterY = chatBox.BBox.Y + (int)(chatBox.BBox.Height * 0.75);
            int clickX = screenX + inputCenterX + Random.Shared.Next(-5, 6);
            int clickY = screenY + inputCenterY + Random.Shared.Next(-5, 6);
            return (true, clickX, clickY);
        }

        private static bool IsSameDetection(
            YoloResult first,
            Rectangle firstScreenBBox,
            YoloResult second,
            Rectangle secondScreenBBox,
            int maxCenterDistance = 30)
        {
            if (first.LabelName != second.LabelName) return false;

            int firstCenterX = firstScreenBBox.X + firstScreenBBox.Width / 2;
            int firstCenterY = firstScreenBBox.Y + firstScreenBBox.Height / 2;
            int secondCenterX = secondScreenBBox.X + secondScreenBBox.Width / 2;
            int secondCenterY = secondScreenBBox.Y + secondScreenBBox.Height / 2;
            int deltaX = firstCenterX - secondCenterX;
            int deltaY = firstCenterY - secondCenterY;
            return deltaX * deltaX + deltaY * deltaY < maxCenterDistance * maxCenterDistance;
        }

        // [已禁用] Debug_Yolo 调试目录相关代码
        // private string SaveDebugRawImage(Bitmap bitmap, string prefix)
        // {
        //     try
        //     {
        //         string debugDir = EnsureDebugYoloDir();
        //         string path = Path.Combine(debugDir, $"{prefix}_{DateTime.Now:HHmmss_fff}.png");
        //         bitmap.Save(path, ImageFormat.Png);
        //         return path;
        //     }
        //     catch (Exception ex)
        //     {
        //         System.Diagnostics.Debug.WriteLine($"[Debug_Yolo] 保存原图失败: prefix={prefix}, err={ex.Message}");
        //         return null;
        //     }
        // }

        // [已禁用] Debug_Yolo 调试目录相关代码
        // private string SaveDebugAnnotatedImage(Bitmap bitmap, List<YoloResult> results, string prefix)
        // {
        //     try
        //     {
        //         if (_yoloDetector == null) return null;
        //         string debugDir = EnsureDebugYoloDir();
        //         string path = Path.Combine(debugDir, $"{prefix}_{DateTime.Now:HHmmss_fff}_ann.png");
        //         _yoloDetector.InferenceWrapper.SaveDebugImage(bitmap, results, path);
        //         return path;
        //     }
        //     catch (Exception ex)
        //     {
        //         System.Diagnostics.Debug.WriteLine($"[Debug_Yolo] 保存标注图失败: prefix={prefix}, err={ex.Message}");
        //         return null;
        //     }
        // }

        private void LogLayoutDebug(string message)
        {
            _logAction?.Invoke(message);
            System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
        }

        private void LogOcrDebug(string message, bool mirrorToUi = false)
        {
            string line = $"[{DateTime.Now:HH:mm:ss.fff}] [OCRDBG] {message}";
            System.Diagnostics.Debug.WriteLine(line);
            if (mirrorToUi)
            {
                _logAction?.Invoke(line);
            }
        }

        private static string BuildLayoutLabelStats(IEnumerable<YoloResult> results)
        {
            if (results == null)
            {
                return "none";
            }

            var groups = results
                .GroupBy(r => r.LabelName)
                .OrderByDescending(g => g.Count())
                .ThenBy(g => g.Key)
                .Select(g => $"{g.Key}:Count={g.Count()},MaxConf={g.Max(x => x.Confidence):F2}")
                .ToList();

            return groups.Count == 0 ? "none" : string.Join(" | ", groups);
        }

        private static string BuildTopCandidatesText(IEnumerable<YoloResult> results, string labelName, int topN = 3)
        {
            if (results == null)
            {
                return "none";
            }

            var top = results
                .Where(r => r.LabelName == labelName)
                .OrderByDescending(r => r.Confidence)
                .Take(topN)
                .Select(r => $"{r.Confidence:F2}@{BuildBboxText(r.BBox)}")
                .ToList();

            return top.Count == 0 ? "none" : string.Join(" | ", top);
        }

        // [已禁用] Debug_Yolo 调试目录相关代码
        // private string SaveLayoutDebugDataFile(
        //     string appName,
        //     IntPtr hwnd,
        //     RECT rect,
        //     int attempt,
        //     int maxRetries,
        //     List<YoloResult> results,
        //     YoloResult? groupName,
        //     YoloResult? chatInfo,
        //     YoloResult? chatBox)
        // {
        //     try
        //     {
        //         string debugDir = EnsureDebugYoloDir();
        //         string fileName = $"LayoutVerify_{(appName == "企业微信" ? "WeWork" : "WeChat")}_R{attempt}_{DateTime.Now:HHmmss_fff}.txt";
        //         string path = Path.Combine(debugDir, fileName);
        //
        //         var sb = new StringBuilder();
        //         sb.AppendLine($"Time={DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
        //         sb.AppendLine($"App={appName}");
        //         sb.AppendLine($"Hwnd={hwnd}");
        //         sb.AppendLine($"Attempt={attempt}/{maxRetries}");
        //         sb.AppendLine($"WindowRect=({rect.Left},{rect.Top},{rect.Right},{rect.Bottom})");
        //         sb.AppendLine($"WindowSize={Math.Max(0, rect.Right - rect.Left)}x{Math.Max(0, rect.Bottom - rect.Top)}");
        //         sb.AppendLine($"LabelStats={BuildLayoutLabelStats(results)}");
        //         sb.AppendLine($"Top_GroupName={BuildTopCandidatesText(results, YoloWindowDetector.Label_GroupName)}");
        //         sb.AppendLine($"Top_ChatInfo={BuildTopCandidatesText(results, YoloWindowDetector.Label_ChatInfo)}");
        //         sb.AppendLine($"Top_ChatBox={BuildTopCandidatesText(results, YoloWindowDetector.Label_ChatBox)}");
        //
        //         if (groupName != null && chatInfo != null && chatBox != null)
        //         {
        //             int groupY = groupName.BBox.Y + groupName.BBox.Height / 2;
        //             int infoY = chatInfo.BBox.Y + chatInfo.BBox.Height / 2;
        //             int boxY = chatBox.BBox.Y + chatBox.BBox.Height / 2;
        //             sb.AppendLine($"OrderCheck=GroupY({groupY}) < InfoY({infoY}) < BoxY({boxY}) => {(groupY < infoY && infoY < boxY)}");
        //         }
        //         else
        //         {
        //             sb.AppendLine("OrderCheck=Skipped(MissingCoreLabel)");
        //         }
        //
        //         sb.AppendLine("Detections:");
        //         if (results == null || results.Count == 0)
        //         {
        //             sb.AppendLine("  (none)");
        //         }
        //         else
        //         {
        //             int idx = 1;
        //             foreach (var r in results.OrderByDescending(x => x.Confidence))
        //             {
        //                 sb.AppendLine($"  {idx,2}. Label={r.LabelName}, Conf={r.Confidence:F4}, BBox={BuildBboxText(r.BBox)}");
        //                 idx++;
        //             }
        //         }
        //
        //         File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        //         return path;
        //     }
        //     catch (Exception ex)
        //     {
        //         System.Diagnostics.Debug.WriteLine($"[Debug_Yolo] 保存布局文本失败: app={appName}, attempt={attempt}/{maxRetries}, err={ex.Message}");
        //         return null;
        //     }
        // }


        // ✅ 新增：获取当前窗口顶部的标题文字（完全基于 YOLO）
        public async Task<string> GetWeChatWindowTitleTextAsync(IntPtr targetHwnd, bool isWework)
        {
            try
            {
                if (_yoloDetector == null) return null;
                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) return null;

                // 🚀 [改版] 根据用户反馈，模型使用全窗口训练，因此截取整个窗口
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                int screenX = rect.Left;
                int screenY = rect.Top;

                if (width <= 0 || height <= 0) return null;

                // 🛡️ [防御] 防止截取到桌面
                if (IsDesktopPixelSize(rect) || IsSystemWindowClass(targetHwnd))
                {
                    _logAction?.Invoke($"⚠️ [GetTitle] 拦截到桌面/系统窗口截取请求: {targetHwnd}");
                    return null;
                }

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // 1. YOLO 识别
                    var yoloResults = _yoloDetector.Detect(bitmap);
                    CacheDetections(targetHwnd, rect, yoloResults);
                    
                    // 2. 保存调试图 [已禁用]
                    // string debugDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Debug_Yolo", DateTime.Now.ToString("yyyyMMdd"));
                    // Directory.CreateDirectory(debugDir);
                    // string debugFile = Path.Combine(debugDir, $"Title_{DateTime.Now:HHmmss_fff}.png");
                    // _yoloDetector.InferenceWrapper.SaveDebugImage(bitmap, yoloResults, debugFile);

                    // 🚨 [调整] 场景宽松验证 (Relaxed Scene Validation)
                    // 原则：只要有 "群名" 和 "输入框" 且置信度及格，就认为是聊天窗口。
                    // "聊天信息" (Label_ChatInfo) 并非所有群/人都有(或识别不稳定)，故不强求。
                    float threshold = 0.6f; 
                    bool hasGroupName = yoloResults.Any(r => r.LabelName == YoloWindowDetector.Label_GroupName && r.Confidence > threshold);
                    bool hasChatBox = yoloResults.Any(r => r.LabelName == YoloWindowDetector.Label_ChatBox && r.Confidence > threshold);
                    
                    // 仅记录信息用于调试，不作为硬性拦截条件
                    var chatInfo = yoloResults.FirstOrDefault(r => r.LabelName == YoloWindowDetector.Label_ChatInfo);
                    string chatInfoLog = chatInfo != null ? $"{chatInfo.Confidence:F2}" : "None";

                    if (!hasGroupName || !hasChatBox)
                    {
                        _logAction?.Invoke($"⚠️ [GetTitle] 场景验证不通过: 群名={hasGroupName}, 输入框={hasChatBox} (阈值>{threshold:F1}), InfoConf={chatInfoLog}");
                        return null; 
                    }

                    // 3. 筛选 "群聊名字"
                    var target = yoloResults
                        .Where(r => r.LabelName == YoloWindowDetector.Label_GroupName)
                        .OrderByDescending(r => r.Confidence)
                        .FirstOrDefault();

                    if (target != null)
                    {
                        var bbox = target.BBox;
                        
                        // 🔧 [回退] 用户要求去掉裁切扩大，直接使用 YOLO 原始框
                        using (var crop = new Bitmap(bbox.Width, bbox.Height))
                        using (var g = Graphics.FromImage(crop))
                        {
                            g.DrawImage(bitmap, new Rectangle(0, 0, bbox.Width, bbox.Height), bbox, GraphicsUnit.Pixel);
                        string ocrText = await PerformAdaptiveOcrAsync(crop, $"GetTitle/{(isWework ? "WeWork" : "WeChat")}");
                        string cleaned = CleanGroupName(ocrText);
                        RememberGroupTitleOcrSample(
                            targetHwnd,
                            isWework,
                            width,
                            height,
                            bbox,
                            target.Confidence,
                            BuildDetectionSummary(yoloResults),
                            ocrText,
                            cleaned,
                            crop);
                        _logAction?.Invoke($"🏷️ YOLO 识别窗口标题: '{cleaned}' (置信度:{target.Confidence:P0})");
                        return cleaned;
                        }
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 获取标题栏失败: {ex.Message}");
                return null;
            }
        }

        // ✅ [新增] 专门计算输入框点击坐标的方法 (增强：YOLO 辅助)
        public async Task<(bool success, int x, int y)> GetInputBoxClickCoordinatesAsync(IntPtr hwnd, bool isWework)
        {
            if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out RECT rect)) return (false, 0, 0);

            // 🚀 [改版] 尝试 YOLO 定位 聊天输入框
            try
            {
                if (_yoloDetector != null)
                {
                // 🚀 [改版] 截取整个窗口 (Full Window)
                int w = rect.Right - rect.Left;
                int h = rect.Bottom - rect.Top;
                int screenX = rect.Left;
                int screenY = rect.Top;

                if (w > 0 && h > 0)
                {
                    if (TryGetCachedDetections(hwnd, rect, out var cachedResults))
                    {
                        var cachedCoordinates = GetInputCoordinatesFromDetections(cachedResults, screenX, screenY);
                        if (cachedCoordinates.success)
                        {
                            _logAction?.Invoke($"⚡ 复用 YOLO 缓存定位输入框: ({cachedCoordinates.x}, {cachedCoordinates.y})");
                            return cachedCoordinates;
                        }
                    }

                    // 🛡️ [防御] 防止截取到桌面
                    if (IsDesktopPixelSize(rect) || IsSystemWindowClass(hwnd))
                    {
                         _logAction?.Invoke($"⚠️ [GetInput] 拦截到桌面/系统窗口截取请求: {hwnd}");
                         return (false, 0, 0);
                    }

                    using (var bitmap = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                    {
                        using (var g = Graphics.FromImage(bitmap))
                        {
                            g.CopyFromScreen(screenX, screenY, 0, 0, new Size(w, h), CopyPixelOperation.SourceCopy);
                        }

                            var results = _yoloDetector.Detect(bitmap);
                            CacheDetections(hwnd, rect, results);
                            
                            // 调试保存 [已禁用]
                            // string debugDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Debug_Yolo", DateTime.Now.ToString("yyyyMMdd"));
                            // Directory.CreateDirectory(debugDir);
                            // string debugFile = Path.Combine(debugDir, $"InputBox_{DateTime.Now:HHmmss_fff}.png");
                            // _yoloDetector.InferenceWrapper.SaveDebugImage(bitmap, results, debugFile);

                            var coordinates = GetInputCoordinatesFromDetections(results, screenX, screenY);
                            if (coordinates.success)
                            {
                                _logAction?.Invoke($"🎯 YOLO 定位输入框成功(微幅随机): ({coordinates.x}, {coordinates.y})");
                                return coordinates;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"⚠️ YOLO 输入框检测失败，使用兜底逻辑: {ex.Message}");
            }

            // 📐 坐标计算兜底逻辑 (如果 YOLO 失败)
            int x = rect.Left + (isWework ? 350 : 310);
            int y = rect.Bottom - (isWework ? 80 : 70);

            // 简单的越界检查
            if (x >= rect.Right) x = rect.Right - 50;
            if (y >= rect.Bottom) y = rect.Bottom - 20;

            return (true, x, y);
        }

        /// <summary>
        /// ✅ [新增] 验证聊天窗口布局是否正常 (基于 YOLO)
        /// 规则：
        /// 1. 必须识别到：群聊名字 (GroupName)
        /// 2. 必须识别到：聊天信息 (ChatInfo)
        /// 3. 必须识别到：聊天框 (ChatBox)
        /// 4. 顺序必须是：GroupName (上) -> ChatInfo (中) -> ChatBox (下)
        /// 5. 失败则认为窗口异常 (退出登录/遮挡/界面错乱)
        /// </summary>
        public async Task<bool> VerifyChatWindowLayoutAsync(IntPtr hwnd, bool isWework)
        {
            if (_yoloDetector == null) return true; // 没有 YOLO 只能盲信
            if (hwnd == IntPtr.Zero) return false;

            string appName = isWework ? "企业微信" : "微信";
            int maxRetries = 3;
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    LogLayoutDebug(
                        $"🔎 [布局验证] 开始第{i + 1}/{maxRetries}次: App={appName}, Hwnd={hwnd}, " +
                        $"ConfTh={LAYOUT_YOLO_CONF_THRESHOLD:F2}, IouTh={LAYOUT_YOLO_IOU_THRESHOLD:F2}");

                    if (!GetWindowRect(hwnd, out RECT rect)) return false;
                    int w = rect.Right - rect.Left;
                    int h = rect.Bottom - rect.Top;
                    
                    if (w <= 0 || h <= 0) return false;
                    LogLayoutDebug($"📐 [布局验证] 第{i + 1}次窗口尺寸: {w}x{h}, Rect=({rect.Left},{rect.Top},{rect.Right},{rect.Bottom})");

                    using (var bitmap = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                    {
                        using (var g = Graphics.FromImage(bitmap))
                        {
                            g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(w, h), CopyPixelOperation.SourceCopy);
                        }

                        var results = _yoloDetector.Detect(bitmap, LAYOUT_YOLO_CONF_THRESHOLD, LAYOUT_YOLO_IOU_THRESHOLD);
                        
                        string debugPrefix = $"LayoutVerify_{(isWework ? "WeWork" : "WeChat")}_R{i + 1}";
                        // [已禁用] Debug_Yolo 调试图保存
                        // string layoutRawPath = SaveDebugRawImage(bitmap, debugPrefix);
                        // string layoutAnnPath = SaveDebugAnnotatedImage(bitmap, results, debugPrefix);
                        string layoutSummary = BuildDetectionSummary(results, 12);
                        string labelStats = BuildLayoutLabelStats(results);
                        LogLayoutDebug($"🧾 [布局验证] 第{i + 1}次YOLO结果: {layoutSummary}");
                        LogLayoutDebug($"📊 [布局验证] 第{i + 1}次标签统计: {labelStats}");
                        LogLayoutDebug($"📌 [布局验证] GroupTop={BuildTopCandidatesText(results, YoloWindowDetector.Label_GroupName)}");
                        LogLayoutDebug($"📌 [布局验证] InfoTop={BuildTopCandidatesText(results, YoloWindowDetector.Label_ChatInfo)}");
                        LogLayoutDebug($"📌 [布局验证] BoxTop={BuildTopCandidatesText(results, YoloWindowDetector.Label_ChatBox)}");
                        // [已禁用] Debug_Yolo 调试图日志
                        // if (!string.IsNullOrEmpty(layoutRawPath) || !string.IsNullOrEmpty(layoutAnnPath))
                        // {
                        //     LogLayoutDebug($"🖼️ [布局验证] 第{i + 1}次调试图: Raw={layoutRawPath}, Ann={layoutAnnPath}");
                        // }

                        // 获取置信度最高的组件
                        var groupName = results.Where(r => r.LabelName == YoloWindowDetector.Label_GroupName).OrderByDescending(r => r.Confidence).FirstOrDefault();
                        var chatInfo = results.Where(r => r.LabelName == YoloWindowDetector.Label_ChatInfo).OrderByDescending(r => r.Confidence).FirstOrDefault();
                        var chatBox = results.Where(r => r.LabelName == YoloWindowDetector.Label_ChatBox).OrderByDescending(r => r.Confidence).FirstOrDefault();

                        // [已禁用] Debug_Yolo 布局数据保存
                        // string layoutDataPath = SaveLayoutDebugDataFile(appName, hwnd, rect, i + 1, maxRetries, results, groupName, chatInfo, chatBox);
                        // if (!string.IsNullOrEmpty(layoutDataPath))
                        // {
                        //     LogLayoutDebug($"📝 [布局验证] 第{i + 1}次布局数据: {layoutDataPath}");
                        // }

                        // 规则1: 核心组件必须存在
                        if (groupName == null || chatInfo == null || chatBox == null)
                        {
                            string groupInfo = groupName != null ? $"{groupName.Confidence:F2}@{BuildBboxText(groupName.BBox)}" : "null";
                            string chatInfoText = chatInfo != null ? $"{chatInfo.Confidence:F2}@{BuildBboxText(chatInfo.BBox)}" : "null";
                            string boxInfo = chatBox != null ? $"{chatBox.Confidence:F2}@{BuildBboxText(chatBox.BBox)}" : "null";
                            LogLayoutDebug($"❌ [布局验证] 核心组件缺失 (第{i+1}次): Group={groupInfo}, Info={chatInfoText}, Box={boxInfo}");
                             if (i < maxRetries - 1) 
                             {
                                 await Task.Delay(500);
                                 continue;
                             }
                             return false;
                        }

                        // 规则2: Y轴顺序检查 (上 -> 下)
                        // GroupName.Y < ChatInfo.Y < ChatBox.Y
                        int groupY = groupName.BBox.Y + groupName.BBox.Height / 2;
                        int infoY = chatInfo.BBox.Y + chatInfo.BBox.Height / 2;
                        int boxY = chatBox.BBox.Y + chatBox.BBox.Height / 2;

                        if (groupY < infoY && infoY < boxY)
                        {
                            LogLayoutDebug($"✅ [布局验证] 窗口布局正常 (Group -> Info -> Box), Y轴中心: Group={groupY}, Info={infoY}, Box={boxY}");
                            return true;
                        }
                        else
                        {
                            LogLayoutDebug($"❌ [布局验证] 组件顺序错误: GroupY={groupY}, InfoY={infoY}, BoxY={boxY}");
                            if (i < maxRetries - 1) 
                            { 
                                await Task.Delay(500); 
                                continue; 
                            }
                            return false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogLayoutDebug($"💥 [布局验证] 发生异常: {ex.Message}");
                    return false;
                }
            }
            return false;
        }











        /// <summary>
        /// ✅ [智能优化版] 模糊匹配
        /// 针对 OCR 误差、长文本截断、包含关系进行了专门优化
        /// </summary>
        /// <param name="expected">目标搜索词 (例如: "美诗安轩官方旗舰店")</param>
        /// <param name="actual">OCR识别出的文本 (例如: "美诗安轩官方..(54人)")</param>
        /// <returns>是否匹配</returns>
        public bool IsFuzzyMatch(string expected, string actual)
        {
            if (string.IsNullOrWhiteSpace(actual)) return false;
            if (string.IsNullOrWhiteSpace(expected)) return false;

            // 1. 快速检查：未处理前如果包含，直接返回 (最快)
            if (actual.Contains(expected, StringComparison.OrdinalIgnoreCase) ||
                expected.Contains(actual, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // ✅ [新增] 规范化特殊字符（OCR 经常把这些字符识别错或遗漏）
            string NormalizeSpecialChars(string s)
            {
                // 波浪号统一：全角～ → 半角~，或直接移除
                s = s.Replace("～", "~").Replace("〜", "~");
                // 破折号统一
                s = s.Replace("—", "-").Replace("–", "-").Replace("一", "-");
                // 省略号统一
                s = s.Replace("…", "..").Replace("。。", "..");
                // 中文括号统一
                s = s.Replace("（", "(").Replace("）", ")");
                s = s.Replace("【", "[").Replace("】", "]");
                return s;
            }

            string normalizedTarget = NormalizeSpecialChars(expected);
            string normalizedOCR = NormalizeSpecialChars(actual);

            // 规范化后再检查包含关系
            if (normalizedOCR.Contains(normalizedTarget, StringComparison.OrdinalIgnoreCase) ||
                normalizedTarget.Contains(normalizedOCR, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 2. 深度清洗：
            // - 去除所有空白字符 (\s)
            // - 去除常见标点符号 (包括中文标点、截断用的点、波浪号等)
            // - 统一转小写
            string pattern = @"\s+|[.,;:'""()\-\[\]{}<>/\\|、，。；：""（）—…\.~～〜]";

            // ✅ 定义局部函数，供后续分词匹配使用
            string Clean(string s) => Regex.Replace(s, pattern, "").ToLowerInvariant();

            string cleanTarget = Clean(normalizedTarget);
            string cleanOCR = Clean(normalizedOCR);

            // 防止清洗后为空
            if (string.IsNullOrEmpty(cleanTarget) || string.IsNullOrEmpty(cleanOCR)) return false;

            // 3. 【核心优化】智能前缀匹配 (专门解决 "张旭彬...官方旗舰店" 变成 "张旭彬...官方.." 的问题)
            // 逻辑：如果清洗后的 OCR 结果，是 目标词 的"开头部分"，且长度足够长，视为匹配。
            int minPrefixLen = 4; // 至少匹配前4个字才算数，防止匹配到"张三"这种泛滥的词
            if (cleanTarget.Length >= minPrefixLen && cleanOCR.Length >= minPrefixLen)
            {
                // 截取两者的最短长度进行比较
                int checkLen = Math.Min(cleanTarget.Length, cleanOCR.Length);

                // 这里的 checkLen 可能会比较长，我们主要看 OCR 是否匹配 Target 的前半部分
                string targetPrefix = cleanTarget.Substring(0, checkLen);
                string ocrPrefix = cleanOCR.Substring(0, checkLen);

                if (targetPrefix == ocrPrefix)
                {
                    // System.Diagnostics.Debug.WriteLine($"[Fuzzy] 前缀完全匹配成功: {targetPrefix}");
                    return true;
                }

                // 容错前缀：如果前6个字里，有5个字一样，也算匹配 (应对OCR把开头某个字识别错的情况)
                if (checkLen >= 5)
                {
                    int prefixDist = LevenshteinDistance(targetPrefix, ocrPrefix);
                    if (prefixDist <= 1) // 允许错1个字
                    {
                        // System.Diagnostics.Debug.WriteLine($"[Fuzzy] 前缀容错匹配成功 (错{prefixDist}字)");
                        return true;
                    }
                }
            }

            // 4. 包含关系 (清洗后)
            if (cleanOCR.Contains(cleanTarget)) return true;

            // 反向包含 (针对 target 很长，OCR 只是其中一部分的情况)
            // 但要求 OCR 至少有一定长度，防止 target="A" ocr="ABCDEFG" 这种误判
            // 🎯 [优化] 降低阈值到 2，应对 OCR 将群名切分为多个短语的情况 (如 "中通" "快递")
            if (cleanTarget.Contains(cleanOCR) && cleanOCR.Length >= 2) return true;

            // 5. 【增强】分词乱序匹配 (解决 "A B" 变成 "B A" 或 "BA" 的问题)
            // 场景：SearchText="tb522008016 郭润斌88"  OCR="郭润斌88..tb522008016"
            // 逻辑：将 Target 按空格/标点拆分，如果 OCR 包含所有核心片段，视为匹配
            var parts = expected.Split(new[] { ' ', ',', '，', '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 1) // 只有当 Target 本身就是复合词时才用这个逻辑
            {
                int matchCount = 0;
                int totalLen = 0;
                foreach (var part in parts)
                {
                    string cleanPart = Clean(part);
                    if (string.IsNullOrEmpty(cleanPart)) continue;

                    // 忽略太短的连接词
                    if (cleanPart.Length < 2 && parts.Length > 2) continue; 

                    totalLen++;
                    if (cleanOCR.Contains(cleanPart))
                    {
                        matchCount++;
                    }
                }

                // 如果片段匹配率超过 80% (或全部匹配)，视为成功
                // 例如 3个词匹配了3个，或者 5个词匹配了4个
                if (totalLen > 0 && (double)matchCount / totalLen >= 0.99) // 这里要求极其严格，必须全中(针对用户案例)或者允许极少缺失
                {
                    // 再次检查长度差异，防止匹配到"包含"但"多出太多内容"的情况（虽然CheckSearchResult通常不会）
                    return true;
                }
                
                // 宽松模式：针对用户这种 "tb522008016 郭润斌88 雄兴风扇" -> "郭润斌88雄兴风扇tb522008016"
                // 只要所有部分都找到了，不管顺序，直接 True
                if (matchCount == totalLen && totalLen > 0) return true;
            }

            // 6. 莱文斯坦距离 (兜底逻辑)
            int dist = LevenshteinDistance(cleanTarget, cleanOCR);
            int maxLength = Math.Max(cleanTarget.Length, cleanOCR.Length);
            double similarity = 1.0 - (double)dist / maxLength;

            // 【优化】动态阈值：
            // 字符串越短，要求越严格；字符串越长，允许误差越大。
            double threshold = 0.5;
            if (maxLength <= 4) threshold = 0.75;      // 4个字以内，必须很像 (允许错1个)
            else if (maxLength <= 8) threshold = 0.6;  // 8个字以内，允许错一点
            else threshold = 0.4;                      // 超长字符串，允许错更多 (适应截断和乱码)

            // System.Diagnostics.Debug.WriteLine($"[Fuzzy] 相似度: {similarity:F2} (阈值: {threshold})");
            return similarity >= threshold;
        }


        /// <summary>
        /// 计算两个字符串的莱文斯坦距离 (编辑距离)
        /// </summary>
        private int LevenshteinDistance(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            if (n == 0) return m;
            if (m == 0) return n;

            for (int i = 0; i <= n; d[i, 0] = i++) { }
            for (int j = 0; j <= m; d[0, j] = j++) { }

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            return d[n, m];
        }



        /// <summary>
        /// ✅ [YOLO 精准版] 发送验证保险机制
        /// 通过 YOLO 定位 ChatInfo（聊天信息区）和 ChatBox（聊天框）的精确区域，
        /// 快速截取 3 屏防刷屏，OCR 识别后使用模糊匹配对比我们发送的关键词。
        /// 只有 3 屏中完全找不到我们发送的文本才判定为失败。
        /// </summary>
        /// <param name="hwnd">目标聊天窗口句柄</param>
        /// <param name="isWework">是否企业微信</param>
        /// <param name="keyword">我们发送的关键词（用于对比验证）</param>
        /// <returns>
        /// success: 是否验证成功
        /// verifyMethod: 验证方式描述
        /// chatInfoText: 聊天信息区 OCR 文本（最后一屏）
        /// chatBoxText: 聊天框 OCR 文本（最后一屏）
        /// </returns>
        public async Task<(bool success, string verifyMethod, string chatInfoText, string chatBoxText)> VerifySendWithYoloAsync(
            IntPtr hwnd, bool isWework, string keyword, CancellationToken token = default)
        {
            Action<string> Log = (msg) => System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [YoloVerify] {msg}");

            if (hwnd == IntPtr.Zero)
            {
                Log("❌ 窗口句柄无效");
                return (false, "窗口无效", null, null);
            }

            if (_yoloDetector == null)
            {
                Log("❌ YOLO 未初始化，无法验证");
                return (false, "YOLO未初始化", null, null);
            }

            // 快速连续截取 3 屏，几乎无间隔，防止消息被刷屏上去识别不到
            const int FRAME_COUNT = 3;
            const int FRAME_INTERVAL_MS = 30; // 极短间隔，保证速度

            string lastChatInfoText = null;
            string lastChatBoxText = null;

            for (int frame = 0; frame < FRAME_COUNT; frame++)
            {
                if (token.IsCancellationRequested)
                {
                    Log("🛑 验证取消");
                    return (false, "已取消", lastChatInfoText, lastChatBoxText);
                }

                var frameSw = Stopwatch.StartNew();
                if (frame > 0)
                {
                    try { await Task.Delay(FRAME_INTERVAL_MS, token); }
                    catch (OperationCanceledException) { return (false, "已取消", lastChatInfoText, lastChatBoxText); }
                }

                try
                {
                    if (!GetWindowRect(hwnd, out RECT rect)) continue;
                    int w = rect.Right - rect.Left;
                    int h = rect.Bottom - rect.Top;
                    if (w <= 0 || h <= 0) continue;

                    // 截取整个窗口
                    using (var bitmap = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                    {
                        using (var g = Graphics.FromImage(bitmap))
                        {
                            g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(w, h), CopyPixelOperation.SourceCopy);
                        }

                        // YOLO 检测
                        var detectSw = Stopwatch.StartNew();
                        var results = _yoloDetector.Detect(bitmap);
                        Log($"🧪 第{frame + 1}屏 YOLO完成: 耗时={detectSw.ElapsedMilliseconds}ms, 目标数={results?.Count ?? 0}, 窗口={w}x{h}");

                        var chatInfo = results
                            .Where(r => r.LabelName == YoloWindowDetector.Label_ChatInfo)
                            .OrderByDescending(r => r.Confidence)
                            .FirstOrDefault();
                        var chatBox = results
                            .Where(r => r.LabelName == YoloWindowDetector.Label_ChatBox)
                            .OrderByDescending(r => r.Confidence)
                            .FirstOrDefault();

                        if (chatInfo == null && chatBox == null)
                        {
                            Log($"⚠️ 第{frame + 1}屏: YOLO 未检测到 ChatInfo 和 ChatBox");
                            continue;
                        }

                        // === 验证1: ChatInfo 区域（聊天信息）底部 1/3 ===
                        if (chatInfo != null)
                        {
                            // 裁剪 ChatInfo 底部 1/3（最新消息区域）
                            int infoH = chatInfo.BBox.Height;
                            int cropH = Math.Max(infoH / 3, 50); // 至少 50px
                            int cropY = chatInfo.BBox.Y + infoH - cropH;
                            int cropX = chatInfo.BBox.X;
                            int cropW = chatInfo.BBox.Width;

                            // 边界安全检查
                            if (cropX < 0) cropX = 0;
                            if (cropY < 0) cropY = 0;
                            if (cropX + cropW > w) cropW = w - cropX;
                            if (cropY + cropH > h) cropH = h - cropY;

                            if (cropW > 0 && cropH > 0)
                            {
                                using (var chatInfoCrop = bitmap.Clone(
                                    new Rectangle(cropX, cropY, cropW, cropH),
                                    bitmap.PixelFormat))
                                {
                                    Log($"🧪 第{frame + 1}屏 ChatInfo裁剪: {cropW}x{cropH}@({cropX},{cropY}), Conf={chatInfo.Confidence:F2}");
                                    lastChatInfoText = await PerformOcrAsync(chatInfoCrop, $"YoloVerify/F{frame + 1}/ChatInfo", token);
                                    if (token.IsCancellationRequested || string.Equals(lastChatInfoText, "OCR识别取消", StringComparison.Ordinal))
                                    {
                                        return (false, "已取消", lastChatInfoText, lastChatBoxText);
                                    }
                                    Log($"📝 第{frame + 1}屏 ChatInfo OCR: {lastChatInfoText?.Substring(0, Math.Min(80, lastChatInfoText?.Length ?? 0))}");
                                    if (string.Equals(lastChatInfoText, "OCR识别超时", StringComparison.Ordinal))
                                    {
                                        Log($"⚠️ 第{frame + 1}屏 ChatInfo OCR 超时(8000ms)");
                                    }

                                    // 模糊匹配：对比我们发送的关键词
                                    if (!string.IsNullOrEmpty(lastChatInfoText) && IsFuzzyMatch(keyword, lastChatInfoText))
                                    {
                                        Log($"✅ 第{frame + 1}屏: ChatInfo 上屏验证成功 (模糊匹配命中)");
                                        return (true, $"YOLO上屏验证(第{frame + 1}屏)", lastChatInfoText, lastChatBoxText);
                                    }
                                }
                            }
                        }

                        // === 验证2: ChatBox 区域（聊天输入框）===
                        if (chatBox != null)
                        {
                            int boxX = chatBox.BBox.X;
                            int boxY = chatBox.BBox.Y;
                            int boxW = chatBox.BBox.Width;
                            int boxH = chatBox.BBox.Height;

                            // 边界安全检查
                            if (boxX < 0) boxX = 0;
                            if (boxY < 0) boxY = 0;
                            if (boxX + boxW > w) boxW = w - boxX;
                            if (boxY + boxH > h) boxH = h - boxY;

                            if (boxW > 0 && boxH > 0)
                            {
                                using (var chatBoxCrop = bitmap.Clone(
                                    new Rectangle(boxX, boxY, boxW, boxH),
                                    bitmap.PixelFormat))
                                {
                                    Log($"🧪 第{frame + 1}屏 ChatBox裁剪: {boxW}x{boxH}@({boxX},{boxY}), Conf={chatBox.Confidence:F2}");
                                    lastChatBoxText = await PerformOcrAsync(chatBoxCrop, $"YoloVerify/F{frame + 1}/ChatBox", token);
                                    if (token.IsCancellationRequested || string.Equals(lastChatBoxText, "OCR识别取消", StringComparison.Ordinal))
                                    {
                                        return (false, "已取消", lastChatInfoText, lastChatBoxText);
                                    }
                                    Log($"📝 第{frame + 1}屏 ChatBox OCR: {lastChatBoxText?.Substring(0, Math.Min(80, lastChatBoxText?.Length ?? 0))}");
                                    if (string.Equals(lastChatBoxText, "OCR识别超时", StringComparison.Ordinal))
                                    {
                                        Log($"⚠️ 第{frame + 1}屏 ChatBox OCR 超时(8000ms)");
                                    }

                                    // 如果 ChatBox 中不再包含我们的关键词，说明已经清空发送了
                                    if (!string.IsNullOrEmpty(lastChatBoxText) && !IsFuzzyMatch(keyword, lastChatBoxText))
                                    {
                                        Log($"✅ 第{frame + 1}屏: ChatBox 清空验证成功 (输入框不含关键词)");
                                        return (true, $"YOLO清空验证(第{frame + 1}屏)", lastChatInfoText, lastChatBoxText);
                                    }
                                    // ChatBox OCR 为空文本也算清空
                                    if (string.IsNullOrWhiteSpace(lastChatBoxText) || lastChatBoxText == "未识别到文字")
                                    {
                                        Log($"✅ 第{frame + 1}屏: ChatBox 为空，清空验证成功");
                                        return (true, $"YOLO清空验证(第{frame + 1}屏)", lastChatInfoText, lastChatBoxText);
                                    }
                                }
                            }
                        }

                        Log($"⏳ 第{frame + 1}屏: 未验证成功，继续下一屏... 本屏总耗时={frameSw.ElapsedMilliseconds}ms");
                    }
                }
                catch (Exception ex)
                {
                    Log($"💥 第{frame + 1}屏截取异常: {ex.Message}");
                }
            }

            // 3 屏全部未匹配到我们发送的文本 → 验证失败
            Log($"❌ 3屏均未检测到发送的关键词 '{keyword}'");
            return (false, "3屏均未匹配", lastChatInfoText, lastChatBoxText);
        }


        /// <summary>
        /// ✅ [新增] 截取整个窗口并进行 OCR 识别
        /// 用于检测全屏覆盖的安全验证页面
        /// </summary>
        public async Task<string> CaptureFullWindowAndOcrAsync(IntPtr targetHwnd)
        {
            try
            {
                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect))
                {
                    System.Diagnostics.Debug.WriteLine("[安全检测截图] 窗口句柄无效或无法获取窗口尺寸");
                    return null;
                }

                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;

                System.Diagnostics.Debug.WriteLine($"[安全检测截图] 窗口尺寸: {width}x{height}, 位置: ({rect.Left}, {rect.Top})");

                if (width <= 0 || height <= 0) return null;

                // 截取整个窗口
                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // ✅ [调试] 保存截图到文件以便验证
                    /*
                    try
                    {
                        string debugFolder = Path.Combine(_baseDirectory, "SecurityCheck_Debug");
                        Directory.CreateDirectory(debugFolder);
                        string debugPath = Path.Combine(debugFolder, $"security_check_{DateTime.Now:HHmmss_fff}.png");
                        bitmap.Save(debugPath, ImageFormat.Png);
                        System.Diagnostics.Debug.WriteLine($"[安全检测截图] 已保存到: {debugPath}");
                    }
                    catch (Exception saveEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[安全检测截图] 保存截图失败: {saveEx.Message}");
                    }
                    */

                    // 直接用原图进行 OCR (安全验证页面的文字通常较大，不需要放大)
                    string result = await PerformOcrAsync(bitmap, "SecurityCheck/FullWindow");
                    System.Diagnostics.Debug.WriteLine($"[安全检测截图] OCR 结果 (前300字): {result?.Substring(0, Math.Min(300, result?.Length ?? 0))}");
                    return result;
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 全窗口截图失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[安全检测截图] 异常: {ex.Message}");
                return null;
            }
        }





        /// <summary>
        /// 【V3 - 终极版】核心方法：
        /// 1. 直接截取窗口顶部已裁剪好的小图，提升效率。
        /// 2. 异步启动OCR识别任务，防止UI卡顿。
        /// 3. 通过回调函数返回识别结果。
        /// </summary>
        /// <param name="targetHwnd">要截图的目标窗口的句柄</param>
        /// <param name="storeName">商家名称</param>
        /// <param name="processName">窗口的进程名 ("WeChat" 或 "WXWork")</param>
        /// <param name="onOcrComplete">OCR识别完成后的回调委托</param>
        public void CaptureWindowTop(IntPtr targetHwnd, string storeName, string processName, Action<BusinessInfo> onOcrComplete)
        {
            try
            {
                if (targetHwnd == IntPtr.Zero)
                {
                    _logAction?.Invoke("❌ 截图失败: 窗口句柄无效。");
                    return;
                }

                if (!GetWindowRect(targetHwnd, out RECT rect) || rect.Right - rect.Left <= 0)
                {
                    _logAction?.Invoke($"❌ 截图失败: 无法获取窗口 '{storeName}' 的尺寸。");
                    return;
                }

                string appIdentifier = GetAppIdentifier(processName);
                int rightCrop = GetRightCropAmount(appIdentifier);
                int leftCrop = _config.WeChatCropLeft;
                int cropHeight = _config.WeChatCropHeight;

                int cropWidth = rect.Right - rect.Left - leftCrop - rightCrop;
                if (cropWidth <= 0)
                {
                    _logAction?.Invoke($"❌ 截图失败: 窗口宽度不足，无法裁剪。");
                    return;
                }

                // 直接创建裁剪后尺寸的Bitmap
                using (var bitmap = new Bitmap(cropWidth, cropHeight, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        // 从屏幕的指定位置直接复制已裁剪的区域到Bitmap中
                        graphics.CopyFromScreen(rect.Left + leftCrop, rect.Top, 0, 0, new Size(cropWidth, cropHeight), CopyPixelOperation.SourceCopy);
                    }

                    // ✅ [已禁用调试] 不再保存截图到文件
                    /*
                    string dateFolder = $"{DateTime.Now:yyyyMMdd}_OCR_Screenshots";
                    string fullDirectoryPath = Path.Combine(_baseDirectory, dateFolder);
                    Directory.CreateDirectory(fullDirectoryPath);

                    string safeStoreName = string.Join("_", storeName.Split(Path.GetInvalidFileNameChars()));
                    string initialFileName = $"{safeStoreName}_{appIdentifier}.png";
                    string filePath = Path.Combine(fullDirectoryPath, initialFileName);

                    bitmap.Save(filePath, ImageFormat.Png);
                    _logAction?.Invoke($"✅ 截图 '{initialFileName}' 已保存，正在启动后台OCR...");
                    
                    // 文件模式的 OCR
                    Task.Run(() => PerformOcrAndRenameAsync(filePath, storeName, appIdentifier, onOcrComplete));
                    */

                    // ✅ [优化] 直接在内存中进行 OCR，不保存文件
                    _logAction?.Invoke($"✅ 正在后台 OCR 识别 '{storeName}'...");
                    var bitmapCopy = (Bitmap)bitmap.Clone();
                    _ = Task.Run(async () =>
                    {
                        using (bitmapCopy)
                        {
                            BusinessInfo ocrResult = new BusinessInfo { StoreName = storeName, Source = appIdentifier };
                            try
                            {
                                string rawText = await PerformAdaptiveOcrAsync(bitmapCopy, $"CaptureTop/Inline/{storeName}");
                                ocrResult.GroupName = CleanGroupName(rawText);
                            }
                            catch (Exception ex)
                            {
                                ocrResult.GroupName = $"[OCR识别失败: {ex.Message}]";
                                _logAction?.Invoke($"💥 OCR处理失败 '{storeName}': {ex.Message}");
                            }
                            finally
                            {
                                onOcrComplete?.Invoke(ocrResult);
                            }
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 截图或OCR启动时发生严重错误: {ex.Message}");
            }
        }



        /// <summary>
        /// 【后台任务】对指定图片执行OCR，成功后重命名文件，并通过回调返回结果。
        /// ✅ 修复：使用 MemoryStream 加载图片，防止文件被锁定导致重命名失败
        /// </summary>
        private async Task PerformOcrAndRenameAsync(string imagePath, string storeName, string appIdentifier, Action<BusinessInfo> onOcrComplete)
        {
            BusinessInfo ocrResult = new BusinessInfo { StoreName = storeName, Source = appIdentifier };
            try
            {
                string recognizedGroupName = null;

                // 1. 读取文件到内存，随即释放文件句柄
                byte[] fileBytes;
                try
                {
                    fileBytes = File.ReadAllBytes(imagePath);
                }
                catch (IOException)
                {
                    // 如果文件刚生成可能被短暂占用，稍等一下重试
                    await Task.Delay(100);
                    fileBytes = File.ReadAllBytes(imagePath);
                }

                // 2. 在内存中进行图像处理和 OCR
                using (var ms = new MemoryStream(fileBytes))
                using (var originalBitmap = new Bitmap(ms))
                {
                    // 依然保持放大策略以确保群名识别准确率 (群名文字通常较小)
                    // 🚀 [优化] 统一使用 PreprocessForOcr
                    string rawText = await PerformAdaptiveOcrAsync(originalBitmap, $"CaptureTop/File/{storeName}");
                    recognizedGroupName = CleanGroupName(rawText);
                } // 离开 using 块，Bitmap 资源释放

                ocrResult.GroupName = recognizedGroupName;

                // 3. 重命名文件 (此时文件未被锁定)
                if (!string.IsNullOrEmpty(recognizedGroupName) && !recognizedGroupName.Contains("失败"))
                {
                    string safeGroupName = string.Join("_", recognizedGroupName.Split(Path.GetInvalidFileNameChars()));
                    if (safeGroupName.Length > 50) safeGroupName = safeGroupName.Substring(0, 50); // 限制长度

                    string newFileName = $"{Path.GetFileNameWithoutExtension(imagePath)}_[{safeGroupName}].png";
                    string newFilePath = Path.Combine(Path.GetDirectoryName(imagePath), newFileName);

                    try
                    {
                        if (File.Exists(newFilePath)) File.Delete(newFilePath); // 防止重名冲突
                        File.Move(imagePath, newFilePath);
                    }
                    catch (Exception renameEx)
                    {
                        _logAction?.Invoke($"⚠️ 文件重命名失败: {renameEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                ocrResult.GroupName = $"[OCR识别失败: {ex.Message}]";
                _logAction?.Invoke($"💥 OCR处理失败 '{storeName}': {ex.Message}");
            }
            finally
            {
                // 无论成功与否，都调用回调
                onOcrComplete?.Invoke(ocrResult);
            }
        }



        private ImageOcr GetOrCreateOcrEngine()
        {
            lock (_ocrEngineSync)
            {
                ObjectDisposedException.ThrowIf(_disposed, this);
                return _ocrEngine ??= new ImageOcr();
            }
        }

        /// <summary>
        /// 自动发送失败后采集当前群名区域。仅保存群名小图，不保存聊天正文。
        /// </summary>
        public async Task<string?> CaptureFailedGroupOcrDebugAsync(
            IntPtr targetHwnd,
            bool isWework,
            string? storeName,
            string? expectedGroupName,
            string? failureStage,
            string? retryPhase,
            CancellationToken token = default)
        {
            if (!_config.EnableFailedOcrDebugCapture || _disposed)
            {
                return null;
            }

            try
            {
                return await Task.Run(async () =>
                {
                    token.ThrowIfCancellationRequested();

                    GroupTitleOcrSample? originalSample = GetRecentGroupTitleOcrSample(
                        targetHwnd,
                        isWework,
                        TimeSpan.FromSeconds(20));
                    if (originalSample != null)
                    {
                        return SaveFailedOcrDebugRecord(
                            targetHwnd,
                            isWework,
                            storeName,
                            expectedGroupName,
                            failureStage,
                            retryPhase,
                            originalSample.CapturedAt,
                            originalSample.WindowWidth,
                            originalSample.WindowHeight,
                            originalSample.GroupBox,
                            originalSample.YoloConfidence,
                            originalSample.DetectionSummary,
                            originalSample.RawText,
                            originalSample.CleanedText,
                            originalSample.GroupPngBytes,
                            usedOriginalOcrSample: true);
                    }

                    if (_yoloDetector == null ||
                        targetHwnd == IntPtr.Zero ||
                        !GetWindowRect(targetHwnd, out RECT rect))
                    {
                        return null;
                    }

                    int width = rect.Right - rect.Left;
                    int height = rect.Bottom - rect.Top;
                    if (width <= 0 || height <= 0 ||
                        IsDesktopPixelSize(rect) ||
                        IsSystemWindowClass(targetHwnd))
                    {
                        return null;
                    }

                    token.ThrowIfCancellationRequested();
                    using var windowBitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                    using (var graphics = Graphics.FromImage(windowBitmap))
                    {
                        graphics.CopyFromScreen(
                            rect.Left,
                            rect.Top,
                            0,
                            0,
                            new Size(width, height),
                            CopyPixelOperation.SourceCopy);
                    }

                    List<YoloResult> detections = _yoloDetector.Detect(windowBitmap);
                    CacheDetections(targetHwnd, rect, detections);
                    YoloResult? groupDetection = detections
                        .Where(item => item.LabelName == YoloWindowDetector.Label_GroupName)
                        .OrderByDescending(item => item.Confidence)
                        .FirstOrDefault();

                    string rawOcrText = string.Empty;
                    string cleanedOcrText = string.Empty;
                    Rectangle? capturedGroupBox = null;
                    byte[]? groupPngBytes = null;
                    DateTime capturedAt = DateTime.Now;

                    if (groupDetection != null)
                    {
                        Rectangle validBox = Rectangle.Intersect(
                            new Rectangle(0, 0, width, height),
                            groupDetection.BBox);
                        if (validBox.Width > 0 && validBox.Height > 0)
                        {
                            capturedGroupBox = validBox;
                            using var groupBitmap = new Bitmap(validBox.Width, validBox.Height, PixelFormat.Format32bppArgb);
                            using (var groupGraphics = Graphics.FromImage(groupBitmap))
                            {
                                groupGraphics.DrawImage(
                                    windowBitmap,
                                    new Rectangle(0, 0, validBox.Width, validBox.Height),
                                    validBox,
                                    GraphicsUnit.Pixel);
                            }

                            rawOcrText = await PerformAdaptiveOcrAsync(
                                groupBitmap,
                                $"FailureDebug/{SanitizeDebugToken(storeName ?? "未命名商家", 40)}",
                                token);
                            cleanedOcrText = CleanGroupName(rawOcrText) ?? string.Empty;
                            groupPngBytes = ImageToBytes(groupBitmap);
                        }
                    }

                    return SaveFailedOcrDebugRecord(
                        targetHwnd,
                        isWework,
                        storeName,
                        expectedGroupName,
                        failureStage,
                        retryPhase,
                        capturedAt,
                        width,
                        height,
                        capturedGroupBox,
                        groupDetection?.Confidence,
                        BuildDetectionSummary(detections),
                        rawOcrText,
                        cleanedOcrText,
                        groupPngBytes,
                        usedOriginalOcrSample: false);
                }, token);
            }
            catch (OperationCanceledException)
            {
                return null;
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"⚠️ [失败OCR采集] 保存失败: {ex.Message}");
                return null;
            }
        }

        private void RememberGroupTitleOcrSample(
            IntPtr hwnd,
            bool isWework,
            int windowWidth,
            int windowHeight,
            Rectangle groupBox,
            float yoloConfidence,
            string detectionSummary,
            string? rawText,
            string? cleanedText,
            Bitmap groupBitmap)
        {
            if (!_config.EnableFailedOcrDebugCapture || _disposed)
            {
                return;
            }

            try
            {
                var sample = new GroupTitleOcrSample
                {
                    CapturedAt = DateTime.Now,
                    Hwnd = hwnd,
                    IsWework = isWework,
                    WindowWidth = windowWidth,
                    WindowHeight = windowHeight,
                    GroupBox = groupBox,
                    YoloConfidence = yoloConfidence,
                    DetectionSummary = detectionSummary ?? string.Empty,
                    RawText = rawText ?? string.Empty,
                    CleanedText = cleanedText ?? string.Empty,
                    GroupPngBytes = ImageToBytes(groupBitmap)
                };

                lock (_lastGroupTitleOcrSampleSync)
                {
                    _lastGroupTitleOcrSample = sample;
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"⚠️ [失败OCR采集] 暂存标题样本失败: {ex.Message}");
            }
        }

        private GroupTitleOcrSample? GetRecentGroupTitleOcrSample(
            IntPtr hwnd,
            bool isWework,
            TimeSpan maxAge)
        {
            lock (_lastGroupTitleOcrSampleSync)
            {
                GroupTitleOcrSample? sample = _lastGroupTitleOcrSample;
                if (sample == null ||
                    sample.Hwnd != hwnd ||
                    sample.IsWework != isWework ||
                    DateTime.Now - sample.CapturedAt > maxAge)
                {
                    return null;
                }

                return sample;
            }
        }

        private string SaveFailedOcrDebugRecord(
            IntPtr targetHwnd,
            bool isWework,
            string? storeName,
            string? expectedGroupName,
            string? failureStage,
            string? retryPhase,
            DateTime sampleCapturedAt,
            int windowWidth,
            int windowHeight,
            Rectangle? groupBox,
            float? groupYoloConfidence,
            string detectionSummary,
            string rawOcrText,
            string cleanedOcrText,
            byte[]? groupPngBytes,
            bool usedOriginalOcrSample)
        {
            DateTime recordedAt = DateTime.Now;
            long sequence = Interlocked.Increment(ref _ocrFailureDebugSeq);
            string safeStoreName = SanitizeDebugToken(storeName ?? "未命名商家", 40);
            string safePhase = SanitizeDebugToken(retryPhase ?? "unknown", 16);
            string filePrefix = $"{recordedAt:HHmmss_fff}_{sequence:D5}_{safePhase}_{safeStoreName}";
            string dateDirectory = Path.Combine(FailedOcrDebugDirectory, recordedAt.ToString("yyyyMMdd"));
            Directory.CreateDirectory(dateDirectory);

            string imageFileName = string.Empty;
            if (groupPngBytes is { Length: > 0 })
            {
                imageFileName = filePrefix + ".png";
                File.WriteAllBytes(Path.Combine(dateDirectory, imageFileName), groupPngBytes);
            }

            bool ocrMatchesExpected =
                !string.IsNullOrWhiteSpace(expectedGroupName) &&
                !string.IsNullOrWhiteSpace(cleanedOcrText) &&
                IsFuzzyMatch(expectedGroupName, cleanedOcrText);
            object? groupBoxValue = groupBox.HasValue
                ? new
                {
                    X = groupBox.Value.X,
                    Y = groupBox.Value.Y,
                    Width = groupBox.Value.Width,
                    Height = groupBox.Value.Height
                }
                : null;

            var debugRecord = new
            {
                RecordedAt = recordedAt.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                OcrSampleCapturedAt = sampleCapturedAt.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                UsedOriginalOcrSample = usedOriginalOcrSample,
                StoreName = storeName ?? string.Empty,
                ExpectedGroupName = expectedGroupName ?? string.Empty,
                OcrRawText = rawOcrText ?? string.Empty,
                OcrCleanedText = cleanedOcrText ?? string.Empty,
                OcrMatchesExpected = ocrMatchesExpected,
                App = isWework ? "企业微信" : "微信",
                FailureStage = failureStage ?? string.Empty,
                RetryPhase = retryPhase ?? string.Empty,
                WindowHandle = targetHwnd.ToInt64(),
                WindowSize = new { Width = windowWidth, Height = windowHeight },
                GroupBox = groupBoxValue,
                GroupYoloConfidence = groupYoloConfidence,
                DetectionSummary = detectionSummary ?? string.Empty,
                GroupScreenshot = imageFileName
            };

            string metadataPath = Path.Combine(dateDirectory, filePrefix + ".json");
            string json = JsonSerializer.Serialize(debugRecord, FailedOcrDebugJsonOptions);
            File.WriteAllText(metadataPath, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            _logAction?.Invoke(
                $"🧪 [失败OCR采集] 商家='{storeName}', 预期='{expectedGroupName}', " +
                $"OCR='{cleanedOcrText}', 原始样本={(usedOriginalOcrSample ? "是" : "否")}, " +
                $"阶段={retryPhase}/{failureStage}, 文件={metadataPath}");
            return metadataPath;
        }

        private void ResetOcrEngine(string reason)
        {
            ImageOcr? oldEngine;
            lock (_ocrEngineSync)
            {
                oldEngine = _ocrEngine;
                _ocrEngine = null;
            }

            if (oldEngine == null) return;

            try
            {
                oldEngine.Dispose();
                LogOcrDebug($"OCR_ENGINE_RESET Reason={reason}");
            }
            catch (Exception ex)
            {
                LogOcrDebug($"OCR_ENGINE_RESET_WARN Reason={reason}, Error={ex.Message}");
            }
        }

        private async Task<OcrRunResult> PerformOcrDetailedAsync(
            Bitmap bitmap,
            string? traceTag = null,
            CancellationToken token = default,
            int timeoutMs = OCR_TIMEOUT_MS)
        {
            long requestId = Interlocked.Increment(ref _ocrRequestSeq);
            string tag = string.IsNullOrWhiteSpace(traceTag) ? "Generic" : traceTag.Trim();
            var totalSw = Stopwatch.StartNew();

            if (bitmap == null)
            {
                LogOcrDebug($"#{requestId} FAIL Tag={tag}, Bitmap=null", mirrorToUi: true);
                return new OcrRunResult();
            }

            bool gateEntered = false;
            int activeCount = 0;
            int width = bitmap.Width;
            int height = bitmap.Height;

            try
            {
                var queueSw = Stopwatch.StartNew();
                await _ocrSemaphore.WaitAsync(token);
                gateEntered = true;
                long queueMs = queueSw.ElapsedMilliseconds;

                var encodeSw = Stopwatch.StartNew();
                byte[] bytes = ImageToBytes(bitmap);
                long encodeMs = encodeSw.ElapsedMilliseconds;

                activeCount = Interlocked.Increment(ref _ocrActiveCount);
                LogOcrDebug(
                    $"#{requestId} START Tag={tag}, Size={width}x{height}, Bytes={bytes.Length}, Queue={queueMs}ms, " +
                    $"Encode={encodeMs}ms, Active={activeCount}, Thread={Environment.CurrentManagedThreadId}");

                var tcs = new TaskCompletionSource<OcrRunResult>(TaskCreationOptions.RunContinuationsAsynchronously);
                int callbackEntered = 0;
                int timeoutFlag = 0;
                ImageOcr ocr = GetOrCreateOcrEngine();

                try
                {
                    ocr.Run(bytes, (path, result) =>
                    {
                        long cbDelayMs = totalSw.ElapsedMilliseconds;
                        Interlocked.Exchange(ref callbackEntered, 1);
                        var callbackSw = Stopwatch.StartNew();

                        try
                        {
                            int itemCount = result?.OcrResult?.SingleResult?.Count ?? 0;
                            LogOcrDebug(
                                $"#{requestId} CALLBACK_BEGIN Tag={tag}, Delay={cbDelayMs}ms, " +
                                $"TimedOut={(Volatile.Read(ref timeoutFlag) == 1 ? "Y" : "N")}, Items={itemCount}");

                            var regions = new List<OcrTextRegion>(itemCount);
                            var sb = new StringBuilder();

                            if (result?.OcrResult?.SingleResult != null)
                            {
                                foreach (var item in result.OcrResult.SingleResult)
                                {
                                    if (item == null || string.IsNullOrEmpty(item.SingleStrUtf8)) continue;

                                    sb.Append(item.SingleStrUtf8);
                                    regions.Add(new OcrTextRegion
                                    {
                                        Text = item.SingleStrUtf8,
                                        Left = (int)item.Left,
                                        Top = (int)item.Top,
                                        Right = (int)item.Right,
                                        Bottom = (int)item.Bottom
                                    });
                                }
                            }

                            tcs.TrySetResult(new OcrRunResult
                            {
                                Text = sb.ToString().Trim(),
                                Regions = regions
                            });
                        }
                        catch (Exception ex)
                        {
                            LogOcrDebug($"#{requestId} CALLBACK_EXCEPTION Tag={tag}, Error={ex.Message}", mirrorToUi: true);
                            tcs.TrySetException(ex);
                        }
                        finally
                        {
                            try
                            {
                                if (File.Exists(path)) File.Delete(path);
                            }
                            catch (Exception fileEx)
                            {
                                LogOcrDebug($"#{requestId} CLEANUP_WARN Tag={tag}, DeleteTempFailed={fileEx.Message}");
                            }

                            LogOcrDebug(
                                $"#{requestId} CALLBACK_END Tag={tag}, CbCost={callbackSw.ElapsedMilliseconds}ms, Total={totalSw.ElapsedMilliseconds}ms");
                        }
                    });
                }
                catch
                {
                    ResetOcrEngine($"RunException/{tag}");
                    throw;
                }

                Task timeoutTask = Task.Delay(timeoutMs);
                Task? cancelTask = token.CanBeCanceled ? Task.Delay(Timeout.Infinite, token) : null;
                Task completedTask = cancelTask == null
                    ? await Task.WhenAny(tcs.Task, timeoutTask)
                    : await Task.WhenAny(tcs.Task, timeoutTask, cancelTask);

                if (cancelTask != null && completedTask == cancelTask)
                {
                    Interlocked.Exchange(ref timeoutFlag, 1);
                    LogOcrDebug(
                        $"#{requestId} CANCELED Tag={tag}, Waited={totalSw.ElapsedMilliseconds}ms, " +
                        $"CallbackEntered={(Volatile.Read(ref callbackEntered) == 1 ? "Y" : "N")}",
                        mirrorToUi: true);
                    ResetOcrEngine($"Canceled/{tag}");
                    return new OcrRunResult { Canceled = true };
                }

                if (completedTask == timeoutTask)
                {
                    Interlocked.Exchange(ref timeoutFlag, 1);
                    LogOcrDebug(
                    $"#{requestId} TIMEOUT Tag={tag}, Waited={totalSw.ElapsedMilliseconds}ms, " +
                        $"CallbackEntered={(Volatile.Read(ref callbackEntered) == 1 ? "Y" : "N")}, Size={width}x{height}, Bytes={bytes.Length}",
                        mirrorToUi: true);
                    ResetOcrEngine($"Timeout/{tag}");
                    return new OcrRunResult { TimedOut = true };
                }

                OcrRunResult ocrResult = await tcs.Task;
                LogOcrDebug(
                    $"#{requestId} DONE Tag={tag}, Total={totalSw.ElapsedMilliseconds}ms, TextLen={ocrResult.Text.Length}, Regions={ocrResult.Regions.Count}");
                return ocrResult;
            }
            catch (OperationCanceledException)
            {
                return new OcrRunResult { Canceled = true };
            }
            finally
            {
                if (activeCount > 0)
                {
                    int remain = Interlocked.Decrement(ref _ocrActiveCount);
                    LogOcrDebug($"#{requestId} RELEASE Tag={tag}, Active={remain}, Total={totalSw.ElapsedMilliseconds}ms");
                }

                if (gateEntered) _ocrSemaphore.Release();
            }
        }

        private async Task<string> PerformOcrAsync(Bitmap bitmap, string? traceTag = null, CancellationToken token = default)
        {
            OcrRunResult result = await PerformOcrDetailedAsync(bitmap, traceTag, token);
            if (result.Canceled) return "OCR识别取消";
            if (result.TimedOut) return "OCR识别超时";
            return string.IsNullOrWhiteSpace(result.Text) ? "未识别到文字" : result.Text;
        }

        private async Task<string> PerformAdaptiveOcrAsync(Bitmap bitmap, string traceTag, CancellationToken token = default)
        {
            int fastScale = GetRecommendedOcrScale(bitmap);
            using (var processed = PreprocessForOcr(bitmap, fastScale, highQuality: false))
            {
                string text = await PerformOcrAsync(processed, $"{traceTag}/S{fastScale}", token);
                if (!ShouldRetryOcr(text) || fastScale >= 3 || token.IsCancellationRequested)
                {
                    return text;
                }
            }

            using (var fallback = PreprocessForOcr(bitmap, 3, highQuality: true))
            {
                return await PerformOcrAsync(fallback, $"{traceTag}/S3Fallback", token);
            }
        }

        private static bool ShouldRetryOcr(string? text)
        {
            return string.IsNullOrWhiteSpace(text) || string.Equals(text, "未识别到文字", StringComparison.Ordinal);
        }

        private static int GetRecommendedOcrScale(Bitmap bitmap)
        {
            if (bitmap.Height >= 96) return 1;
            if (bitmap.Height >= 36) return 2;
            return 3;
        }

        #region 辅助方法
        private string GetAppIdentifier(string processName)
        {
            if ("WeChat".Equals(processName, StringComparison.OrdinalIgnoreCase)) return "微信";
            if ("WXWork".Equals(processName, StringComparison.OrdinalIgnoreCase)) return "企业微信";
            return "未知应用";
        }

        private int GetRightCropAmount(string appIdentifier)
        {
            switch (appIdentifier)
            {
                case "企业微信": return _config.WeWorkRightCrop;
                case "微信": return _config.WeChatRightCrop;
                default: return DEFAULT_RIGHT_CROP;
            }
        }

        private Bitmap ScaleImage(Bitmap original, int scaleFactor)
        {
            int newWidth = original.Width * scaleFactor;
            int newHeight = original.Height * scaleFactor;
            var scaled = new Bitmap(newWidth, newHeight, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(scaled))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(original, new Rectangle(0, 0, newWidth, newHeight));
            }
            return scaled;
        }

        private byte[] ImageToBytes(Bitmap bitmap)
        {
            using (var ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }

        private string CleanGroupName(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return rawText;
            // 移除末尾的 (...) 或 (...)
            string pattern = @"[（\(]\s*\d+\s*[）\)]\s*$";
            return Regex.Replace(rawText, pattern, "").Trim();
        }
        #endregion

        /// <summary>
        /// ✅ [通用版] 验证搜索结果列表 (带网络搜索排除逻辑)
        /// 支持：微信 (53,93) 和 企业微信 (78,90)
        /// 🚫 改版：使用 YOLO 检测 "搜索群聊" 或 "最近搜索群聊" 标签
        /// </summary>
        /// <summary>
        /// ✅ [通用版] 验证搜索结果列表 (带网络搜索排除逻辑)
        /// 支持：微信 (53,93) 和 企业微信 (78,90)
        /// 🚫 改版：使用 YOLO 检测 "搜索群聊" 或 "最近搜索群聊" 标签
        /// </summary>
        public async Task<bool> CheckSearchResultAsync(IntPtr targetHwnd, string expectedText, bool isWework)
        {
            var result = await FindAndVerifySearchResultAsync(targetHwnd, expectedText, isWework);
            return result.success;
        }

        /// <summary>
        /// ✅ [合并版] 验证搜索结果并返回点击坐标 (合并 StepA 和 StepB)
        /// 避免 StepA 成功但 StepB 重新截图导致失败的问题
        /// </summary>
        /// <summary>
        /// ✅ [合并版] 验证搜索结果并返回点击坐标 (合并 StepA 和 StepB)
        /// 避免 StepA 成功但 StepB 重新截图导致失败的问题
        /// 🔄 [增强] 引入多侦测 (Multi-Frame Detection) 机制，过滤闪烁误报
        /// </summary>
        public async Task<(bool success, Point? clickPoint, Rectangle? matchedScreenBBox)> FindAndVerifySearchResultAsync(IntPtr targetHwnd, string expectedText, bool isWework)
        {
            try
            {
                string appName = isWework ? "企业微信" : "微信";
                string expected = expectedText?.Trim() ?? string.Empty;
                string expectedToken = SanitizeDebugToken(expected, 24);
                System.Diagnostics.Debug.WriteLine(
                    $"[SearchOCR] 开始搜索结果验证: App={appName}, Hwnd={targetHwnd}, Expected='{expected}', " +
                    $"ConfTh={SEARCH_YOLO_CONF_THRESHOLD:F2}, IouTh={SEARCH_YOLO_IOU_THRESHOLD:F2}");

                if (_yoloDetector == null)
                {
                    _logAction?.Invoke("❌ YOLO Detector 未初始化");
                    return (false, null, null);
                }

                // 定义多帧结果容器
                var frameResults = new List<List<(YoloResult Result, Rectangle ScreenBBox, Rectangle LocalBBox)>>();
                Bitmap finalBitmap = null;
                int finalScreenX = 0, finalScreenY = 0;
                
                const int maxStableFrameCount = 3;
                const float singleFrameFastConfidence = 0.65f;

                for (int i = 0; i < maxStableFrameCount; i++)
                {
                    if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) 
                    {
                        if (finalBitmap != null) finalBitmap.Dispose();
                        return (false, null, null);
                    }

                    // 🛡️ [防御] 防止截取到桌面
                    if (IsDesktopPixelSize(rect) || IsSystemWindowClass(targetHwnd))
                    {
                        System.Diagnostics.Debug.WriteLine($"⚠️ [FindAndVerify] 拦截到桌面/系统窗口截取请求: {targetHwnd}");
                        if (finalBitmap != null) finalBitmap.Dispose();
                        return (false, null, null);
                    }

                    // 截取窗口
                    int width = rect.Right - rect.Left;
                    int height = rect.Bottom - rect.Top;
                    int screenX = rect.Left;
                    int screenY = rect.Top;

                    if (width <= 0 || (rect.Bottom - rect.Top) < 100) 
                    {
                         if (finalBitmap != null) finalBitmap.Dispose();
                         return (false, null, null);
                    }

                    var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // YOLO 检测
                    var results = _yoloDetector.Detect(bitmap, SEARCH_YOLO_CONF_THRESHOLD, SEARCH_YOLO_IOU_THRESHOLD);
                    var currentFrameTargets = results
                        .Where(r => r.LabelName == YoloWindowDetector.Label_SearchGroup || 
                                    r.LabelName == YoloWindowDetector.Label_RecentGroup)
                        .Select(r => (r, new Rectangle(screenX + r.BBox.X, screenY + r.BBox.Y, r.BBox.Width, r.BBox.Height), r.BBox))
                        .ToList();

                    string framePrefix = $"SearchVerify_{(isWework ? "WeWork" : "WeChat")}_F{i + 1}_{expectedToken}";
                    // [已禁用] Debug_Yolo 调试图保存
                    // string frameRawPath = SaveDebugRawImage(bitmap, framePrefix);
                    // string frameAnnPath = SaveDebugAnnotatedImage(bitmap, results, framePrefix);
                    string frameSummary = BuildDetectionSummary(currentFrameTargets.Select(x => x.r), 10);
                    System.Diagnostics.Debug.WriteLine($"[SearchOCR] 帧{i + 1}/{maxStableFrameCount}: 候选={currentFrameTargets.Count}, 明细={frameSummary}");
                    // [已禁用] Debug_Yolo 调试图日志
                    // if (!string.IsNullOrEmpty(frameRawPath) || !string.IsNullOrEmpty(frameAnnPath))
                    // {
                    //     System.Diagnostics.Debug.WriteLine($"[SearchOCR] 帧{i + 1}调试图: Raw='{frameRawPath}', Ann='{frameAnnPath}'");
                    // }
                    
                    frameResults.Add(currentFrameTargets);

                    bool singleFrameFastPath = i == 0 &&
                                               currentFrameTargets.Count == 1 &&
                                               currentFrameTargets[0].r.Confidence >= singleFrameFastConfidence;
                    bool stableAfterMultipleFrames = i > 0 && currentFrameTargets.Any(candidate =>
                        frameResults.Take(frameResults.Count - 1).Any(previousFrame =>
                            previousFrame.Any(previous => IsSameDetection(
                                candidate.r,
                                candidate.Item2,
                                previous.Result,
                                previous.ScreenBBox))));
                    bool keepCurrentFrame = singleFrameFastPath || stableAfterMultipleFrames || i == maxStableFrameCount - 1;

                    // 高置信度单目标走单帧；低置信度目标在两帧稳定后提前结束；仅不稳定场景取第三帧。
                    if (keepCurrentFrame)
                    {
                        finalBitmap = bitmap;
                        finalScreenX = screenX;
                        finalScreenY = screenY;
                        System.Diagnostics.Debug.WriteLine(
                            $"[SearchOCR] 自适应取帧结束: Captured={frameResults.Count}, " +
                            $"FastPath={singleFrameFastPath}, Stable={stableAfterMultipleFrames}");
                        break;
                    }
                    else
                    {
                        bitmap.Dispose();
                    }
                }

                using (finalBitmap) // 确保 verify 完后释放
                {
                    // --- 稳定性分析 ---
                    var lastFrameTargets = frameResults.Last();
                    if (lastFrameTargets.Count == 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"❌ [FindAndVerify] 最后一帧未检测到任何结果，放弃");
                        return (false, null, null);
                    }

                    (YoloResult Result, Rectangle ScreenBBox, Rectangle LocalBBox)? bestStableTarget = null;
                    float bestConfidence = -1f;

                    foreach (var candidate in lastFrameTargets)
                    {
                        int appearanceCount = 1; // 最后一帧已出现
                        
                        // 在前几帧中查找匹配 (基于 IoU 或 中心距离)
                        for (int i = 0; i < frameResults.Count - 1; i++)
                        {
                            var prevTargets = frameResults[i];
                            bool matchFound = prevTargets.Any(prev => IsSameDetection(
                                candidate.Result,
                                candidate.ScreenBBox,
                                prev.Result,
                                prev.ScreenBBox));

                            if (matchFound) appearanceCount++;
                        }

                        System.Diagnostics.Debug.WriteLine(
                            $"[SearchOCR] 稳定性候选: Label={candidate.Result.LabelName}, Conf={candidate.Result.Confidence:F2}, " +
                            $"ScreenBBox={BuildBboxText(candidate.ScreenBBox)}, Appear={appearanceCount}/{frameResults.Count}");

                        int requiredAppearanceCount = frameResults.Count == 1 ? 1 : 2;
                        if (appearanceCount >= requiredAppearanceCount)
                        {
                            if (candidate.Result.Confidence > bestConfidence)
                            {
                                bestConfidence = candidate.Result.Confidence;
                                bestStableTarget = candidate;
                            }
                        }
                    }

                    if (bestStableTarget == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"❌ [FindAndVerify] 检测到目标但不稳定 (未连续出现)，Expected='{expected}'，放弃");
                        return (false, null, null);
                    }

                    var target = bestStableTarget.Value;
                    System.Diagnostics.Debug.WriteLine($"✅ [FindAndVerify] 捕获稳定目标: {target.Result.LabelName} (Conf:{target.Result.Confidence:F2})");

                    // 4. OCR 验证
                    var bbox = target.LocalBBox;
                    if (bbox.Width <= 0 || bbox.Height <= 0) return (false, null, null);

                    using (var crop = new Bitmap(bbox.Width, bbox.Height))
                    using (var g = Graphics.FromImage(crop))
                    {
                        g.DrawImage(finalBitmap, new Rectangle(0, 0, bbox.Width, bbox.Height), bbox, GraphicsUnit.Pixel);
                        
                        string ocrText = await PerformAdaptiveOcrAsync(crop, $"SearchVerify/{expectedToken}");
                        bool match = IsFuzzyMatch(expectedText, ocrText);
                        System.Diagnostics.Debug.WriteLine(
                            $"🔍 [FindAndVerify] OCR候选: Label={target.Result.LabelName}, Conf={target.Result.Confidence:F2}, " +
                            $"BBox={BuildBboxText(bbox)}, OCR='{ocrText}', Expected='{expected}', Match={match}");

                        if (match)
                        {
                            System.Diagnostics.Debug.WriteLine($"✅ [FindAndVerify] 文本匹配成功: '{expectedText}'");
                                
                            int marginX = (int)(bbox.Width * 0.2);
                            int marginY = (int)(bbox.Height * 0.2);
                            if (marginX < 2) marginX = 2;
                            if (marginY < 2) marginY = 2;

                            int safeW = bbox.Width - 2 * marginX;
                            int safeH = bbox.Height - 2 * marginY;
                                
                            if (safeW <= 0) safeW = 1;
                            if (safeH <= 0) safeH = 1;

                            Random rnd = new Random();
                            int offsetX = marginX + rnd.Next(0, safeW);
                            int offsetY = marginY + rnd.Next(0, safeH);

                            int clickX = finalScreenX + bbox.X + offsetX;
                            int clickY = finalScreenY + bbox.Y + offsetY;

                            var matchedScreenBBox = new Rectangle(
                                finalScreenX + bbox.X,
                                finalScreenY + bbox.Y,
                                bbox.Width,
                                bbox.Height
                            );

                            return (true, new Point(clickX, clickY), matchedScreenBBox);
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"❌ [FindAndVerify] 稳定目标文本不匹配: Expected='{expected}', App={appName}");
                    return (false, null, null);
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 搜索验证合并版出错: {ex.Message}");
                return (false, null, null);
            }
        }

        /// <summary>
        /// 点击前二次校验：确认当前点击点仍落在 YOLO 识别的目标群聊项内，并通过 OCR 文本匹配
        /// </summary>
        public async Task<bool> ValidateSearchResultPointAsync(IntPtr targetHwnd, string expectedText, bool isWework, Point screenPoint, int tolerancePixels = 6)
        {
            try
            {
                if (_yoloDetector == null)
                {
                    _logAction?.Invoke("❌ YOLO Detector 未初始化");
                    return false;
                }

                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) return false;

                // 🛡️ [防御] 防止截取到桌面
                if (IsDesktopPixelSize(rect) || IsSystemWindowClass(targetHwnd))
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ [ValidatePoint] 拦截到桌面/系统窗口截取请求: {targetHwnd}");
                    return false;
                }

                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                int screenX = rect.Left;
                int screenY = rect.Top;

                if (width <= 0 || height < 100) return false;

                int localX = screenPoint.X - screenX;
                int localY = screenPoint.Y - screenY;
                string expected = expectedText?.Trim() ?? string.Empty;
                string expectedToken = SanitizeDebugToken(expected, 24);

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    var targets = _yoloDetector.Detect(bitmap, SEARCH_YOLO_CONF_THRESHOLD, SEARCH_YOLO_IOU_THRESHOLD)
                        .Where(r => r.LabelName == YoloWindowDetector.Label_SearchGroup ||
                                    r.LabelName == YoloWindowDetector.Label_RecentGroup)
                        .OrderByDescending(r => r.Confidence)
                        .ToList();

                    string pointPrefix = $"ValidatePoint_{(isWework ? "WeWork" : "WeChat")}_{expectedToken}";
                    // [已禁用] Debug_Yolo 调试图保存
                    // string pointRawPath = SaveDebugRawImage(bitmap, pointPrefix);
                    // string pointAnnPath = SaveDebugAnnotatedImage(bitmap, targets, pointPrefix);
                    string pointSummary = BuildDetectionSummary(targets, 10);
                    System.Diagnostics.Debug.WriteLine(
                        $"[ValidatePoint] App={(isWework ? "企业微信" : "微信")}, Expected='{expected}', Point=({screenPoint.X},{screenPoint.Y}), " +
                        $"候选数={targets.Count}, 明细={pointSummary}, ConfTh={SEARCH_YOLO_CONF_THRESHOLD:F2}, IouTh={SEARCH_YOLO_IOU_THRESHOLD:F2}");
                    // [已禁用] Debug_Yolo 调试图日志
                    // if (!string.IsNullOrEmpty(pointRawPath) || !string.IsNullOrEmpty(pointAnnPath))
                    // {
                    //     System.Diagnostics.Debug.WriteLine($"[ValidatePoint] 调试图: Raw='{pointRawPath}', Ann='{pointAnnPath}'");
                    // }

                    if (targets.Count == 0)
                    {
                        System.Diagnostics.Debug.WriteLine("❌ [ValidatePoint] YOLO 未检测到任何搜索结果标签");
                        return false;
                    }

                    var pointTargets = targets.Where(t =>
                    {
                        var bbox = t.BBox;
                        int left = bbox.X - tolerancePixels;
                        int top = bbox.Y - tolerancePixels;
                        int right = bbox.Right + tolerancePixels;
                        int bottom = bbox.Bottom + tolerancePixels;
                        return localX >= left && localX <= right && localY >= top && localY <= bottom;
                    }).ToList();

                    if (pointTargets.Count == 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"❌ [ValidatePoint] 点击点({screenPoint.X},{screenPoint.Y}) 不在任何候选群聊框内");
                        return false;
                    }

                    foreach (var target in pointTargets)
                    {
                        var bbox = target.BBox;
                        if (bbox.Width <= 0 || bbox.Height <= 0) continue;

                        using (var crop = new Bitmap(bbox.Width, bbox.Height))
                        using (var g = Graphics.FromImage(crop))
                        {
                            g.DrawImage(bitmap, new Rectangle(0, 0, bbox.Width, bbox.Height), bbox, GraphicsUnit.Pixel);
                            string ocrText = await PerformAdaptiveOcrAsync(crop, $"ValidatePoint/{expectedToken}");
                            bool match = IsFuzzyMatch(expectedText, ocrText);
                            System.Diagnostics.Debug.WriteLine($"🔍 [ValidatePoint] 候选 OCR:'{ocrText}' -> {(match ? "✅ 匹配" : "❌ 不匹配")}");
                            if (match)
                            {
                                System.Diagnostics.Debug.WriteLine($"✅ [ValidatePoint] 点击点二次校验通过: ({screenPoint.X},{screenPoint.Y})");
                                return true;
                            }
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"❌ [ValidatePoint] 点击点命中目标框，但 OCR 不匹配期望文本 '{expectedText}'");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 点击前二次校验失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ✅ [新增] 在指定区域查找关键词的坐标 (用于定位 "群聊" 等动态位置)
        /// </summary>
        /// <returns>关键词中心点在屏幕上的坐标 (如果是 null 则未找到)</returns>
        public async Task<Point?> FindKeywordPositionAsync(IntPtr targetHwnd, string keyword, int[] relativeSearchArea = null)
        {
            try
            {
                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) return null;

                // 默认搜索区域：整个窗口的左侧 (例如前 400px，高度覆盖全部)
                // 默认搜索区域：整个窗口的左侧 (例如前 400px，高度覆盖全部)
                int relX = relativeSearchArea?[0] ?? 0;
                int relY = relativeSearchArea?[1] ?? 60; // 避开顶栏
                int width = relativeSearchArea?[2] ?? 400; // 搜索栏通常在左侧
                int height = relativeSearchArea?[3] ?? (rect.Bottom - rect.Top - 60);

                // 越界保护
                if (width <= 0) width = 200;
                if (height <= 0) height = 200;

                // 🛡️ [防御] 防止截取到桌面
                if (IsDesktopPixelSize(rect) || IsSystemWindowClass(targetHwnd))
                {
                    // FindKeyword 比较特殊，有时候也许是在桌面找东西？但本项目应该都是在微信窗口找
                    _logAction?.Invoke($"⚠️ [FindKeyword] 拦截到桌面/系统窗口截取请求: {targetHwnd}");
                    return null;
                }

                int screenX = rect.Left + relX;
                int screenY = rect.Top + relY;

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var g = Graphics.FromImage(bitmap))
                    {
                        g.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // OCR
                    using (var scaled = ScaleImage(bitmap, 2)) // 2倍放大通常够了
                    {
                        OcrRunResult ocrResult = await PerformOcrDetailedAsync(
                            scaled,
                            $"FindKeyword/{keyword}",
                            timeoutMs: 5000);

                        OcrTextRegion? match = ocrResult.Regions.FirstOrDefault(item => IsFuzzyMatch(keyword, item.Text));
                        if (match == null) return null;

                        int finalX = screenX + (match.Left + match.Right) / 4;
                        int finalY = screenY + (match.Top + match.Bottom) / 4;
                        return new Point(finalX, finalY);
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 FindKeywordPositionAsync 异常: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// ✅ [核心功能] 查找弹窗上的文字坐标 (如 "使用原文件")
        /// 🚫 改版：截取窗口大部分区域，由 YOLO 在大图中检测
        /// </summary>
        public async Task<Point?> FindPopupTextPositionAsync(IntPtr targetHwnd, string keyword)
        {
            try
            {
                if (_yoloDetector == null) return null;
                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) return null;

                // 🚀 [改版] 截取整个窗口 (Full Window)
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                int screenX = rect.Left;
                int screenY = rect.Top;

                if (width <= 0 || height <= 0) return null;

                // 🛡️ [防御] 防止截取到桌面
                if (IsDesktopPixelSize(rect) || IsSystemWindowClass(targetHwnd))
                {
                    _logAction?.Invoke($"⚠️ [FindPopup] 拦截到桌面/系统窗口截取请求: {targetHwnd}");
                    return null;
                }

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var g = Graphics.FromImage(bitmap))
                    {
                        g.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // 1. YOLO 识别
                    var yoloResults = _yoloDetector.Detect(bitmap);
                    
                    // 2. 保存调试图 [已注释]
                    // string debugDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Debug_Yolo", DateTime.Now.ToString("yyyyMMdd"));
                    // Directory.CreateDirectory(debugDir);
                    // string debugFile = Path.Combine(debugDir, $"Search_{DateTime.Now:HHmmss_fff}.png");
                    // _yoloDetector.InferenceWrapper.SaveDebugImage(bitmap, yoloResults, debugFile);
                    // _logAction?.Invoke($"🖼️ YOLO 识别图已存: {debugFile}");

                    // 3. 筛选 "在线文档"
                    var popupRect = yoloResults
                        .Where(r => r.LabelName == YoloWindowDetector.Label_OnlineDoc)
                        .OrderByDescending(r => r.Confidence)
                        .Select(r => (Rectangle?)r.BBox)
                        .FirstOrDefault();
                    
                    if (popupRect.HasValue)
                    {
                        var rectVal = popupRect.Value;
                        _logAction?.Invoke($"✅ YOLO 定位到弹窗区域: {rectVal}");

                        // 裁剪出弹窗区域
                        using (var crop = bitmap.Clone(rectVal, bitmap.PixelFormat))
                        using (var scaled = ScaleImage(crop, 2)) // 2倍放大 OCR
                        {
                            OcrRunResult ocrResult = await PerformOcrDetailedAsync(
                                scaled,
                                $"FindPopup/{keyword}",
                                timeoutMs: 3000);
                            OcrTextRegion? match = ocrResult.Regions.FirstOrDefault(item => IsFuzzyMatch(keyword, item.Text));
                            if (match != null)
                            {
                                int finalX = screenX + rectVal.X + (match.Left + match.Right) / 4;
                                int finalY = screenY + rectVal.Y + (match.Top + match.Bottom) / 4;
                                _logAction?.Invoke($"🎯 OCR 精确定位 '{keyword}': ({finalX}, {finalY})");
                                return new Point(finalX, finalY);
                            }
                        }
                        
                        _logAction?.Invoke("🚫 未在弹窗中识别到 '使用原文件'，放弃点击。");
                        return null;
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 FindPopupTextPositionAsync 异常: {ex.Message}");
                return null;
            }
        }




        /// <summary>
        /// ✅ [优化版] 动态定位群聊点击位置
        /// 🚫 改版：完全移除固定坐标配置，改为自适应检测
        /// </summary>
        public async Task<Point?> FindGroupChatClickPositionAsync(IntPtr targetHwnd, string targetGroupName = null, bool isWework = false)
        {
            try
            {
                if (_yoloDetector == null) return null;
                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) return null;

                // 🚀 [改版] 截取整个窗口 (Full Window)
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                int screenX = rect.Left;
                int screenY = rect.Top;

                if (width <= 0 || height <= 0) return null;

                // 🛡️ [防御] 防止截取到桌面
                if (IsDesktopPixelSize(rect) || IsSystemWindowClass(targetHwnd))
                {
                    _logAction?.Invoke($"⚠️ [FindGroup] 拦截到桌面/系统窗口截取请求: {targetHwnd}");
                    return null;
                }

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // 1. YOLO 第一轮识别
                    var yoloResults = _yoloDetector.Detect(bitmap);
                    
                    // 2. 保存调试图 [已注释]
                    // string debugDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Debug_Yolo", DateTime.Now.ToString("yyyyMMdd"));
                    // Directory.CreateDirectory(debugDir);
                    // string debugFile = Path.Combine(debugDir, $"GroupClick_{DateTime.Now:HHmmss_fff}.png");
                    // _yoloDetector.InferenceWrapper.SaveDebugImage(bitmap, yoloResults, debugFile);
                    // _logAction?.Invoke($"🖼️ YOLO 群聊识别图已存: {debugFile}");

                    // 3. 筛选目标：搜索群聊 或 最近搜索群聊
                    var targets = yoloResults
                        .Where(r => r.LabelName == YoloWindowDetector.Label_SearchGroup || 
                                    r.LabelName == YoloWindowDetector.Label_RecentGroup)
                        .OrderByDescending(r => r.Confidence)
                        .ToList();

                    if (targets.Count == 0)
                    {
                        _logAction?.Invoke(
                            $"❌ YOLO 未在当前区域检测到群聊列表项 (ConfTh={SEARCH_YOLO_CONF_THRESHOLD:F2}, IouTh={SEARCH_YOLO_IOU_THRESHOLD:F2})");
                        return null;
                    }

                    // 4. 匹配逻辑
                    if (!string.IsNullOrEmpty(targetGroupName))
                    {
                        _logAction?.Invoke($"🔍 目标群聊: '{targetGroupName}'，尝试 OCR 匹配...");
                        foreach (var target in targets)
                        {
                            var bbox = target.BBox;
                            using (var crop = new Bitmap(bbox.Width, bbox.Height))
                            using (var g = Graphics.FromImage(crop))
                            {
                                g.DrawImage(bitmap, new Rectangle(0, 0, bbox.Width, bbox.Height), bbox, GraphicsUnit.Pixel);
                                string ocrText = await PerformAdaptiveOcrAsync(crop, $"FindGroup/{targetGroupName ?? "Auto"}");
                                bool match = IsFuzzyMatch(targetGroupName, ocrText);

                                _logAction?.Invoke($"   - [{target.LabelName}] OCR:'{ocrText}' -> {(match ? "✅ 匹配" : "❌ 忽略")}");

                                if (match)
                                {
                                    int mX = Math.Max(2, (int)(bbox.Width * 0.2));
                                    int mY = Math.Max(2, (int)(bbox.Height * 0.2));
                                    int sW = Math.Max(1, bbox.Width - 2 * mX);
                                    int sH = Math.Max(1, bbox.Height - 2 * mY);

                                    return new Point(
                                        screenX + bbox.X + mX + Random.Shared.Next(0, sW),
                                        screenY + bbox.Y + mY + Random.Shared.Next(0, sH));
                                }
                            }
                        }
                        return null;
                    }
                    else
                    {
                        // 无指定名字，返回置信度最高的
                        var best = targets.FirstOrDefault(t => t.LabelName == YoloWindowDetector.Label_SearchGroup) ?? targets.First();
                        var safeBBox = best.BBox;
                        int marginX = (int)(safeBBox.Width * 0.2);
                        int marginY = (int)(safeBBox.Height * 0.2);
                        if (marginX < 2) marginX = 2;
                        if (marginY < 2) marginY = 2;

                        int safeW = safeBBox.Width - 2 * marginX;
                        int safeH = safeBBox.Height - 2 * marginY;
                        if (safeW <= 0) safeW = 1;
                        if (safeH <= 0) safeH = 1;

                        Random rnd = new Random();
                        int cx = screenX + safeBBox.X + marginX + rnd.Next(0, safeW);
                        int cy = screenY + safeBBox.Y + marginY + rnd.Next(0, safeH);

                        _logAction?.Invoke($"✅ 自动锁定最高置信度目标: {best.LabelName} ({best.Confidence:P0}) -> 随机点 ({cx},{cy})");
                        return new Point(cx, cy);
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 FindGroupChatClickPositionAsync 异常: {ex.Message}");
                return null;
            }
        }



     


        /// <summary>
        /// 🔍 结构化查找：找到锚点关键词，然后返回它正下方的第一个文本项的坐标
        /// </summary>
        public async Task<System.Drawing.Point?> FindItemBelowKeywordAsync(IntPtr targetHwnd, string anchorKeyword, int[] relativeSearchArea = null)
        {
            try
            {
                if (!GetWindowRect(targetHwnd, out RECT rect)) return null;

                int relX = relativeSearchArea != null ? relativeSearchArea[0] : 0;
                int relY = relativeSearchArea != null ? relativeSearchArea[1] : 0;
                int width = relativeSearchArea != null ? relativeSearchArea[2] : (rect.Right - rect.Left);
                int height = relativeSearchArea != null ? relativeSearchArea[3] : (rect.Bottom - rect.Top);

                int screenX = rect.Left + relX;
                int screenY = rect.Top + relY;

                if (width <= 0 || height <= 0) return null;

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var g = Graphics.FromImage(bitmap)) { g.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy); }

                    OcrRunResult ocrResult = await PerformOcrDetailedAsync(
                        bitmap,
                        $"FindStructure/{anchorKeyword}",
                        timeoutMs: 2000);
                    OcrTextRegion? anchorItem = ocrResult.Regions.FirstOrDefault(item =>
                        item.Text.Contains(anchorKeyword, StringComparison.OrdinalIgnoreCase));
                    if (anchorItem == null) return null;

                    OcrTextRegion? targetItem = ocrResult.Regions
                        .Where(item => !ReferenceEquals(item, anchorItem))
                        .Where(item => item.Top > anchorItem.Bottom && item.Top < anchorItem.Bottom + 150)
                        .OrderBy(item => item.Top - anchorItem.Bottom)
                        .FirstOrDefault();
                    if (targetItem == null) return null;

                    int centerX = screenX + (targetItem.Left + targetItem.Right) / 2;
                    int centerY = screenY + (targetItem.Top + targetItem.Bottom) / 2;
                    _logAction?.Invoke($"✅ [FindStructure] Found target below '{anchorKeyword}' -> '{targetItem.Text}'");
                    return new System.Drawing.Point(centerX, centerY);
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 结构化查找失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 在源图像中查找目标图像第一次出现的位置 (安全版，从上往下搜)
        /// </summary>
        private Rectangle? FindFirstImageOccurrence(Bitmap source, Bitmap template, int tolerance = 30)
        {
            BitmapData sourceData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            BitmapData templateData = template.LockBits(new Rectangle(0, 0, template.Width, template.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

            try
            {
                int sourceBytesCount = sourceData.Stride * source.Height;
                byte[] sourceBytes = new byte[sourceBytesCount];
                Marshal.Copy(sourceData.Scan0, sourceBytes, 0, sourceBytesCount);

                int templateBytesCount = templateData.Stride * template.Height;
                byte[] templateBytes = new byte[templateBytesCount];
                Marshal.Copy(templateData.Scan0, templateBytes, 0, templateBytesCount);

                int tWidth = template.Width;
                int tHeight = template.Height;
                int sWidth = source.Width;
                int sHeight = source.Height;
                int sStride = sourceData.Stride;
                int tStride = templateData.Stride;

                // 从上往下搜索 (即 Y 从 0 开始增加)
                for (int y = 0; y <= sHeight - tHeight; y++)
                {
                    for (int x = 0; x <= sWidth - tWidth; x++)
                    {
                        if (IsMatchSafe(sourceBytes, templateBytes, x, y, tWidth, tHeight, sStride, tStride, tolerance))
                        {
                            return new Rectangle(x, y, tWidth, tHeight);
                        }
                    }
                }
            }
            finally
            {
                source.UnlockBits(sourceData);
                template.UnlockBits(templateData);
            }

            return null;
        }

        /// <summary>
        /// 在源图像中查找目标图像最后一次出现的位置 (安全版，无需 unsafe)
        /// </summary>
        private System.Drawing.Rectangle? FindLastImageOccurrence(Bitmap source, Bitmap template, int tolerance = 30)
        {
            if (source == null || template == null || source.Width < template.Width || source.Height < template.Height)
                return null;

            Rectangle sourceRect = new Rectangle(0, 0, source.Width, source.Height);
            Rectangle templateRect = new Rectangle(0, 0, template.Width, template.Height);
            
            BitmapData sourceData = source.LockBits(sourceRect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            BitmapData templateData = template.LockBits(templateRect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

            try
            {
                int sourceBytesCount = sourceData.Stride * source.Height;
                byte[] sourceBytes = new byte[sourceBytesCount];
                Marshal.Copy(sourceData.Scan0, sourceBytes, 0, sourceBytesCount);

                int templateBytesCount = templateData.Stride * template.Height;
                byte[] templateBytes = new byte[templateBytesCount];
                Marshal.Copy(templateData.Scan0, templateBytes, 0, templateBytesCount);

                int tWidth = template.Width;
                int tHeight = template.Height;
                int sWidth = source.Width;
                int sHeight = source.Height;
                int sStride = sourceData.Stride;
                int tStride = templateData.Stride;

                // 从下往上搜索
                for (int y = sHeight - tHeight; y >= 0; y--)
                {
                    for (int x = sWidth - tWidth; x >= 0; x--)
                    {
                        if (IsMatchSafe(sourceBytes, templateBytes, x, y, tWidth, tHeight, sStride, tStride, tolerance))
                        {
                            return new Rectangle(x, y, tWidth, tHeight);
                        }
                    }
                }
            }
            finally
            {
                source.UnlockBits(sourceData);
                template.UnlockBits(templateData);
            }

            return null;
        }

        private bool IsMatchSafe(byte[] sourceBytes, byte[] templateBytes, int startX, int startY, 
            int tWidth, int tHeight, int sStride, int tStride, int tolerance)
        {
            // 检查中心点
            int centerX = tWidth / 2;
            int centerY = tHeight / 2;
            if (!PixelMatchSafe(sourceBytes, templateBytes, startX + centerX, startY + centerY, centerX, centerY, sStride, tStride, tolerance)) return false;

            // 检查四角
            if (!PixelMatchSafe(sourceBytes, templateBytes, startX, startY, 0, 0, sStride, tStride, tolerance)) return false;
            if (!PixelMatchSafe(sourceBytes, templateBytes, startX + tWidth - 1, startY, tWidth - 1, 0, sStride, tStride, tolerance)) return false;

            // 全像素检查
            for (int y = 0; y < tHeight; y++)
            {
                int sRowIdx = (startY + y) * sStride + startX * 4;
                int tRowIdx = y * tStride;

                for (int x = 0; x < tWidth; x++)
                {
                    int sIdx = sRowIdx + x * 4;
                    int tIdx = tRowIdx + x * 4;

                    byte ta = templateBytes[tIdx + 3];
                    if (ta < 10) continue; // 透明跳过

                    byte sb = sourceBytes[sIdx];
                    byte sg = sourceBytes[sIdx + 1];
                    byte sr = sourceBytes[sIdx + 2];

                    byte tb = templateBytes[tIdx];
                    byte tg = templateBytes[tIdx + 1];
                    byte tr = templateBytes[tIdx + 2];

                    if (Math.Abs(sb - tb) > tolerance ||
                        Math.Abs(sg - tg) > tolerance ||
                        Math.Abs(sr - tr) > tolerance)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        private bool PixelMatchSafe(byte[] sourceBytes, byte[] templateBytes, int sx, int sy, int tx, int ty, int sStride, int tStride, int tolerance)
        {
             int sIdx = sy * sStride + sx * 4;
             int tIdx = ty * tStride + tx * 4;

             byte ta = templateBytes[tIdx + 3];
             if (ta < 10) return true; // 透明跳过

             if (Math.Abs(sourceBytes[sIdx] - templateBytes[tIdx]) > tolerance ||       // B
                 Math.Abs(sourceBytes[sIdx + 1] - templateBytes[tIdx + 1]) > tolerance ||   // G
                 Math.Abs(sourceBytes[sIdx + 2] - templateBytes[tIdx + 2]) > tolerance)     // R
             {
                 return false;
             }
             return true;
        }

        public void Dispose()
        {
            ImageOcr? ocrToDispose;
            lock (_ocrEngineSync)
            {
                if (_disposed) return;
                _disposed = true;
                ocrToDispose = _ocrEngine;
                _ocrEngine = null;
            }

            try { ocrToDispose?.Dispose(); } catch { }

            lock (_detectionCacheSync)
            {
                _detectionCache = null;
            }

            _yoloDetector?.Dispose();
        }
        

        /// <summary>
        /// 判断是否为全屏桌面尺寸
        /// </summary>
        private bool IsDesktopPixelSize(RECT rect)
        {
            int w = rect.Right - rect.Left;
            int h = rect.Bottom - rect.Top;
            int screenW = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
            int screenH = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;
            // 如果尺寸完全等于屏幕分辨率，极有可能是桌面
            // (虽然全屏应用也是，但在我们的场景下 WeWork 极少全屏，且如果是全屏通常也不会正好是桌面句柄)
            return w == screenW && h == screenH;
        }

        /// <summary>
        /// 判断是否为系统窗口 (桌面/任务栏)
        /// </summary>
        private bool IsSystemWindowClass(IntPtr hwnd)
        {
            try
            {
                StringBuilder sb = new StringBuilder(256);
                GetClassName(hwnd, sb, sb.Capacity);
                string cls = sb.ToString();
                return cls == "Progman" || cls == "WorkerW" || cls == "Shell_TrayWnd";
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"⚠️ [IsSystemWindowClass] 异常: {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// ✅ [图像预处理] 放大并添加白边，提升 OCR 识别率
        /// </summary>
        private Bitmap PreprocessForOcr(Bitmap original, int scaleFactor, bool highQuality = false)
        {
            const int maxLongEdge = 1800;
            scaleFactor = Math.Clamp(scaleFactor, 1, 3);
            int originalLongEdge = Math.Max(original.Width, original.Height);
            if (originalLongEdge > 0 && originalLongEdge * scaleFactor > maxLongEdge)
            {
                scaleFactor = Math.Max(1, maxLongEdge / originalLongEdge);
            }

            int newWidth = original.Width * scaleFactor;
            int newHeight = original.Height * scaleFactor;

            // 小图保留适量白边，避免边缘字符被检测器裁掉。
            int padding = scaleFactor >= 3 ? 16 : 12;
            int paddedWidth = newWidth + padding * 2;
            int paddedHeight = newHeight + padding * 2;

            var processed = new Bitmap(paddedWidth, paddedHeight, PixelFormat.Format24bppRgb);
            
            using (var g = Graphics.FromImage(processed))
            {
                g.Clear(Color.White);
                g.InterpolationMode = highQuality
                    ? System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
                    : System.Drawing.Drawing2D.InterpolationMode.Bilinear;
                g.PixelOffsetMode = highQuality
                    ? System.Drawing.Drawing2D.PixelOffsetMode.HighQuality
                    : System.Drawing.Drawing2D.PixelOffsetMode.Half;
                g.CompositingQuality = highQuality
                    ? System.Drawing.Drawing2D.CompositingQuality.HighQuality
                    : System.Drawing.Drawing2D.CompositingQuality.HighSpeed;
                g.DrawImage(original, new Rectangle(padding, padding, newWidth, newHeight));
            }
            
            return processed;
        }
    }
}
