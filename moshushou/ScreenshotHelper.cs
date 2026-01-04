using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading.Tasks;
using System.Threading.Tasks;
// using System.Windows; // 移除以避免 Point/Size 歧义
using WeChatOcr; // 确保已通过 NuGet 安装 WeChatOcr.Lite

namespace moshushou
{
    public class ScreenshotHelper
    {
        #region Win32 API Imports
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
        #endregion

        // *** NEW ***: 定义截图裁剪的精确参数
        // private const int LEFT_CROP = 270; // 移入 Config
        // private const int CROP_HEIGHT = 53; // 移入 Config
        // private const int WECHAT_RIGHT_CROP = 125; // 移入 Config
        // private const int WEWORK_RIGHT_CROP = 100; // 移入 Config
        private const int DEFAULT_RIGHT_CROP = 300; // 备用值

        private readonly string _baseDirectory;
        private readonly Action<string> _logAction;
        private readonly SearchConfig _config;


        public ScreenshotHelper(string baseStorageDirectory, SearchConfig config, Action<string> logAction = null)
        {
            _baseDirectory = baseStorageDirectory;
            _config = config;
            _logAction = logAction;
        }


        // ✅ 新增：获取当前窗口顶部的标题文字（复用 CaptureWindowTop 的裁剪逻辑）
        public async Task<string> GetWeChatWindowTitleTextAsync(IntPtr targetHwnd, bool isWework)
        {
            try
            {
                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) return null;

                // 复用 CaptureWindowTop 的逻辑确定裁剪参数
                string appIdentifier = isWework ? "企业微信" : "微信";
                int rightCrop = GetRightCropAmount(appIdentifier);
                int leftCrop = _config.WeChatCropLeft; // 默认 270
                
                int cropHeight = _config.WeChatCropHeight; // 默认 53

                int cropWidth = rect.Right - rect.Left - leftCrop - rightCrop;
                if (cropWidth <= 0) return null;

                // 截图
                using (var bitmap = new Bitmap(cropWidth, cropHeight, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        // 核心：这里的坐标 (rect.Left + LEFT_CROP) 就是原代码中识别群名的位置
                        graphics.CopyFromScreen(rect.Left + leftCrop, rect.Top, 0, 0, new Size(cropWidth, cropHeight), CopyPixelOperation.SourceCopy);
                    }

                    // 放大并 OCR
                    using (var scaledMap = ScaleImage(bitmap, 3))
                    {
                        string ocrText = await PerformOcrAsync(scaledMap);
                        // 清理结果（移除括号等干扰）
                        return CleanGroupName(ocrText);
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 获取标题栏失败: {ex.Message}");
                return null;
            }
        }





        // ✅ [新增] 专门计算输入框点击坐标的方法
        public bool GetInputBoxClickCoordinates(IntPtr hwnd, bool isWework, out int x, out int y)
        {
            x = 0; y = 0;
            if (hwnd == IntPtr.Zero) return false;

            if (!GetWindowRect(hwnd, out RECT rect)) return false;

            // ---------------------------------------------------------
            // 📐 坐标计算核心逻辑
            // ---------------------------------------------------------
            // 微信 PC 版：
            //   侧边栏(图标+列表) 约 250~300px
            //   偏移 310px 比较稳健 (270 + 40)
            //
            // 企业微信：
            //   侧边栏通常更宽
            //   偏移 380px 比较稳健 (310 + 70)
            // ---------------------------------------------------------

            int xOffset = isWework ? 380 : 310;
            // 如果是微信 4.0，可能需要微调，暂时用硬编码，后续也可移入 Config
            if (!isWework && _config.WeChatProcessName != "WeChat") 
            {
                 // 预留位置
            }
            
            int yOffset = 70; // 距离底部的高度

            x = rect.Left + xOffset;
            y = rect.Bottom - yOffset;

            // 简单的越界检查 (防止窗口被缩得太小)
            if (x >= rect.Right) x = rect.Right - 50;
            if (y >= rect.Bottom) y = rect.Bottom - 20;

            return true;
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
            if (actual.Contains(expected) || expected.Contains(actual)) return true;

            // 2. 深度清洗：
            // - 去除所有空白字符 (\s)
            // - 去除常见标点符号 (包括中文标点和截断用的点)
            // - 统一转小写
            string pattern = @"\s+|[.,;:'""()\-\[\]{}<>/\\|、，。；：“”（）—…\.]";
            string cleanTarget = Regex.Replace(expected, pattern, "").ToLower();
            string cleanOCR = Regex.Replace(actual, pattern, "").ToLower();

            // 防止清洗后为空
            if (string.IsNullOrEmpty(cleanTarget) || string.IsNullOrEmpty(cleanOCR)) return false;

            // 3. 【核心优化】智能前缀匹配 (专门解决 "张旭彬...官方旗舰店" 变成 "张旭彬...官方.." 的问题)
            // 逻辑：如果清洗后的 OCR 结果，是 目标词 的“开头部分”，且长度足够长，视为匹配。
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

            // 5. 莱文斯坦距离 (兜底逻辑)
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
        /// ✅ 解决报错：宽松匹配方法
        /// </summary>
        public bool IsTextMatch(string fullText, string keyword)
        {
            if (string.IsNullOrEmpty(fullText) || string.IsNullOrEmpty(keyword)) return false;

            // 移除空白和标点，忽略大小写
            string Clean(string s) => Regex.Replace(s, @"\s+|[.,;:'""()（）]", "").ToLower();

            return Clean(fullText).Contains(Clean(keyword));
        }

        /// <summary>
        /// ✅ [修改版] 截取右侧窗口，并按高度分割验证
        /// 改动：移除放大逻辑，直接使用原图进行神经网络OCR识别
        /// </summary>
        public async Task<(string topText, string bottomText)> CaptureSplitVerificationAsync(IntPtr targetHwnd, bool isWework)
        {
            try
            {
                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) return (null, null);

                string appIdentifier = isWework ? "企业微信" : "微信";
                int rightCrop = GetRightCropAmount(appIdentifier);
                int leftCrop = _config.WeChatCropLeft;
                
                int leftStart = rect.Left + leftCrop;
                int totalWidth = (rect.Right - rect.Left) - leftCrop - rightCrop;
                int totalHeight = rect.Bottom - rect.Top;

                if (totalWidth <= 0 || totalHeight <= 0) return (null, null);

                // 分割线：离底部 250px (涵盖输入框)
                int splitHeightFromBottom = 250;
                if (totalHeight < 500) splitHeightFromBottom = (int)(totalHeight * 0.4);

                int topHeight = totalHeight - splitHeightFromBottom;
                int bottomHeight = splitHeightFromBottom;

                string tText = "", bText = "";

                // --- 截取上半部分 (聊天区) ---
                using (var bmpTop = new Bitmap(totalWidth, topHeight, PixelFormat.Format32bppArgb))
                {
                    using (var g = Graphics.FromImage(bmpTop))
                    {
                        g.CopyFromScreen(leftStart, rect.Top, 0, 0, new Size(totalWidth, topHeight), CopyPixelOperation.SourceCopy);
                    }
                    // ⚡ 原图直接识别 (不放大)
                    tText = await PerformOcrAsync(bmpTop);
                }

                // --- 截取下半部分 (输入区) ---
                using (var bmpBottom = new Bitmap(totalWidth, bottomHeight, PixelFormat.Format32bppArgb))
                {
                    using (var g = Graphics.FromImage(bmpBottom))
                    {
                        g.CopyFromScreen(leftStart, rect.Top + topHeight, 0, 0, new Size(totalWidth, bottomHeight), CopyPixelOperation.SourceCopy);
                    }
                    // ⚡ 原图直接识别 (不放大)
                    bText = await PerformOcrAsync(bmpBottom);
                }

                return (tText, bText);
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 分割验证截图失败: {ex.Message}");
                return (null, null);
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

                    string dateFolder = $"{DateTime.Now:yyyyMMdd}_OCR_Screenshots";
                    string fullDirectoryPath = Path.Combine(_baseDirectory, dateFolder);
                    Directory.CreateDirectory(fullDirectoryPath);

                    string safeStoreName = string.Join("_", storeName.Split(Path.GetInvalidFileNameChars()));
                    string initialFileName = $"{safeStoreName}_{appIdentifier}.png";
                    string filePath = Path.Combine(fullDirectoryPath, initialFileName);

                    bitmap.Save(filePath, ImageFormat.Png);
                    _logAction?.Invoke($"✅ 截图 '{initialFileName}' 已保存，正在启动后台OCR...");

                    // *** 核心 ***: 使用Task.Run在后台线程执行耗时的OCR操作
                    Task.Run(() => PerformOcrAndRenameAsync(filePath, storeName, appIdentifier, onOcrComplete));
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
                    using (var finalBitmapToOcr = ScaleImage(originalBitmap, 3))
                    {
                        string rawText = await PerformOcrAsync(finalBitmapToOcr);
                        recognizedGroupName = CleanGroupName(rawText);
                    }
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



        private async Task<string> PerformOcrAsync(Bitmap bitmap)
        {
            var bytes = ImageToBytes(bitmap);
            var tcs = new TaskCompletionSource<string>();
            var ocr = new ImageOcr();

            ocr.Run(bytes, (path, result) =>
            {
                try
                {
                    if (result?.OcrResult?.SingleResult == null)
                    {
                        tcs.TrySetResult("未识别到文字");
                        return;
                    }
                    var sb = new StringBuilder();
                    foreach (var item in result.OcrResult.SingleResult)
                    {
                        if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8)) sb.Append(item.SingleStrUtf8);
                    }
                    tcs.TrySetResult(sb.ToString().Trim());
                }
                catch (Exception ex) { tcs.TrySetException(ex); }
                finally
                {
                    try { if (File.Exists(path)) File.Delete(path); } catch { /* ignore */ }
                }
            });

            // 设置一个超时，防止OCR进程卡死
            var timeoutTask = Task.Delay(8000);
            var completedTask = await Task.WhenAny(tcs.Task, timeoutTask);

            if (completedTask == timeoutTask)
            {
                return "OCR识别超时";
            }
            return await tcs.Task;
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
        /// 🚫 新增：排除"搜索网络结果"区域的匹配项
        /// </summary>
        public async Task<bool> CheckSearchResultAsync(IntPtr targetHwnd, string expectedText, bool isWework)
        {
            try
            {
                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) return false;

                int relX, relY, width, height;

                if (isWework)
                {
                    // 🏢 企业微信坐标
                    int[] rects = _config.WeWorkSearchResultRect ?? new int[] { 78, 90, 394, 58 };
                    relX = rects[0]; relY = rects[1];
                    width = rects[2];
                    height = rects[3];
                }
                else
                {
                    // 💬 微信坐标 - 强制使用用户指定参数 (忽略Config干扰)
                    int[] rects = new int[] { 74, 57, 326, 400 };
                    relX = rects[0]; relY = rects[1];
                    width = rects[2];
                    height = rects[3];
                }

                int screenX = rect.Left + relX;
                int screenY = rect.Top + relY;

                if (width <= 0 || height <= 0) return false;

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // 使用 3 倍放大进行 OCR
                    using (var scaledMap = ScaleImage(bitmap, 3))
                    {
                        var bytes = ImageToBytes(scaledMap);
                        // ✅ [修复] 强制异步延续，防止 OCR 回调线程（非托管/临时）卡死后续逻辑
                        var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                        var ocr = new ImageOcr();

                        ocr.Run(bytes, (path, result) =>
                        {
                            try
                            {
                                if (result?.OcrResult?.SingleResult == null)
                                {
                                    System.Diagnostics.Debug.WriteLine("❌ [CheckSearchResult] OCR 结果为空");
                                    tcs.TrySetResult(false);
                                    return;
                                }

                                // 🚫 步骤 1: 定位"搜索网络结果"区域
                                float searchWebOriginalY_Start = -1;
                                float searchWebOriginalY_End = -1;

                                foreach (var item in result.OcrResult.SingleResult)
                                {
                                    if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8))
                                    {
                                        string text = item.SingleStrUtf8;
                                        // 支持多种 OCR 识别变体
                                        if (text.Contains("搜索网络结果") || text.Contains("搜索网络") || 
                                            text.Contains("搜一搜") || text.Contains("网络搜索") || 
                                            text.Contains("网络结果") || text.Contains("六搜索") ||
                                            text.Contains("不搜索") || text.Contains("搜索网"))
                                        {
                                            searchWebOriginalY_Start = item.Top / 3.0f;
                                            System.Diagnostics.Debug.WriteLine($"🚫 [CheckSearchResult] 发现网络搜索标记 '{text}'，原图StartY={searchWebOriginalY_Start}");
                                            break;
                                        }
                                    }
                                }

                                // 🔍 步骤 1.5: 用图像搜索 search.png 来确定排除区域
                                // ✅ 修复：无论 OCR 是否找到起点，都需要用图像搜索确定终点
                                try
                                {
                                    string iconPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "search.png");
                                    if (System.IO.File.Exists(iconPath))
                                    {
                                        using (var iconBmp = new Bitmap(iconPath))
                                        {
                                            // 如果 OCR 没找到起点，用图像搜索找第一个图标作为起点
                                            if (searchWebOriginalY_Start < 0)
                                            {
                                                var firstRect = FindFirstImageOccurrence(bitmap, iconBmp);
                                                if (firstRect.HasValue)
                                                {
                                                    searchWebOriginalY_Start = firstRect.Value.Top - 10;
                                                    System.Diagnostics.Debug.WriteLine($"🚫 [CheckSearchResult] (图像兜底) 找到第一个search图标，设定排除起点Y={searchWebOriginalY_Start}");
                                                }
                                            }

                                            // ✅ 关键修复：无论起点来源，都要用图像搜索找终点
                                            if (searchWebOriginalY_Start >= 0)
                                            {
                                                var lastRect = FindLastImageOccurrence(bitmap, iconBmp);
                                                if (lastRect.HasValue)
                                                {
                                                    searchWebOriginalY_End = lastRect.Value.Bottom;
                                                    System.Diagnostics.Debug.WriteLine($"🚫 [CheckSearchResult] 找到最后一个search图标，原图EndY={searchWebOriginalY_End}");
                                                }
                                                else
                                                {
                                                    // 如果图像搜索也没找到，设置默认范围（起点+80px）
                                                    searchWebOriginalY_End = searchWebOriginalY_Start + 80;
                                                    System.Diagnostics.Debug.WriteLine($"🚫 [CheckSearchResult] 未找到search图标，使用默认范围EndY={searchWebOriginalY_End}");
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // search.png 不存在时，设置默认范围
                                        if (searchWebOriginalY_Start >= 0)
                                        {
                                            searchWebOriginalY_End = searchWebOriginalY_Start + 80;
                                            System.Diagnostics.Debug.WriteLine($"⚠️ [CheckSearchResult] search.png 不存在，使用默认范围EndY={searchWebOriginalY_End}");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"⚠️ [CheckSearchResult] 图像搜索出错: {ex.Message}");
                                    // 出错时设置默认范围
                                    if (searchWebOriginalY_Start >= 0 && searchWebOriginalY_End < searchWebOriginalY_Start)
                                    {
                                        searchWebOriginalY_End = searchWebOriginalY_Start + 80;
                                    }
                                }

                                // ✅ 步骤 2: 在安全区域内查找目标
                                bool foundInSafeZone = false;

                                foreach (var item in result.OcrResult.SingleResult)
                                {
                                    if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8))
                                    {
                                        if (IsFuzzyMatch(expectedText, item.SingleStrUtf8))
                                        {
                                            float itemOriginalTop = item.Top / 3.0f;
                                            float itemOriginalCenterY = (item.Top + item.Bottom) / 2.0f / 3.0f;

                                            // 安全检查：确保不在网络搜索区域内
                                            bool isInDangerZone = false;
                                            if (searchWebOriginalY_Start >= 0 && searchWebOriginalY_End > searchWebOriginalY_Start)
                                            {
                                                // 精确区间排除：只排除 [起点-5, 终点+5] 区间内的内容
                                                if (itemOriginalCenterY >= searchWebOriginalY_Start - 5 && itemOriginalCenterY <= searchWebOriginalY_End + 5)
                                                {
                                                    isInDangerZone = true;
                                                }
                                            }

                                            if (isInDangerZone)
                                            {
                                                System.Diagnostics.Debug.WriteLine($"🚫 [CheckSearchResult] 目标 '{item.SingleStrUtf8}' 在网络搜索区域内 (Y={itemOriginalTop})，忽略");
                                                continue; // 继续查找其他匹配项
                                            }

                                            System.Diagnostics.Debug.WriteLine($"✅ [CheckSearchResult] 在安全区域找到目标: '{item.SingleStrUtf8}' (Y={itemOriginalTop})");
                                            foundInSafeZone = true;
                                            break;
                                        }
                                    }
                                }

                                if (!foundInSafeZone)
                                {
                                    System.Diagnostics.Debug.WriteLine($"❌ [CheckSearchResult] 未在安全区域找到目标 '{expectedText}'");
                                }

                                tcs.TrySetResult(foundInSafeZone);
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"💥 [CheckSearchResult] OCR 处理异常: {ex.Message}");
                                tcs.TrySetException(ex);
                            }
                            finally
                            {
                                try { if (File.Exists(path)) File.Delete(path); } catch { }
                            }
                        });

                        var completedTask = await Task.WhenAny(tcs.Task, Task.Delay(8000));
                        
                        // ✅ [防止GC] 保持 ocr 对象存活，防止在回调回来之前被回收导致 native crash
                        GC.KeepAlive(ocr);

                        if (completedTask == tcs.Task)
                        {
                            return await tcs.Task;
                        }
                        
                        System.Diagnostics.Debug.WriteLine("⚠️ [CheckSearchResult] OCR 超时");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 搜索验证出错: {ex.Message}");
                return false;
            }
            finally
            {
                 System.Diagnostics.Debug.WriteLine("🏁 [DEBUG_TRACE] CheckSearchResultAsync 方法结束 (Finally)");
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
                int relX = relativeSearchArea?[0] ?? 0;
                int relY = relativeSearchArea?[1] ?? 60; // 避开顶栏
                int width = relativeSearchArea?[2] ?? 400; // 搜索栏通常在左侧
                int height = relativeSearchArea?[3] ?? (rect.Bottom - rect.Top - 60);

                int screenX = rect.Left + relX;
                int screenY = rect.Top + relY;

                // 🔧 DEBUG: 打印详细坐标信息
                _logAction?.Invoke($"🔍 [FindKeyword] WindowRect: {rect.Left},{rect.Top} size {rect.Right-rect.Left}x{rect.Bottom-rect.Top}");
                _logAction?.Invoke($"🔍 [FindKeyword] SearchArea Screen: {screenX},{screenY} {width}x{height}");

                if (width <= 0 || height <= 0) return null;

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    var bytes = ImageToBytes(bitmap);
                    // ✅ [修复] 强制异步延续
                    var tcs = new TaskCompletionSource<Point?>(TaskCreationOptions.RunContinuationsAsynchronously);
                    var ocr = new ImageOcr();

                    ocr.Run(bytes, (path, result) =>
                    {
                        try
                        {
                            if (result?.OcrResult?.SingleResult != null)
                            {
                                foreach (var item in result.OcrResult.SingleResult)
                                {
                                    if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8))
                                    {
                                        if (item.SingleStrUtf8.Contains(keyword))
                                        {
                                            // 尝试直接访问坐标属性
                                            // 如果 WeChatOcr 的 item 定义是 flat 的
                                            if (item.Left >= 0) 
                                            {
                                                int itemX = (int)(item.Left + (item.Right - item.Left) / 2);
                                                int itemY = (int)(item.Top + (item.Bottom - item.Top) / 2);
                                                
                                                int centerX = screenX + itemX;
                                                int centerY = screenY + itemY;

                                                _logAction?.Invoke($"✅ [FindKeyword] Found '{keyword}' at Rel({itemX},{itemY}) -> Screen({centerX}, {centerY})");
                                                tcs.TrySetResult(new Point(centerX, centerY));
                                                return;
                                            }
                                        }
                                    }
                                }
                            }
                            tcs.TrySetResult(null);
                        }
                        catch (Exception ex) { tcs.TrySetException(ex); }
                        finally
                        {
                            try { if (File.Exists(path)) File.Delete(path); } catch {}
                        }
                    });

                    var completedTask = await Task.WhenAny(tcs.Task, Task.Delay(5000));
                    
                    GC.KeepAlive(ocr); // ✅ 防止过早回收

                    if (completedTask == tcs.Task) return await tcs.Task;
                    return null;
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 查找关键词 '{keyword}' 失败: {ex.Message}");
                return null;
            }
        }


        /// <summary>
        /// ✅ [优化版] 动态定位群聊点击位置
        /// 策略优先级:
        /// 1. 查找"最常使用"或"群聊"锚点，点击其下方
        /// 2. 如果没有锚点，直接搜索目标群聊名称并点击
        /// 3. 排除"搜索网络结果"下方的内容
        /// </summary>
        /// <param name="targetHwnd">微信/企业微信窗口句柄</param>
        /// <param name="targetGroupName">要搜索的目标群聊名称 (可选)</param>
        /// <param name="isWework">是否为企业微信</param>
        /// <returns>屏幕点击坐标, null 表示未找到</returns>
        public async Task<Point?> FindGroupChatClickPositionAsync(IntPtr targetHwnd, string targetGroupName = null, bool isWework = false)
        {
            System.Diagnostics.Debug.WriteLine($"🔍 [FindGroupChat] 方法入口, hwnd={targetHwnd}, 目标群聊='{targetGroupName}', isWework={isWework}");
            try
            {
                if (targetHwnd == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine("❌ [FindGroupChat] hwnd 为 Zero，返回 null");
                    return null;
                }

                bool rectResult = GetWindowRect(targetHwnd, out RECT rect);
                System.Diagnostics.Debug.WriteLine($"🔍 [FindGroupChat] GetWindowRect 结果: {rectResult}");
                
                if (!rectResult)
                {
                    System.Diagnostics.Debug.WriteLine("❌ [FindGroupChat] GetWindowRect 失败，返回 null");
                    return null;
                }
                
                // 🛡️ [自动纠错] 根据进程名强制修正 isWework 标志
                try
                {
                    GetWindowThreadProcessId(targetHwnd, out uint processId);
                    using (var process = Process.GetProcessById((int)processId))
                    {
                        string processName = process.ProcessName;
                        bool isRealWework = "WXWork".Equals(processName, StringComparison.OrdinalIgnoreCase);
                        bool isRealWeChat = "WeChat".Equals(processName, StringComparison.OrdinalIgnoreCase);

                        if (isRealWework && !isWework)
                        {
                            System.Diagnostics.Debug.WriteLine($"⚠️ [FindGroupChat] 检测到窗口是企业微信，但参数为微信。自动修正 isWework=True");
                            isWework = true;
                        }
                        else if (isRealWeChat && isWework)
                        {
                            System.Diagnostics.Debug.WriteLine($"⚠️ [FindGroupChat] 检测到窗口是微信，但参数为企业微信。自动修正 isWework=False");
                            isWework = false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ [FindGroupChat] 进程名检测/纠错失败: {ex.Message}");
                }

                // 搜索下拉区域：根据是微信还是企业微信使用不同的配置
                int relX, relY, width, height;
                
                if (isWework)
                {
                    // 🏢 企业微信配置
                    relX = 78;
                    relY = 90; // 用户回退了之前的修改，保持 90
                    width = 394;
                    height = 400;
                    System.Diagnostics.Debug.WriteLine("🏢 [FindGroupChat] 使用企业微信截图区域配置");
                }
                else
                {
                    // 💬 微信配置
                    // ⚠️ 强制使用用户指定坐标，屏蔽可能存在的旧 Config (如 0 或 53)
                    // int[] userRects = _config.WeChatSearchResultRect;
                    int[] userRects = null; // FORCE NULL to skip config block
                    
                    if (userRects != null && userRects.Length >= 4)
                    {
                        relX = userRects[0];
                        relY = userRects[1]; 
                        width = userRects[2];
                        height = 350;
                        System.Diagnostics.Debug.WriteLine($"💬 [FindGroupChat] 使用用户 Config 坐标: X={relX}, Y={relY}, W={width}");
                    }
                    else
                    {
                        relX = 74;
                        relY = 57;
                        width = 326; // 400 - 74
                        height = 350;
                        System.Diagnostics.Debug.WriteLine($"💬 [FindGroupChat] 强制使用用户指定坐标: X=74, Y=57, W=326 (已屏蔽Config)");
                    }
                }



                // � [DEBUG 核心示踪剂] 强制弹窗，证明新代码在运行
                // 验证完后必须删除！
                // 🔔 [DEBUG 核心示踪剂] 已移除
                // MessageBox.Show(...)

                int screenX = rect.Left + relX;
                int screenY = rect.Top + relY;

                // 🔧 DEBUG: 输出到 VS 输出窗口
                System.Diagnostics.Debug.WriteLine($"🔍 [FindGroupChat] 窗口坐标: rect.Left={rect.Left}, rect.Top={rect.Top}, rect.Right={rect.Right}, rect.Bottom={rect.Bottom}");
                System.Diagnostics.Debug.WriteLine($"🔍 [FindGroupChat] 窗口尺寸: {rect.Right - rect.Left}x{rect.Bottom - rect.Top}");
                System.Diagnostics.Debug.WriteLine($"🔍 [FindGroupChat] 截图区域: 相对(relX={relX}, relY={relY}) -> 屏幕({screenX},{screenY}) {width}x{height}");
                _logAction?.Invoke($"🔍 [FindGroupChat] 窗口坐标: rect.Left={rect.Left}, rect.Top={rect.Top}");
                _logAction?.Invoke($"🔍 [FindGroupChat] 截图区域: 屏幕({screenX},{screenY}) {width}x{height}");

                if (width <= 0 || height <= 0) return null;

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // 🔧 DEBUG: 保存原始截图供调试
                    string debugDir = Path.Combine(_baseDirectory, "Debug_GroupChat");
                    try
                    {
                        Directory.CreateDirectory(debugDir);
                        string origFilename = $"GroupChat_原图_{DateTime.Now:HHmmss_fff}.png";
                        bitmap.Save(Path.Combine(debugDir, origFilename), ImageFormat.Png);
                        _logAction?.Invoke($"🧐 [调试] 原始截图已保存: {origFilename}");
                    }
                    catch { }

                    // ✅ 关键优化：使用 3 倍放大提高 OCR 准确率
                    using (var scaledBitmap = ScaleImage(bitmap, 3))
                    {
                        // 🔧 DEBUG: 保存放大后截图供调试
                        try
                        {
                            string scaledFilename = $"GroupChat_放大3x_{DateTime.Now:HHmmss_fff}.png";
                            scaledBitmap.Save(Path.Combine(debugDir, scaledFilename), ImageFormat.Png);
                            _logAction?.Invoke($"🧐 [调试] 放大3x截图已保存: {scaledFilename}");
                        }
                        catch { }

                        var bytes = ImageToBytes(scaledBitmap);
                        // ✅ [修复] 强制异步延续
                        var tcs = new TaskCompletionSource<Point?>(TaskCreationOptions.RunContinuationsAsynchronously);
                        var ocr = new ImageOcr();

                        ocr.Run(bytes, (path, result) =>
                        {
                            try
                            {
                                if (result?.OcrResult?.SingleResult == null)
                                {
                                    System.Diagnostics.Debug.WriteLine("❌ [FindGroupChat] OCR 结果为空");
                                    tcs.TrySetResult(null);
                                    return;
                                }

                                // 🔧 DEBUG: 输出所有 OCR 识别结果
                                System.Diagnostics.Debug.WriteLine($"📋 [FindGroupChat] OCR 识别到 {result.OcrResult.SingleResult.Count} 个文本块:");
                                foreach (var item in result.OcrResult.SingleResult)
                                {
                                    if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8))
                                    {
                                        // 注意：坐标是放大后的，需要除以 3 才是原图坐标
                                        System.Diagnostics.Debug.WriteLine($"   - '{item.SingleStrUtf8}' at ({item.Left/3},{item.Top/3})-({item.Right/3},{item.Bottom/3})");
                                    }
                                }

                                // 🚫 排除逻辑：先找"搜索网络结果"的位置，及其下方的图标结束位置
                                float searchWebOriginalY_Start = -1; // 原图坐标系的排除起点
                                float searchWebOriginalY_End = -1;   // 原图坐标系的排除终点

                                // 1. 先定位"搜索网络结果"文字位置
                                foreach (var item in result.OcrResult.SingleResult)
                                {
                                    if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8))
                                    {
                                        string text = item.SingleStrUtf8;
                                        // 增加更多关键词
                                        if (text.Contains("搜索网络结果") || text.Contains("搜索网络") || text.Contains("搜一搜") || text.Contains("网络搜索") || text.Contains("网络结果"))
                                        {
                                            searchWebOriginalY_Start = item.Top / 3.0f; // 转换为原图坐标
                                            System.Diagnostics.Debug.WriteLine($"🚫 [FindGroupChat] 发现网络搜索标记 '{text}'，原图StartY={searchWebOriginalY_Start}");
                                            break;
                                        }
                                    }
                                }

                                    try 
                                    {
                                        string iconPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "search.png");
                                        if (System.IO.File.Exists(iconPath))
                                        {
                                            using (var iconBmp = new Bitmap(iconPath))
                                            {
                                                // 🅰️ 如果 OCR 没找到文本起点，尝试用图像搜索找第一个 search.png 作为起点
                                                if (searchWebOriginalY_Start < 0)
                                                {
                                                    var firstRect = FindFirstImageOccurrence(bitmap, iconBmp);
                                                    if (firstRect.HasValue)
                                                    {
                                                        searchWebOriginalY_Start = firstRect.Value.Top - 10; 
                                                        System.Diagnostics.Debug.WriteLine($"🚫 [FindGroupChat] (图像兜底) 找到第一个search图标，设定排除起点Y={searchWebOriginalY_Start}");
                                                    }
                                                }

                                                // 🅱️ ✅ 修复：无论起点来源，都要用图像搜索找终点
                                                if (searchWebOriginalY_Start >= 0)
                                                {
                                                    var lastRect = FindLastImageOccurrence(bitmap, iconBmp);
                                                    if (lastRect.HasValue)
                                                    {
                                                        searchWebOriginalY_End = lastRect.Value.Bottom; 
                                                        System.Diagnostics.Debug.WriteLine($"🚫 [FindGroupChat] 找到最后一个search图标，原图EndY={searchWebOriginalY_End}");
                                                    }
                                                    else
                                                    {
                                                        // 如果图像搜索也没找到，设置默认范围（起点+80px）
                                                        searchWebOriginalY_End = searchWebOriginalY_Start + 80;
                                                        System.Diagnostics.Debug.WriteLine($"🚫 [FindGroupChat] 未找到search图标，使用默认范围EndY={searchWebOriginalY_End}");
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"⚠️ [FindGroupChat] search.png 不存在于 {iconPath}");
                                            // search.png 不存在时，设置默认范围
                                            if (searchWebOriginalY_Start >= 0)
                                            {
                                                searchWebOriginalY_End = searchWebOriginalY_Start + 80;
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"⚠️ [FindGroupChat] 图像搜索出错: {ex.Message}");
                                        // 出错时设置默认范围
                                        if (searchWebOriginalY_Start >= 0 && searchWebOriginalY_End < searchWebOriginalY_Start)
                                        {
                                            searchWebOriginalY_End = searchWebOriginalY_Start + 80;
                                        }
                                    }
                                    
                                    // ✅ 企业微信特殊处理：优先匹配目标名称
                                if (isWework)
                                {
                                    System.Diagnostics.Debug.WriteLine($"🏢 [FindGroupChat] 企业微信模式，目标: '{targetGroupName}'");
                                    
                                    dynamic bestItem = null;
                                    float minY = float.MaxValue;
                                    dynamic firstItem = null;

                                    foreach (var item in result.OcrResult.SingleResult)
                                    {
                                        if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8))
                                        {
                                            string text = item.SingleStrUtf8;
                                            
                                            if (item.Top < minY)
                                            {
                                                minY = item.Top;
                                                firstItem = item;
                                            }

                                            // 尝试匹配目标
                                            if (!string.IsNullOrEmpty(targetGroupName) && IsFuzzyMatch(targetGroupName, text))
                                            {
                                                bestItem = item;
                                                System.Diagnostics.Debug.WriteLine($"✅ [FindGroupChat] 企业微信匹配成功: '{text}'");
                                                break; // 找到就停
                                            }
                                        }
                                    }

                                    if (bestItem != null)
                                    {
                                        // 命中目标
                                        int itemCenterX = (int)((bestItem.Left + bestItem.Right) / 2 / 3);
                                        int itemCenterY = (int)((bestItem.Top + bestItem.Bottom) / 2 / 3);
                                        int weworkClickX = screenX + itemCenterX;
                                        int weworkClickY = screenY + itemCenterY;

                                        System.Diagnostics.Debug.WriteLine($"✅ [FindGroupChat] 锁定目标坐标: Screen({weworkClickX},{weworkClickY})");
                                        tcs.TrySetResult(new Point(weworkClickX, weworkClickY));
                                        return;
                                    }
                                    else
                                    {
                                         // ❌ 未找到目标
                                         System.Diagnostics.Debug.WriteLine($"❌ [FindGroupChat] 企业微信未找到与 '{targetGroupName}' 匹配的项，停止操作防误触。");
                                         
                                         if (firstItem != null)
                                            System.Diagnostics.Debug.WriteLine($"ℹ️ [FindGroupChat] 排首位的项是: '{firstItem.SingleStrUtf8}' (未命中目标)");

                                         tcs.TrySetResult(null);
                                         return;
                                    }
                                }

                                // --- 以下是微信 (WeChat) 的逻辑 --- 

                                // [重复代码已移除] 之前已经在上方查找过 searchWebOriginalY_Start
                                
                                // 📊 如果没找到搜索网络结果，输出调试信息
                                if (searchWebOriginalY_Start < 0)
                                {
                                    System.Diagnostics.Debug.WriteLine("⚠️ [FindGroupChat] 未识别到'搜索网络结果'文字");
                                }

                                // ✅ 1. [直接匹配] 尝试直接查找目标群聊名称 (必须在"搜索网络结果"上方)
                                // 这可以解决锚点(如"最常使用")OCR识别失败导致无法点击的问题
                                if (!string.IsNullOrEmpty(targetGroupName))
                                {
                                    foreach (var item in result.OcrResult.SingleResult)
                                    {
                                        if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8))
                                        {
                                             // 这里的 item 是放大3倍后的坐标
                                             if (IsFuzzyMatch(targetGroupName, item.SingleStrUtf8))
                                             {
                                                 float itemOriginalBottom = item.Bottom / 3.0f;
                                                 float itemOriginalTop = item.Top / 3.0f;
                                                 float itemOriginalCenterY = (item.Top + item.Bottom) / 2.0f / 3.0f;
                                                 
                                                 // 安全检查：确保不在网络搜索区域 [起点, 终点] 内
                                                 bool isSafe = true;
                                                 if (searchWebOriginalY_Start >= 0 && searchWebOriginalY_End > searchWebOriginalY_Start)
                                                 {
                                                     // 精确区间排除：只排除 [起点-5, 终点+5] 区间内的内容
                                                     if (itemOriginalCenterY >= searchWebOriginalY_Start - 5 && itemOriginalCenterY <= searchWebOriginalY_End + 5)
                                                     {
                                                         System.Diagnostics.Debug.WriteLine($"🚫 [FindGroupChat] (直接匹配) 目标 '{item.SingleStrUtf8}' 在网络搜索区间内 (Y={itemOriginalCenterY}, 区间=[{searchWebOriginalY_Start}, {searchWebOriginalY_End}])，禁止点击。");
                                                         isSafe = false;
                                                     }
                                                 }

                                                 if (isSafe)
                                                 {
                                                      int itemCenterX = (int)((item.Left + item.Right) / 2 / 3);
                                                      int itemCenterY = (int)((item.Top + item.Bottom) / 2 / 3);
                                                      int directClickX = screenX + itemCenterX; 
                                                      int directClickY = screenY + itemCenterY;

                                                      // 🎯 [优化] 如果目标在最顶部 (Top < 80px)，通常是"最常使用"或最佳匹配，直接回车更稳
                                                      if (itemOriginalTop < 80)
                                                      {
                                                          System.Diagnostics.Debug.WriteLine($"🎯 [FindGroupChat] (微信) 目标 '{item.SingleStrUtf8}' 位于顶部 (Y={itemOriginalTop}<80)，返回特殊坐标(-1,-1)表示直接回车");
                                                          tcs.TrySetResult(new Point(-1, -1));
                                                      }
                                                      else
                                                      {
                                                          System.Diagnostics.Debug.WriteLine($"✅ [FindGroupChat] (微信) 直接命中目标: '{item.SingleStrUtf8}' Screen({directClickX},{directClickY})");
                                                          tcs.TrySetResult(new Point(directClickX, directClickY));
                                                      }
                                                      return;
                                                 }
                                             }
                                        }
                                    }
                                }

                                // 2. [锚点定位] 按优先级查找锚点关键词 (必须在"搜索网络结果"上方)
                                string[] anchorKeywords = { "最常使用", "群聊" };
                                dynamic anchorItem = null;
                                string foundKeyword = null;

                                foreach (var keyword in anchorKeywords)
                                {
                                    foreach (var item in result.OcrResult.SingleResult)
                                    {
                                        if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8))
                                        {
                                            // ⬇️ 排除逻辑：只排除 [Start, End] 区间内的内容
                                            // item.Top 是放大后的，先转回原图
                                            float itemOriginalTop = item.Top / 3.0f;
                                            float itemOriginalBottom = item.Bottom / 3.0f;
                                            float itemOriginalCenterY = (itemOriginalTop + itemOriginalBottom) / 2.0f;

                                            if (searchWebOriginalY_Start >= 0 && searchWebOriginalY_End > searchWebOriginalY_Start)
                                            {
                                                // 精确区间排除：只排除 [起点-5, 终点+5] 区间内的内容
                                                if (itemOriginalCenterY >= searchWebOriginalY_Start - 5 && itemOriginalCenterY <= searchWebOriginalY_End + 5)
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"🚫 [FindGroupChat] 排除位于网络搜索区间的项: '{item.SingleStrUtf8}' (Y={itemOriginalCenterY})");
                                                    continue;
                                                }
                                            }

                                            if (item.SingleStrUtf8.Contains(keyword))
                                            {
                                                anchorItem = item;
                                                foundKeyword = keyword;
                                                System.Diagnostics.Debug.WriteLine($"✅ [FindGroupChat] 找到锚点: '{item.SingleStrUtf8}' at 原图坐标({item.Left/3},{item.Top/3})-({item.Right/3},{item.Bottom/3})");
                                                
                                                // 🎯 关键优化：如果是"最常使用"，直接回车即可进入
                                                if (keyword == "最常使用")
                                                {
                                                    System.Diagnostics.Debug.WriteLine("🎯 [FindGroupChat] 检测到'最常使用'，返回特殊坐标(-1,-1)表示直接回车");
                                                    tcs.TrySetResult(new Point(-1, -1));  // 特殊标记
                                                    return;
                                                }
                                                
                                                break;
                                            }
                                        }
                                    }
                                    if (anchorItem != null) break;
                                }

                                if (anchorItem == null)
                                {
                                    System.Diagnostics.Debug.WriteLine("❌ [FindGroupChat] 未找到 '最常使用' 或 '群聊' 锚点。");
                                    System.Diagnostics.Debug.WriteLine("🛑 [安全模式] 为防止误点网络搜索结果，严格禁止无锚点点击。");
                                    _logAction?.Invoke("⚠️ 未找到'群聊'或'最常使用'分类，判定为未搜到目标，停止操作。");
                                    
                                    tcs.TrySetResult(null);
                                    return;
                                }

                                // 计算点击位置（注意：OCR 坐标是放大后的，需要除以 3）
                                // 修正：X坐标不再取中心，而是取左侧靠右一点的位置，防止偏右
                                // 锚点本身在左侧，所以 safeClickX 应该是 截图左边界 + 锚点文本的一半宽度
                                int ocrAnchorCenterX = (int)((anchorItem.Left + anchorItem.Right) / 2 / 3);
                                int clickX = screenX + ocrAnchorCenterX; 

                                // 兜底：如果算出来太偏右，强制限制在左侧 150px 范围内
                                int maxOffset = 150;
                                if (clickX > rect.Left + maxOffset) clickX = rect.Left + 100;
                                
                                int originalBottom = (int)(anchorItem.Bottom / 3);
                                int clickY = screenY + originalBottom + 20;

                                _logAction?.Invoke($"✅ [FindGroupChat] 锚点 '{foundKeyword}' 原图Bottom={originalBottom}, 计算点击坐标: Screen({clickX}, {clickY})");
                                tcs.TrySetResult(new Point(clickX, clickY));
                            }
                            catch (Exception ex)
                            {
                                _logAction?.Invoke($"💥 [FindGroupChat] OCR 处理异常: {ex.Message}");
                                tcs.TrySetException(ex);
                            }
                            finally
                            {
                                try { if (File.Exists(path)) File.Delete(path); } catch { }
                            }
                        });

                        var completedTask = await Task.WhenAny(tcs.Task, Task.Delay(5000));
                        
                        GC.KeepAlive(ocr); // ✅ 防止过早回收

                        if (completedTask == tcs.Task) return await tcs.Task;
                        
                        _logAction?.Invoke("⚠️ [FindGroupChat] OCR 超时");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 [FindGroupChat] 异常: {ex.Message}");
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
                    
                    byte[] bytes;
                    using (var ms = new MemoryStream()) { bitmap.Save(ms, ImageFormat.Png); bytes = ms.ToArray(); }

                    var tcs = new TaskCompletionSource<System.Drawing.Point?>();
                    var ocr = new ImageOcr(); // ✅ Fix: Instantiate local ocr

                    ocr.Run(bytes, (path, result) =>
                    {
                        try
                        {
                            if (result?.OcrResult?.SingleResult != null)
                            {
                                var items = result.OcrResult.SingleResult;
                                dynamic anchorItem = null;
                                foreach (var item in items)
                                {
                                    if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8) && item.SingleStrUtf8.Contains(anchorKeyword))
                                    {
                                        anchorItem = item;
                                        break;
                                    }
                                }

                                if (anchorItem != null)
                                {
                                    int anchorBottom = (int)anchorItem.Bottom;

                                    dynamic targetItem = null;
                                    int minDistance = int.MaxValue;

                                    foreach (var item in items)
                                    {
                                        if (item == anchorItem) continue;
                                        if (string.IsNullOrEmpty(item.SingleStrUtf8)) continue;

                                        int itemTop = (int)item.Top;

                                        // 垂直判定: 在下方且距离不远
                                        if (itemTop > anchorBottom && itemTop < anchorBottom + 150)
                                        {
                                            int distance = itemTop - anchorBottom;
                                            if (distance < minDistance)
                                            {
                                                targetItem = item;
                                                minDistance = distance;
                                            }
                                        }
                                    }

                                    if (targetItem != null)
                                    {
                                        int itemX = (int)(targetItem.Left + (targetItem.Right - targetItem.Left) / 2);
                                        int itemY = (int)(targetItem.Top + (targetItem.Bottom - targetItem.Top) / 2);
                                        int centerX = screenX + itemX;
                                        int centerY = screenY + itemY;
                                        _logAction?.Invoke($"✅ [FindStructure] Found target below '{anchorKeyword}' -> '{targetItem.SingleStrUtf8}'");
                                        tcs.TrySetResult(new System.Drawing.Point(centerX, centerY));
                                        return;
                                    }
                                }
                            }
                            tcs.TrySetResult(null);
                        }
                        catch (Exception ex) { tcs.TrySetException(ex); }
                        finally { try { if (File.Exists(path)) File.Delete(path); } catch {} }
                    });

                    if (await Task.WhenAny(tcs.Task, Task.Delay(2000)) == tcs.Task) return await tcs.Task;
                    return null;
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
             if (ta < 10) return true;

             if (Math.Abs(sourceBytes[sIdx] - templateBytes[tIdx]) > tolerance ||       // B
                 Math.Abs(sourceBytes[sIdx+1] - templateBytes[tIdx+1]) > tolerance ||   // G
                 Math.Abs(sourceBytes[sIdx+2] - templateBytes[tIdx+2]) > tolerance)     // R
             {
                 return false;
             }
             return true;
        }
    }
}