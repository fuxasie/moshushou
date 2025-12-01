using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WeChatOcr; // 确保已通过 NuGet 安装 WeChatOcr.Lite

namespace moshushou
{
    public class ScreenshotHelper
    {
        #region Win32 API Imports
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
        #endregion

        // *** NEW ***: 定义截图裁剪的精确参数
        private const int LEFT_CROP = 270;
        private const int CROP_HEIGHT = 53;
        private const int WECHAT_RIGHT_CROP = 125;
        private const int WEWORK_RIGHT_CROP = 100;
        private const int DEFAULT_RIGHT_CROP = 300; // 备用值

        private readonly string _baseDirectory;
        private readonly Action<string> _logAction;




        public ScreenshotHelper(string baseStorageDirectory, Action<string> logAction = null)
        {
            _baseDirectory = baseStorageDirectory;
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

                int cropWidth = rect.Right - rect.Left - LEFT_CROP - rightCrop;
                if (cropWidth <= 0) return null;

                // 截图
                using (var bitmap = new Bitmap(cropWidth, CROP_HEIGHT, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        // 核心：这里的坐标 (rect.Left + LEFT_CROP) 就是原代码中识别群名的位置
                        graphics.CopyFromScreen(rect.Left + LEFT_CROP, rect.Top, 0, 0, new Size(cropWidth, CROP_HEIGHT), CopyPixelOperation.SourceCopy);
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
            if (cleanTarget.Contains(cleanOCR) && cleanOCR.Length >= 4) return true;

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
                int leftStart = rect.Left + LEFT_CROP;
                int totalWidth = (rect.Right - rect.Left) - LEFT_CROP - rightCrop;
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

                int cropWidth = rect.Right - rect.Left - LEFT_CROP - rightCrop;
                if (cropWidth <= 0)
                {
                    _logAction?.Invoke($"❌ 截图失败: 窗口宽度不足，无法裁剪。");
                    return;
                }

                // 直接创建裁剪后尺寸的Bitmap
                using (var bitmap = new Bitmap(cropWidth, CROP_HEIGHT, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        // 从屏幕的指定位置直接复制已裁剪的区域到Bitmap中
                        graphics.CopyFromScreen(rect.Left + LEFT_CROP, rect.Top, 0, 0, new Size(cropWidth, CROP_HEIGHT), CopyPixelOperation.SourceCopy);
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
                case "企业微信": return WEWORK_RIGHT_CROP;
                case "微信": return WECHAT_RIGHT_CROP;
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
        /// ✅ [通用版] 验证搜索结果列表
        /// 支持：微信 (53,93) 和 企业微信 (78,90)
        /// </summary>
        public async Task<bool> CheckSearchResultAsync(IntPtr targetHwnd, string expectedText, bool isWework)
        {
            try
            {
                if (targetHwnd == IntPtr.Zero || !GetWindowRect(targetHwnd, out RECT rect)) return false;

                int relX, relY, width, height;

                if (isWework)
                {
                    // 🏢 企业微信坐标 (用户提供: 78,90 - 472,148)
                    relX = 78;
                    relY = 90;
                    width = 472 - 78;  // 394
                    height = 148 - 90; // 58
                }
                else
                {
                    // 💬 微信坐标 (原有: 53,93 - 277,150)
                    relX = 53;
                    relY = 93;
                    width = 277 - 53;  // 224
                    height = 150 - 93; // 57
                }

                int screenX = rect.Left + relX;
                int screenY = rect.Top + relY;

                // 确保尺寸有效
                if (width <= 0 || height <= 0) return false;

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(screenX, screenY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    // 使用 2 倍放大进行 OCR (搜索列表字体通常较清晰，2倍足够，也可改3倍)
                    using (var scaledMap = ScaleImage(bitmap, 3))
                    {
                        string ocrText = await PerformOcrAsync(scaledMap);
                        System.Diagnostics.Debug.WriteLine($"OCR结果: {ocrText}");
                        return IsFuzzyMatch(expectedText, ocrText);
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"💥 搜索验证出错: {ex.Message}");
                return false;
            }
        }




     


    }
}