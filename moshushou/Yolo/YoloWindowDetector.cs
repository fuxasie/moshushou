using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace moshushou.Yolo
{
    public class YoloWindowDetector : IDisposable
    {
        private readonly YoloV11Wrapper _wrapper;
        public YoloV11Wrapper InferenceWrapper => _wrapper;
        // 标签常量
        public const string Label_OnlineDoc = "在线文档";
        public const string Label_SearchGroup = "搜索群聊";
        public const string Label_RecentGroup = "最近搜索群聊";
        public const string Label_GroupName = "群聊名字";
        public const string Label_ChatInfo = "聊天信息";
        public const string Label_ChatBox = "聊天框";

        public YoloWindowDetector(string modelPath = null)
        {
            if (string.IsNullOrEmpty(modelPath))
            {
                // 默认路径：当前运行目录下的 wxonnx 文件夹
                modelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wxonnx", "yolo11n_wxscreen_fixed.onnx");
            }
            
            _wrapper = new YoloV11Wrapper(modelPath);
        }

        public List<YoloResult> Detect(Bitmap bitmap)
        {
            var results = _wrapper.Predict(bitmap);
            
            // 业务逻辑处理：合并
            // 要求：最近搜索群聊 和 搜索群聊 进行合并
            // 实际上如果只是为了点击，它们都是"目标"，可以直接视为同一类，或者上层逻辑都接受
            // 这里我们不做修改，直接返回，由上层判断 label 名称
            
            return results;
        }

        public List<YoloResult> Detect(Bitmap bitmap, float confThreshold, float iouThreshold = 0.45f)
        {
            var results = _wrapper.Predict(bitmap, confThreshold, iouThreshold);
            return results;
        }

        /// <summary>
        /// 查找"在线文档"弹窗中心点 (用于点击 "使用原文件")
        /// </summary>
        public Point? FindOnlineDocPopup(Bitmap bitmap)
        {
            var rect = FindOnlineDocPopupBBox(bitmap);
            if (rect.HasValue) return GetCenter(rect.Value);
            return null;
        }

        public Rectangle? FindOnlineDocPopupBBox(Bitmap bitmap)
        {
            var results = _wrapper.Predict(bitmap);
            var target = results
                .Where(r => r.LabelName == Label_OnlineDoc)
                .OrderByDescending(r => r.Confidence)
                .FirstOrDefault();

            return target?.BBox;
        }

        /// <summary>
        /// 查找群聊名字 (用于确认进入了群聊)
        /// </summary>
        public Point? FindGroupChatName(Bitmap bitmap)
        {
            var results = _wrapper.Predict(bitmap);
            var target = results
                .Where(r => r.LabelName == Label_GroupName)
                .OrderByDescending(r => r.Confidence)
                .FirstOrDefault();

            if (target != null)
            {
                return GetCenter(target.BBox);
            }
            return null;
        }

        /// <summary>
        /// 检查是否有搜索结果 (搜索群聊 或 最近搜索群聊)
        /// 返回第一个匹配项的中心点，优先 "搜索群聊"
        /// </summary>
        public Point? FindSearchResult(Bitmap bitmap)
        {
            var results = _wrapper.Predict(bitmap);
            
            // 优先找 "搜索群聊"
            var target = results
                .Where(r => r.LabelName == Label_SearchGroup)
                .OrderByDescending(r => r.Confidence)
                .FirstOrDefault();

            if (target != null) return GetCenter(target.BBox);

            // 其次找 "最近搜索群聊"
            target = results
                .Where(r => r.LabelName == Label_RecentGroup)
                .OrderByDescending(r => r.Confidence)
                .FirstOrDefault();

            if (target != null) return GetCenter(target.BBox);

            return null;
        }

        private Point GetCenter(Rectangle rect)
        {
            return new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
        }

        public void Dispose()
        {
            _wrapper?.Dispose();
        }
    }
}
