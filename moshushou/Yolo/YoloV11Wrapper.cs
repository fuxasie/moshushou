using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;

namespace moshushou.Yolo
{
    public class YoloResult
    {
        public int LabelId { get; set; }
        public string LabelName { get; set; }
        public float Confidence { get; set; }
        public Rectangle BBox { get; set; } // 在原图上的坐标
    }

    public class YoloV11Wrapper : IDisposable
    {
        private readonly InferenceSession _session;
        private readonly string[] _labels;
        private const int ModelInputSize = 640;

        public YoloV11Wrapper(string modelPath)
        {
            if (!File.Exists(modelPath))
                throw new FileNotFoundException("Model file not found", modelPath);

            var options = new SessionOptions();
            // 可以根据需要配置 options，例如开启 CUDA
             try
            {
               // options.AppendExecutionProvider_CPU(0); // 默认 CPU
            }
            catch {}

            _session = new InferenceSession(modelPath, options);
            
            // 按照 YAML 定义的类别顺序
            _labels = new string[] 
            { 
                "在线文档", 
                "搜索群聊", 
                "最近搜索群聊", 
                "群聊名字", 
                "聊天信息", 
                "聊天框" 
            };
        }

        public List<YoloResult> Predict(Bitmap bitmap, float confThreshold = 0.25f, float iouThreshold = 0.45f)
        {
            if (bitmap == null) return new List<YoloResult>();

            // 1. 预处理
            var (tensor, scaleX, scaleY, padX, padY) = Preprocess(bitmap);

            // 2. 推理
            var inputs = new List<NamedOnnxValue>
            {
                NamedOnnxValue.CreateFromTensor("images", tensor)
            };

            using var results = _session.Run(inputs);
            
            // YOLOv8/11 output shape: [1, 4 + num_classes, 8400]
            // 4 boxes (xc, yc, w, h) + 6 classes = 10 channels
            var output = results.First().AsTensor<float>();
            
            // 3. 后处理
            return Postprocess(output, scaleX, scaleY, padX, padY, confThreshold, iouThreshold);
        }

        private (DenseTensor<float> tensor, float scaleX, float scaleY, int padX, int padY) Preprocess(Bitmap image)
        {
            int w = image.Width;
            int h = image.Height;
            
            // Letterbox resize strategy
            float scale = Math.Min((float)ModelInputSize / w, (float)ModelInputSize / h);
            int newW = (int)(w * scale);
            int newH = (int)(h * scale);
            
            // Padding
            int padX = (ModelInputSize - newW) / 2;
            int padY = (ModelInputSize - newH) / 2;
            
            // Create inputs
            var tensor = new DenseTensor<float>(new[] { 1, 3, ModelInputSize, ModelInputSize });

            using (var resized = new Bitmap(ModelInputSize, ModelInputSize, PixelFormat.Format24bppRgb))
            using (var g = Graphics.FromImage(resized))
            {
                g.Clear(Color.FromArgb(114, 114, 114)); // YOLO padding color
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Bilinear;
                g.DrawImage(image, new Rectangle(padX, padY, newW, newH));
                
                // Lock bits for fast access
                BitmapData inputData = resized.LockBits(new Rectangle(0, 0, ModelInputSize, ModelInputSize), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                
                unsafe
                {
                    byte* ptr = (byte*)inputData.Scan0;
                    int stride = inputData.Stride;
                    
                    // Parallel loop for speedup could be used here, but keeping simple for now
                    for (int y = 0; y < ModelInputSize; y++)
                    {
                        byte* row = ptr + (y * stride);
                        for (int x = 0; x < ModelInputSize; x++)
                        {
                            // BGR layout in Bitmap, RGB expected by model
                            // Normalization 0-255 -> 0.0-1.0
                            tensor[0, 0, y, x] = row[x * 3 + 2] / 255.0f; // R
                            tensor[0, 1, y, x] = row[x * 3 + 1] / 255.0f; // G
                            tensor[0, 2, y, x] = row[x * 3 + 0] / 255.0f; // B
                        }
                    }
                }
                
                resized.UnlockBits(inputData);
            }

            return (tensor, 1.0f / scale, 1.0f / scale, padX, padY);
        }

        private List<YoloResult> Postprocess(Tensor<float> output, float scaleX, float scaleY, int padX, int padY, float confThres, float iouThres)
        {
            // Output shape: [1, 10, 8400] -> [Batch, Channels, Anchors]
            // We need to transpose logic conceptually or iterate correctly.
            // Stride is typically: [Channels * Anchors, Anchors, 1] for [Batch, Channels, Anchors]?
            // Tensor impl usually flat.
            
            // Dimensions check
            // Expected: [1, 10, 8400]
            var dimensions = output.Dimensions;
            int channels = dimensions[1]; // 4 box + 6 classes = 10
            int anchors = dimensions[2];  // 8400
            
            var proposals = new List<YoloResult>();

            // Iterate over all anchors
            for (int i = 0; i < anchors; i++)
            {
                // Find class with max confidence
                float maxScore = 0;
                int maxClassId = -1;
                
                // Classes start at index 4
                for (int c = 0; c < _labels.Length; c++)
                {
                    // output[0, 4 + c, i]
                    float score = output[0, 4 + c, i];
                    if (score > maxScore)
                    {
                        maxScore = score;
                        maxClassId = c;
                    }
                }

                if (maxScore < confThres) continue;

                // Extract box (xc, yc, w, h)
                float xc = output[0, 0, i];
                float yc = output[0, 1, i];
                float w = output[0, 2, i];
                float h = output[0, 3, i];

                // Convert to top-left (x, y) relative to input image (before padding removal)
                float x = xc - w / 2;
                float y = yc - h / 2;

                // Remove padding and scale back to original image
                float xOrig = (x - padX) * scaleX;
                float yOrig = (y - padY) * scaleY;
                float wOrig = w * scaleX;
                float hOrig = h * scaleY;

                proposals.Add(new YoloResult
                {
                    LabelId = maxClassId,
                    LabelName = _labels[maxClassId],
                    Confidence = maxScore,
                    BBox = new Rectangle((int)xOrig, (int)yOrig, (int)wOrig, (int)hOrig)
                });
            }

            return NMS(proposals, iouThres);
        }

        private List<YoloResult> NMS(List<YoloResult> boxes, float iouThreshold)
        {
            var result = new List<YoloResult>();
            var sortedBoxes = boxes.OrderByDescending(b => b.Confidence).ToList();

            while (sortedBoxes.Count > 0)
            {
                var current = sortedBoxes[0];
                result.Add(current);
                sortedBoxes.RemoveAt(0);

                for (int i = sortedBoxes.Count - 1; i >= 0; i--)
                {
                    if (CalculateIoU(current.BBox, sortedBoxes[i].BBox) > iouThreshold)
                    {
                        sortedBoxes.RemoveAt(i);
                    }
                }
            }

            return result;
        }

        private float CalculateIoU(Rectangle rect1, Rectangle rect2)
        {
            var intersection = Rectangle.Intersect(rect1, rect2);
            float intersectionArea = intersection.Width * intersection.Height;
            if (intersection.Width <= 0 || intersection.Height <= 0) intersectionArea = 0;

            float unionArea = rect1.Width * rect1.Height + rect2.Width * rect2.Height - intersectionArea;
            
            if (unionArea <= 0) return 0;

            return intersectionArea / unionArea;
        }

        /// <summary>
        /// [新增] 保存带有检测框的调试图片
        /// </summary>
        public void SaveDebugImage(Bitmap original, List<YoloResult> results, string outputPath)
        {
            try
            {
                using (var canvas = new Bitmap(original))
                using (var g = Graphics.FromImage(canvas))
                {
                    foreach (var res in results)
                    {
                        var pen = new Pen(Color.Red, 3);
                        g.DrawRectangle(pen, res.BBox);
                        
                        string info = $"{res.LabelName} {res.Confidence:P0}";
                        var font = new Font("Consolas", 14, FontStyle.Bold);
                        var size = g.MeasureString(info, font);
                        
                        g.FillRectangle(Brushes.Red, res.BBox.X, res.BBox.Y - (int)size.Height, (int)size.Width, (int)size.Height);
                        g.DrawString(info, font, Brushes.White, res.BBox.X, res.BBox.Y - (int)size.Height);
                    }
                    canvas.Save(outputPath, ImageFormat.Png);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving debug image: {ex.Message}");
            }
        }

        public void Dispose()
        {
            _session?.Dispose();
        }
    }
}
