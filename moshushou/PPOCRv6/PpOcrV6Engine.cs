using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;

namespace moshushou.Ocr
{
    public sealed class PpOcrTextLine
    {
        public string Text { get; init; } = string.Empty;
        public Rectangle Bounds { get; init; }
        public float Confidence { get; init; }
    }

    /// <summary>
    /// PP-OCRv6 medium local OCR pipeline.
    /// Detection and recognition models are the official PaddlePaddle ONNX
    /// exports of the PP-OCRv6 medium safetensors model family.
    /// </summary>
    public sealed class PpOcrV6Engine : IDisposable
    {
        public const string DisplayName = "PP-OCRv6 medium";
        public const string DetectionModelId = "PaddlePaddle/PP-OCRv6_medium_det_safetensors";
        public const string RecognitionModelId = "PaddlePaddle/PP-OCRv6_medium_rec_safetensors";

        private const string DetectionModelFileName = "PP-OCRv6_medium_det.onnx";
        private const string RecognitionModelFileName = "PP-OCRv6_medium_rec.onnx";
        private const string RecognitionProcessorFileName = "PP-OCRv6_medium_rec.preprocessor.json";

        private const float DetectionPixelThreshold = 0.20f;
        private const float DetectionBoxThreshold = 0.45f;
        private const float DetectionUnclipRatio = 1.40f;
        private const int DetectionLimitSideLength = 736;
        private const int DetectionMaxSideLength = 4000;
        private const int RecognitionHeight = 48;
        private const int RecognitionDefaultWidth = 320;
        private const int RecognitionMaxWidth = 3200;
        private const int MaxTextLines = 256;

        private readonly string _modelDirectory;
        private readonly SemaphoreSlim _inferenceGate = new(1, 1);
        private readonly object _initializationLock = new();

        private InferenceSession? _detector;
        private InferenceSession? _recognizer;
        private string _detectorInputName = "x";
        private string _recognizerInputName = "x";
        private IReadOnlyList<string> _characters = Array.Empty<string>();
        private int _recognizerFixedWidth;
        private bool _disposed;

        public PpOcrV6Engine(string? modelDirectory = null)
        {
            _modelDirectory = string.IsNullOrWhiteSpace(modelDirectory)
                ? GetDefaultModelDirectory()
                : Path.GetFullPath(modelDirectory);
        }

        public static string GetDefaultModelDirectory()
        {
            return Path.Combine(AppContext.BaseDirectory, "PPOCRv6", "models");
        }

        public string ModelDirectory => _modelDirectory;

        public bool ModelFilesAvailable =>
            File.Exists(Path.Combine(_modelDirectory, DetectionModelFileName)) &&
            File.Exists(Path.Combine(_modelDirectory, RecognitionModelFileName)) &&
            File.Exists(Path.Combine(_modelDirectory, RecognitionProcessorFileName));

        public Task WarmUpAsync(CancellationToken token = default)
        {
            return Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();
                EnsureInitialized();
            }, token);
        }

        public async Task<string> RecognizeAsync(Bitmap bitmap, CancellationToken token = default)
        {
            IReadOnlyList<PpOcrTextLine> lines = await RecognizeDetailedAsync(bitmap, token).ConfigureAwait(false);
            if (lines.Count == 0)
            {
                return string.Empty;
            }

            var text = new StringBuilder();
            foreach (PpOcrTextLine line in lines)
            {
                text.Append(line.Text);
            }
            return text.ToString().Trim();
        }

        public async Task<IReadOnlyList<PpOcrTextLine>> RecognizeDetailedAsync(
            Bitmap bitmap,
            CancellationToken token = default)
        {
            ArgumentNullException.ThrowIfNull(bitmap);
            ThrowIfDisposed();

            await _inferenceGate.WaitAsync(token).ConfigureAwait(false);
            try
            {
                token.ThrowIfCancellationRequested();
                return await Task.Run(() => RecognizeCore(bitmap, token), token).ConfigureAwait(false);
            }
            finally
            {
                _inferenceGate.Release();
            }
        }

        private IReadOnlyList<PpOcrTextLine> RecognizeCore(Bitmap bitmap, CancellationToken token)
        {
            EnsureInitialized();
            token.ThrowIfCancellationRequested();

            List<Rectangle> boxes = DetectTextBoxes(bitmap, token);
            if (boxes.Count == 0)
            {
                return Array.Empty<PpOcrTextLine>();
            }

            var result = new List<PpOcrTextLine>(boxes.Count);
            foreach (Rectangle box in boxes.Take(MaxTextLines))
            {
                token.ThrowIfCancellationRequested();
                using Bitmap crop = CloneAs24Bpp(bitmap, box);
                (string text, float confidence) = RecognizeLine(crop);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                result.Add(new PpOcrTextLine
                {
                    Text = text.Trim(),
                    Bounds = box,
                    Confidence = confidence
                });
            }

            return result;
        }

        private void EnsureInitialized()
        {
            ThrowIfDisposed();
            if (_detector != null && _recognizer != null && _characters.Count > 0)
            {
                return;
            }

            lock (_initializationLock)
            {
                if (_detector != null && _recognizer != null && _characters.Count > 0)
                {
                    return;
                }

                string detPath = Path.Combine(_modelDirectory, DetectionModelFileName);
                string recPath = Path.Combine(_modelDirectory, RecognitionModelFileName);
                string processorPath = Path.Combine(_modelDirectory, RecognitionProcessorFileName);
                var missingFiles = new[] { detPath, recPath, processorPath }
                    .Where(path => !File.Exists(path))
                    .Select(Path.GetFileName)
                    .ToArray();

                if (missingFiles.Length > 0)
                {
                    throw new FileNotFoundException(
                        $"PP-OCRv6 模型不完整，缺少: {string.Join(", ", missingFiles)}。模型目录: {_modelDirectory}");
                }

                var options = new SessionOptions
                {
                    GraphOptimizationLevel = GraphOptimizationLevel.ORT_ENABLE_ALL,
                    ExecutionMode = ExecutionMode.ORT_SEQUENTIAL,
                    InterOpNumThreads = 1,
                    IntraOpNumThreads = Math.Clamp(Environment.ProcessorCount / 2, 1, 4)
                };

                InferenceSession? detector = null;
                InferenceSession? recognizer = null;
                try
                {
                    detector = new InferenceSession(detPath, options);
                    recognizer = new InferenceSession(recPath, options);
                    IReadOnlyList<string> characters = LoadCharacterDictionary(processorPath);

                    _detectorInputName = detector.InputMetadata.Keys.First();
                    _recognizerInputName = recognizer.InputMetadata.Keys.First();
                    NodeMetadata recInput = recognizer.InputMetadata[_recognizerInputName];
                    _recognizerFixedWidth = recInput.Dimensions.Length >= 4 && recInput.Dimensions[3] > 0
                        ? recInput.Dimensions[3]
                        : 0;

                    _characters = characters;
                    _detector = detector;
                    _recognizer = recognizer;
                    detector = null;
                    recognizer = null;

                    Debug.WriteLine(
                        $"[PP-OCRv6] Loaded. DetInput={_detectorInputName}, RecInput={_recognizerInputName}, " +
                        $"RecFixedWidth={_recognizerFixedWidth}, Characters={_characters.Count}");
                }
                finally
                {
                    detector?.Dispose();
                    recognizer?.Dispose();
                    options.Dispose();
                }
            }
        }

        private static IReadOnlyList<string> LoadCharacterDictionary(string processorPath)
        {
            using FileStream stream = File.OpenRead(processorPath);
            using JsonDocument document = JsonDocument.Parse(stream);
            if (!document.RootElement.TryGetProperty("character_list", out JsonElement listElement) ||
                listElement.ValueKind != JsonValueKind.Array)
            {
                throw new InvalidDataException("PP-OCRv6 识别模型缺少 character_list。");
            }

            var characters = new List<string>(listElement.GetArrayLength());
            foreach (JsonElement item in listElement.EnumerateArray())
            {
                characters.Add(item.GetString() ?? string.Empty);
            }

            if (characters.Count < 2 || !string.Equals(characters[0], "blank", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("PP-OCRv6 识别字符表格式无效。");
            }

            return characters;
        }

        private List<Rectangle> DetectTextBoxes(Bitmap bitmap, CancellationToken token)
        {
            (DenseTensor<float> input, int inputWidth, int inputHeight) = CreateDetectionInput(bitmap);
            NamedOnnxValue inputValue = NamedOnnxValue.CreateFromTensor(_detectorInputName, input);
            using IDisposableReadOnlyCollection<DisposableNamedOnnxValue> outputs =
                _detector!.Run(new[] { inputValue });

            Tensor<float> probabilityMap = outputs.First().AsTensor<float>();
            int[] dimensions = probabilityMap.Dimensions.ToArray();
            if (dimensions.Length < 2)
            {
                throw new InvalidDataException(
                    $"PP-OCRv6 检测输出维度无效: [{string.Join(",", dimensions)}]");
            }

            int mapHeight = dimensions[^2];
            int mapWidth = dimensions[^1];
            if (mapWidth <= 0 || mapHeight <= 0)
            {
                return new List<Rectangle>();
            }

            float[] probabilities = probabilityMap.ToArray();
            int mapOffset = Math.Max(0, probabilities.Length - (mapWidth * mapHeight));
            List<DetectionComponent> components = ExtractDetectionComponents(
                probabilities,
                mapOffset,
                mapWidth,
                mapHeight,
                token);

            float mapToOriginalX = bitmap.Width / (float)mapWidth;
            float mapToOriginalY = bitmap.Height / (float)mapHeight;
            var boxes = new List<Rectangle>(components.Count);

            foreach (DetectionComponent component in components)
            {
                float averageScore = component.SumScore / Math.Max(1, component.PixelCount);
                if (averageScore < DetectionBoxThreshold)
                {
                    continue;
                }

                int componentWidth = component.MaxX - component.MinX + 1;
                int componentHeight = component.MaxY - component.MinY + 1;
                if (componentWidth < 2 || componentHeight < 2)
                {
                    continue;
                }

                float expandX = componentWidth * (DetectionUnclipRatio - 1f) / 2f;
                float expandY = componentHeight * (DetectionUnclipRatio - 1f) / 2f;
                int left = (int)Math.Floor((component.MinX - expandX) * mapToOriginalX);
                int top = (int)Math.Floor((component.MinY - expandY) * mapToOriginalY);
                int right = (int)Math.Ceiling((component.MaxX + 1 + expandX) * mapToOriginalX);
                int bottom = (int)Math.Ceiling((component.MaxY + 1 + expandY) * mapToOriginalY);

                Rectangle box = ClampRectangle(
                    Rectangle.FromLTRB(left, top, right, bottom),
                    bitmap.Width,
                    bitmap.Height);
                if (box.Width >= 3 && box.Height >= 3)
                {
                    boxes.Add(box);
                }
            }

            return MergeAndSortBoxes(boxes, bitmap.Width, bitmap.Height);
        }

        private static List<DetectionComponent> ExtractDetectionComponents(
            float[] probabilities,
            int offset,
            int width,
            int height,
            CancellationToken token)
        {
            int pixelCount = checked(width * height);
            var visited = new bool[pixelCount];
            var queue = new int[pixelCount];
            var components = new List<DetectionComponent>();

            for (int start = 0; start < pixelCount; start++)
            {
                if ((start & 0x3FFF) == 0)
                {
                    token.ThrowIfCancellationRequested();
                }

                if (visited[start] || probabilities[offset + start] < DetectionPixelThreshold)
                {
                    continue;
                }

                int head = 0;
                int tail = 0;
                queue[tail++] = start;
                visited[start] = true;

                int minX = width;
                int minY = height;
                int maxX = 0;
                int maxY = 0;
                int count = 0;
                float scoreSum = 0;

                while (head < tail)
                {
                    int index = queue[head++];
                    int y = index / width;
                    int x = index - (y * width);
                    float score = probabilities[offset + index];

                    minX = Math.Min(minX, x);
                    minY = Math.Min(minY, y);
                    maxX = Math.Max(maxX, x);
                    maxY = Math.Max(maxY, y);
                    count++;
                    scoreSum += score;

                    int y0 = Math.Max(0, y - 1);
                    int y1 = Math.Min(height - 1, y + 1);
                    int x0 = Math.Max(0, x - 1);
                    int x1 = Math.Min(width - 1, x + 1);
                    for (int neighborY = y0; neighborY <= y1; neighborY++)
                    {
                        int rowOffset = neighborY * width;
                        for (int neighborX = x0; neighborX <= x1; neighborX++)
                        {
                            int neighborIndex = rowOffset + neighborX;
                            if (visited[neighborIndex] ||
                                probabilities[offset + neighborIndex] < DetectionPixelThreshold)
                            {
                                continue;
                            }

                            visited[neighborIndex] = true;
                            queue[tail++] = neighborIndex;
                        }
                    }
                }

                if (count >= 4)
                {
                    components.Add(new DetectionComponent(
                        minX,
                        minY,
                        maxX,
                        maxY,
                        count,
                        scoreSum));
                }
            }

            return components;
        }

        private static List<Rectangle> MergeAndSortBoxes(
            List<Rectangle> boxes,
            int imageWidth,
            int imageHeight)
        {
            if (boxes.Count <= 1)
            {
                return boxes;
            }

            var pending = boxes
                .OrderBy(box => box.Top)
                .ThenBy(box => box.Left)
                .ToList();

            bool changed;
            do
            {
                changed = false;
                for (int i = 0; i < pending.Count && !changed; i++)
                {
                    for (int j = i + 1; j < pending.Count; j++)
                    {
                        Rectangle a = pending[i];
                        Rectangle b = pending[j];
                        float verticalOverlap = Math.Max(
                            0,
                            Math.Min(a.Bottom, b.Bottom) - Math.Max(a.Top, b.Top));
                        float overlapRatio = verticalOverlap / Math.Max(1f, Math.Min(a.Height, b.Height));
                        int horizontalGap = Math.Max(0, Math.Max(a.Left, b.Left) - Math.Min(a.Right, b.Right));
                        int mergeGap = Math.Max(6, (int)(Math.Max(a.Height, b.Height) * 1.25f));

                        if (overlapRatio >= 0.45f && horizontalGap <= mergeGap)
                        {
                            pending[i] = ClampRectangle(
                                Rectangle.Union(a, b),
                                imageWidth,
                                imageHeight);
                            pending.RemoveAt(j);
                            changed = true;
                            break;
                        }
                    }
                }
            }
            while (changed);

            float rowTolerance = pending.Count == 0
                ? 4
                : Math.Max(4, pending.Select(box => box.Height).OrderBy(value => value).ElementAt(pending.Count / 2) * 0.6f);

            return pending
                .OrderBy(box => (int)Math.Round((box.Top + box.Height / 2f) / rowTolerance))
                .ThenBy(box => box.Left)
                .Take(MaxTextLines)
                .ToList();
        }

        private (string text, float confidence) RecognizeLine(Bitmap line)
        {
            (DenseTensor<float> input, int contentWidth) = CreateRecognitionInput(line);
            NamedOnnxValue inputValue = NamedOnnxValue.CreateFromTensor(_recognizerInputName, input);
            using IDisposableReadOnlyCollection<DisposableNamedOnnxValue> outputs =
                _recognizer!.Run(new[] { inputValue });

            Tensor<float> logits = outputs.First().AsTensor<float>();
            int[] dimensions = logits.Dimensions.ToArray();
            if (dimensions.Length != 3)
            {
                throw new InvalidDataException(
                    $"PP-OCRv6 识别输出维度无效: [{string.Join(",", dimensions)}], ContentWidth={contentWidth}");
            }

            int timeSteps = dimensions[1];
            int classCount = dimensions[2];
            if (classCount != _characters.Count)
            {
                throw new InvalidDataException(
                    $"PP-OCRv6 字符表不匹配: Model={classCount}, Dictionary={_characters.Count}");
            }

            var text = new StringBuilder();
            int previousIndex = -1;
            float confidenceSum = 0;
            int confidenceCount = 0;

            for (int step = 0; step < timeSteps; step++)
            {
                int bestIndex = 0;
                float bestLogit = logits[0, step, 0];
                for (int classIndex = 1; classIndex < classCount; classIndex++)
                {
                    float value = logits[0, step, classIndex];
                    if (value > bestLogit)
                    {
                        bestLogit = value;
                        bestIndex = classIndex;
                    }
                }

                if (bestIndex != 0 && bestIndex != previousIndex)
                {
                    text.Append(_characters[bestIndex]);
                    confidenceSum += ComputeSoftmaxConfidence(logits, step, classCount, bestLogit);
                    confidenceCount++;
                }
                previousIndex = bestIndex;
            }

            float confidence = confidenceCount == 0 ? 0 : confidenceSum / confidenceCount;
            return (text.ToString(), confidence);
        }

        private static float ComputeSoftmaxConfidence(
            Tensor<float> logits,
            int step,
            int classCount,
            float maximum)
        {
            double denominator = 0;
            for (int classIndex = 0; classIndex < classCount; classIndex++)
            {
                denominator += Math.Exp(logits[0, step, classIndex] - maximum);
            }

            return denominator <= 0 ? 0 : (float)(1d / denominator);
        }

        private static (DenseTensor<float> tensor, int width, int height) CreateDetectionInput(Bitmap image)
        {
            (int width, int height) = CalculateDetectionSize(image.Width, image.Height);
            var tensor = new DenseTensor<float>(new[] { 1, 3, height, width });

            using var resized = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (Graphics graphics = Graphics.FromImage(resized))
            {
                graphics.Clear(Color.White);
                graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphics.DrawImage(image, new Rectangle(0, 0, width, height));
            }

            BitmapData data = resized.LockBits(
                new Rectangle(0, 0, width, height),
                ImageLockMode.ReadOnly,
                PixelFormat.Format24bppRgb);
            try
            {
                int stride = data.Stride;
                byte[] pixels = new byte[Math.Abs(stride) * height];
                Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
                for (int y = 0; y < height; y++)
                {
                    int row = y * stride;
                    for (int x = 0; x < width; x++)
                    {
                        int pixel = row + (x * 3);
                        tensor[0, 0, y, x] = ((pixels[pixel] / 255f) - 0.485f) / 0.229f;
                        tensor[0, 1, y, x] = ((pixels[pixel + 1] / 255f) - 0.456f) / 0.224f;
                        tensor[0, 2, y, x] = ((pixels[pixel + 2] / 255f) - 0.406f) / 0.225f;
                    }
                }
            }
            finally
            {
                resized.UnlockBits(data);
            }

            return (tensor, width, height);
        }

        private (DenseTensor<float> tensor, int contentWidth) CreateRecognitionInput(Bitmap image)
        {
            float ratio = image.Width / (float)Math.Max(1, image.Height);
            int naturalWidth = Math.Clamp(
                (int)Math.Ceiling(RecognitionHeight * ratio),
                1,
                RecognitionMaxWidth);
            int inputWidth = _recognizerFixedWidth > 0
                ? _recognizerFixedWidth
                : RoundUpToMultiple(Math.Max(32, naturalWidth), 32);
            inputWidth = Math.Clamp(inputWidth, 32, RecognitionMaxWidth);
            int resizedWidth = Math.Min(inputWidth, naturalWidth);

            var tensor = new DenseTensor<float>(new[] { 1, 3, RecognitionHeight, inputWidth });
            using var resized = new Bitmap(resizedWidth, RecognitionHeight, PixelFormat.Format24bppRgb);
            using (Graphics graphics = Graphics.FromImage(resized))
            {
                graphics.Clear(Color.White);
                graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphics.DrawImage(image, new Rectangle(0, 0, resizedWidth, RecognitionHeight));
            }

            BitmapData data = resized.LockBits(
                new Rectangle(0, 0, resizedWidth, RecognitionHeight),
                ImageLockMode.ReadOnly,
                PixelFormat.Format24bppRgb);
            try
            {
                int stride = data.Stride;
                byte[] pixels = new byte[Math.Abs(stride) * RecognitionHeight];
                Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
                for (int y = 0; y < RecognitionHeight; y++)
                {
                    int row = y * stride;
                    for (int x = 0; x < resizedWidth; x++)
                    {
                        int pixel = row + (x * 3);
                        tensor[0, 0, y, x] = (pixels[pixel] / 255f - 0.5f) / 0.5f;
                        tensor[0, 1, y, x] = (pixels[pixel + 1] / 255f - 0.5f) / 0.5f;
                        tensor[0, 2, y, x] = (pixels[pixel + 2] / 255f - 0.5f) / 0.5f;
                    }
                }
            }
            finally
            {
                resized.UnlockBits(data);
            }

            return (tensor, resizedWidth);
        }

        private static (int width, int height) CalculateDetectionSize(int sourceWidth, int sourceHeight)
        {
            int safeWidth = Math.Max(1, sourceWidth);
            int safeHeight = Math.Max(1, sourceHeight);
            float scale = 1f;
            int shortSide = Math.Min(safeWidth, safeHeight);
            int longSide = Math.Max(safeWidth, safeHeight);

            if (shortSide < DetectionLimitSideLength)
            {
                scale = DetectionLimitSideLength / (float)shortSide;
            }

            if (longSide * scale > DetectionMaxSideLength)
            {
                scale = DetectionMaxSideLength / (float)longSide;
            }

            int width = Math.Max(32, RoundToMultiple((int)Math.Round(safeWidth * scale), 32));
            int height = Math.Max(32, RoundToMultiple((int)Math.Round(safeHeight * scale), 32));
            return (width, height);
        }

        private static int RoundToMultiple(int value, int multiple)
        {
            return Math.Max(multiple, (int)Math.Round(value / (double)multiple) * multiple);
        }

        private static int RoundUpToMultiple(int value, int multiple)
        {
            return Math.Max(multiple, ((value + multiple - 1) / multiple) * multiple);
        }

        private static Bitmap CloneAs24Bpp(Bitmap source, Rectangle bounds)
        {
            Rectangle safeBounds = ClampRectangle(bounds, source.Width, source.Height);
            var result = new Bitmap(safeBounds.Width, safeBounds.Height, PixelFormat.Format24bppRgb);
            using Graphics graphics = Graphics.FromImage(result);
            graphics.Clear(Color.White);
            graphics.DrawImage(
                source,
                new Rectangle(0, 0, result.Width, result.Height),
                safeBounds,
                GraphicsUnit.Pixel);
            return result;
        }

        private static Rectangle ClampRectangle(Rectangle rectangle, int width, int height)
        {
            int left = Math.Clamp(rectangle.Left, 0, Math.Max(0, width - 1));
            int top = Math.Clamp(rectangle.Top, 0, Math.Max(0, height - 1));
            int right = Math.Clamp(rectangle.Right, left + 1, width);
            int bottom = Math.Clamp(rectangle.Bottom, top + 1, height);
            return Rectangle.FromLTRB(left, top, right, bottom);
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            lock (_initializationLock)
            {
                _detector?.Dispose();
                _recognizer?.Dispose();
                _detector = null;
                _recognizer = null;
                _characters = Array.Empty<string>();
            }
            _inferenceGate.Dispose();
        }

        private readonly record struct DetectionComponent(
            int MinX,
            int MinY,
            int MaxX,
            int MaxY,
            int PixelCount,
            float SumScore);
    }
}
